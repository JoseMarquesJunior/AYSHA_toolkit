"""Camada LLM: cliente Anthropic, prompts, cache, orçamento de contexto e log.

Uso típico:

    cfg = LLMConfig.from_env()
    ctx = ProjectContext(ctrl)                 # ctrl = l5x_core.parse_file(...)
    llm = LLMClient(cfg)
    result = llm.explain(ctx, "MainProgram", "MainRoutine")
    result = llm.review(ctx)                   # projeto inteiro (map-reduce)
    result = llm.document(ctx, program="MainProgram")
    result = llm.aoi(ctx, "P_Alarm")                 # análise + documentação + roteiro de debug de uma AOI
    result = llm.chat(ctx, history, "o que liga o motor M1?")

Nada do XML cru chega ao modelo: só o índice produzido por l5x_core.serialize.
O log em CSV registra apenas métricas (modo, modelo, tokens, latência), nunca
conteúdo do programa.

IDs de modelo não ficam no código: MODEL_FAST / MODEL_DEEP vêm do ambiente
(.env), de secrets do Streamlit ou de variáveis de ambiente. Os IDs atuais
estão em https://docs.claude.com/en/docs/about-claude/models
"""

from __future__ import annotations

import csv
import os
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

from l5x_core.analysis import analyze
from l5x_core.aoi import AoiReport, aoi_markdown, aoi_summary_rows, build_aoi_report
from l5x_core.model import Controller, Finding, Metrics, SEVERITIES
from l5x_core.rung_text import TIMER_COUNTER
from l5x_core.serialize import Serialized, aoi_interface_text, estimate_tokens, project_overview, routine_text
from l5x_core.xref import AOI_PREFIX, CrossRef, build_xref

from app import retrieval

PROMPTS_DIR = Path(__file__).parent / "prompts"

LANGUAGES: Dict[str, str] = {
    "pt-BR": "português do Brasil",
    "en": "English",
    "es": "español",
    "zh-CN": "中文（简体）",
    "hi": "हिन्दी (Hindi)",
}

MODES = ("explain", "review", "document", "chat", "aoi")
DEEP_MODES = {"review", "document", "aoi"}  # sempre no MODEL_DEEP; explain/chat usam MODEL_FAST, ou MODEL_DEEP com deep=True

# Preço por milhão de tokens (entrada, saída, leitura de cache), tabela da Anthropic de out/2026.
# Sobrescreva com MODEL_PRICES="id:in:out:cache,...". Modelo desconhecido assume o mais caro.
PRICES_USD: Dict[str, Tuple[float, float, float]] = {
    "claude-sonnet-5-5": (2.0, 10.0, 0.20), "claude-sonnet-5": (2.0, 10.0, 0.20), "claude-sonnet-4-6": (3.0, 15.0, 0.30),
    "claude-opus-5-5": (4.0, 20.0, 0.20), "claude-opus-5": (5.0, 25.0, 0.50), "claude-opus-4-8": (5.0, 25.0, 0.50),
    "claude-fable-5-1": (10.0, 50.0, 0.25), "claude-fable-5": (10.0, 50.0, 0.25),
    "claude-haiku-5-5": (0.10, 0.50, 0.01), "claude-haiku-4-5": (1.0, 5.0, 0.10),
}
DEFAULT_PRICE = (10.0, 50.0, 1.0)
# Teto de tokens de saída por modo (raciocínio é cobrado como saída) e saída típica medida no uso real.
OUTPUT_CAPS: Dict[str, int] = {"explain": 6_000, "chat": 4_000, "review": 16_000, "document": 16_000, "aoi": 12_000}
EXPECTED_OUTPUT: Dict[str, int] = {"explain": 2_500, "chat": 3_500, "review": 12_000, "document": 10_000, "aoi": 10_000}

# ----------------------------------------------------------------------------
# Errors
# ----------------------------------------------------------------------------


class LLMError(Exception):
    """Erro com mensagem pronta para o usuário (português)."""

    def __init__(self, message: str, kind: str = "api", cause: Optional[BaseException] = None):
        super().__init__(message)
        self.message = message
        self.kind = kind  # config | auth | rate_limit | context | connection | api | model
        self.cause = cause


class ConfigError(LLMError):
    def __init__(self, message: str):
        super().__init__(message, kind="config")


# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------


@dataclass
class LLMConfig:
    model_fast: str
    model_deep: str
    api_key: Optional[str] = None
    effort_fast: Optional[str] = None  # low | medium | high | xhigh | max (opcional)
    effort_deep: Optional[str] = None
    max_context_tokens: int = 60_000
    max_output_tokens: int = 64_000  # streaming: o SDK aceita até 128k sem timeout
    max_continuations: int = 2  # pedidos de "continue" quando a resposta para por max_tokens
    log_path: Optional[str] = "llm_calls.csv"
    language: str = "pt-BR"
    credit_usd: float = 0.10  # valor de 1 crédito em custo de API
    markup: float = 4.0  # margem sobre o custo de API ao converter em créditos
    output_caps: Dict[str, int] = field(default_factory=lambda: dict(OUTPUT_CAPS))
    prices: Dict[str, Tuple[float, float, float]] = field(default_factory=lambda: dict(PRICES_USD))

    @classmethod
    def from_env(cls, secrets: Optional[Mapping[str, Any]] = None, dotenv: bool = True) -> "LLMConfig":
        """Lê .env (python-dotenv), variáveis de ambiente e, por último, ``secrets``
        (ex.: ``st.secrets``). Lança ConfigError se faltar MODEL_FAST/MODEL_DEEP."""
        if dotenv:
            try:
                from dotenv import load_dotenv

                load_dotenv()
            except ImportError:  # pragma: no cover
                pass

        def get(name: str, default: Optional[str] = None) -> Optional[str]:
            value: Optional[str] = None
            if secrets is not None:
                try:
                    value = secrets.get(name)  # type: ignore[union-attr]
                except Exception:  # pragma: no cover - secrets podem lançar se não existir arquivo
                    value = None
            if value in (None, ""):
                value = os.environ.get(name)
            value = str(value).strip() if value not in (None, "") else None
            return value if value else default

        model_fast = get("MODEL_FAST")
        model_deep = get("MODEL_DEEP") or model_fast
        if not model_fast:
            raise ConfigError(
                "MODEL_FAST não configurado. Defina MODEL_FAST e MODEL_DEEP no .env ou em .streamlit/secrets.toml "
                "(IDs atuais em https://docs.claude.com/en/docs/about-claude/models)."
            )
        language = get("APP_LANGUAGE", "pt-BR") or "pt-BR"
        if language not in LANGUAGES:
            language = "pt-BR"
        caps = dict(OUTPUT_CAPS)
        for mode in MODES:
            raw = get(f"OUTPUT_TOKENS_{mode.upper()}")
            if raw and raw.isdigit():
                caps[mode] = int(raw)
        prices = dict(PRICES_USD)
        for item in (get("MODEL_PRICES") or "").split(","):
            parts = item.strip().split(":")
            if len(parts) == 4:
                try:
                    prices[parts[0]] = (float(parts[1]), float(parts[2]), float(parts[3]))
                except ValueError:
                    pass
        return cls(
            model_fast=model_fast,
            model_deep=model_deep or model_fast,
            api_key=get("ANTHROPIC_API_KEY"),
            effort_fast=get("LLM_EFFORT_FAST"),
            effort_deep=get("LLM_EFFORT_DEEP"),
            max_context_tokens=int(get("MAX_CONTEXT_TOKENS", "60000") or 60000),
            max_output_tokens=int(get("MAX_OUTPUT_TOKENS", "64000") or 64000),
            max_continuations=int(get("LLM_MAX_CONTINUATIONS", "2") or 0),
            log_path=get("LLM_LOG_PATH", "llm_calls.csv"),
            language=language,
            credit_usd=float(get("CREDIT_USD", "0.10") or 0.10),
            markup=float(get("CREDIT_MARKUP", "4") or 4),
            output_caps=caps,
            prices=prices,
        )

    def model_for(self, mode: str, deep: bool = False) -> str:
        return self.model_deep if (mode in DEEP_MODES or deep) else self.model_fast

    def effort_for(self, mode: str, deep: bool = False) -> Optional[str]:
        return self.effort_deep if (mode in DEEP_MODES or deep) else self.effort_fast

    def output_cap(self, mode: str) -> int:
        return min(self.max_output_tokens, self.output_caps.get(mode, self.max_output_tokens))

    def price_for(self, model: str) -> Tuple[float, float, float]:
        return self.prices.get(model, DEFAULT_PRICE)

    @property
    def has_deep_option(self) -> bool:
        return self.model_deep != self.model_fast


# ----------------------------------------------------------------------------
# Project context (built once per uploaded file / session)
# ----------------------------------------------------------------------------


class ProjectContext:
    """Índice do projeto pronto para o modelo: overview (cacheado), textos de rotina
    sob demanda, findings por escopo, tabelas de I/O e timers."""

    def __init__(
        self,
        ctrl: Controller,
        xref: Optional[CrossRef] = None,
        findings: Optional[List[Finding]] = None,
        metrics: Optional[Metrics] = None,
        overview_tokens: int = 6000,
    ) -> None:
        self.ctrl = ctrl
        self.xref = xref or build_xref(ctrl)
        if findings is None or metrics is None:
            findings, metrics = analyze(ctrl, self.xref)
        self.findings = findings
        self.metrics = metrics
        self.overview: Serialized = project_overview(ctrl, findings, metrics, max_tokens=overview_tokens)
        self._routine_cache: Dict[Tuple[str, str], Serialized] = {}
        self._word_index: Optional[Dict[Tuple[str, str], Set[str]]] = None
        self._desc_index: Dict[Tuple[str, str], Set[str]] = {}
        self._aoi_reports: Dict[str, AoiReport] = {}

    # ---- routines ----------------------------------------------------------
    def routine_text(self, program: str, routine: str) -> Serialized:
        """``program`` = nome do programa ou ``aoi:<Nome>`` (rotina interna de AOI)."""
        key = (program.lower(), routine.lower())
        sr = self._routine_cache.get(key)
        if sr is None:
            sr = routine_text(self.ctrl, program, routine, self.xref)
            self._routine_cache[key] = sr
        return sr

    # ---- FBD: elements without wires ---------------------------------------
    def unconnected_rows(self) -> List[Dict[str, Any]]:
        """Uma linha por elemento FBD sem fio (tag solta na folha), em programas e AOIs."""
        rows: List[Dict[str, Any]] = []
        scopes = [(p.name, p.routines) for p in self.ctrl.programs] + [(f"{AOI_PREFIX}{a.name}", a.routines) for a in self.ctrl.aois]
        for scope, routines in scopes:
            for r in routines:
                for d in r.dangling:
                    rows.append({"programa": scope, "rotina": r.name, "folha": d.sheet, "tipo": d.kind, "elemento": d.operand, "tag": d.is_tag})
        return rows

    # ---- AOIs ---------------------------------------------------------------
    def aoi_names(self) -> List[str]:
        return [a.name for a in self.ctrl.aois]

    def aoi_report(self, name: str) -> AoiReport:
        """Relatório determinístico da AOI (cacheado); KeyError se não existir."""
        key = name.lower()
        rep = self._aoi_reports.get(key)
        if rep is None:
            rep = build_aoi_report(self.ctrl, name, self.xref, self.findings)
            self._aoi_reports[key] = rep
        return rep

    def aoi_markdown(self, name: str) -> str:
        """Documentação + roteiro de debug da AOI gerados só pelo parser (sem custo de LLM)."""
        return aoi_markdown(self.aoi_report(name))

    def aoi_rows(self) -> List[Dict[str, Any]]:
        return aoi_summary_rows(self.ctrl, self.xref, self.findings)

    def aoi_findings(self, name: str) -> List[Finding]:
        return list(self.aoi_report(name).findings)

    def config_sheet(self, name: str, mode: str = "config"):
        """Tabela de configuração das instâncias de uma AOI (None se não houver instâncias com dados)."""
        from l5x_core.aoi_config import build_sheets

        sheets = build_sheets(self.ctrl, self.xref, mode=mode, aoi=name)
        return sheets[0] if sheets else None

    def config_xlsx(self, mode: str = "config") -> bytes:
        """Excel com uma aba por AOI (cacheado por modo)."""
        from l5x_core.aoi_config import build_sheets, xlsx_bytes

        cache = self.__dict__.setdefault("_config_xlsx", {})
        if mode not in cache:
            cache[mode] = xlsx_bytes(build_sheets(self.ctrl, self.xref, mode=mode), mode)
        return cache[mode]

    # ---- relatórios determinísticos (cacheados) ------------------------------
    def _memo(self, key: str, build: Callable[[], Any]) -> Any:
        cache = self.__dict__.setdefault("_memo_cache", {})
        if key not in cache:
            cache[key] = build()
        return cache[key]

    def alarm_table(self) -> Tuple[List[str], List[Dict[str, Any]]]:
        from l5x_core.alarms import alarm_rows

        return self._memo("alarms", lambda: alarm_rows(self.ctrl, self.xref))

    def io_tables(self) -> List[Tuple[str, List[str], List[Dict[str, Any]]]]:
        from l5x_core.io_list import io_tables

        return self._memo("io", lambda: io_tables(self.ctrl, self.xref))

    def drift_rows(self, mode: str = "config") -> List[Dict[str, Any]]:
        from l5x_core.drift import instance_drift

        return self._memo(f"drift_{mode}", lambda: instance_drift(self.ctrl, self.xref, mode=mode))

    def diff_to(self, other: "ProjectContext"):
        """Diff semântico deste projeto (antigo) para ``other`` (novo)."""
        from l5x_core.diff import diff_projects

        return diff_projects(self.ctrl, other.ctrl, self.findings, other.findings)

    def library_drift_with(self, others: Sequence["ProjectContext"]) -> List[Dict[str, Any]]:
        from l5x_core.drift import library_drift

        return library_drift([self.ctrl] + [o.ctrl for o in others])

    def aoi_routines(self, name: str) -> List[Tuple[str, str]]:
        """Rotinas analisáveis da AOI como unidades (``aoi:<Nome>``, rotina)."""
        rep = self.aoi_report(name)
        scope = f"{AOI_PREFIX}{rep.aoi.name}"
        return [(scope, r.name) for r in rep.aoi.routines if r.has_content]

    def analyzed_routines(self, program: Optional[str] = None) -> List[Tuple[str, str]]:
        """Rotinas com conteúdo analisável (RLL/ST não protegidas), na ordem do projeto."""
        out: List[Tuple[str, str]] = []
        for p in self.ctrl.programs:
            if program and p.name.lower() != program.lower():
                continue
            for r in p.routines:
                if r.analyzed:
                    out.append((p.name, r.name))
        return out

    def content_routines(self, program: Optional[str] = None) -> List[Tuple[str, str]]:
        """Rotinas com conteúdo para o modelo (RLL/ST exatas e FBD com netlist), na ordem
        de execução: main e fault primeiro, depois as sub-rotinas na ordem dos JSR
        (busca em largura), por fim as rotinas nunca chamadas."""
        out: List[Tuple[str, str]] = []
        for p in self.ctrl.programs:
            if program and p.name.lower() != program.lower():
                continue
            by_name = {r.name.lower(): r for r in p.routines}
            order: List[str] = []
            seen: Set[str] = set()
            queue = [n for n in (p.main_routine, p.fault_routine) if n and n.lower() in by_name]
            while queue:
                name = queue.pop(0)
                key = name.lower()
                if key in seen or key not in by_name:
                    continue
                seen.add(key)
                order.append(by_name[key].name)
                queue.extend(c.strip() for c in by_name[key].calls if c.strip().lower() in by_name and c.strip().lower() not in seen)
            order.extend(r.name for r in p.routines if r.name.lower() not in seen)
            out.extend((p.name, n) for n in order if by_name[n.lower()].has_content)
        return out

    def all_routines(self, program: Optional[str] = None) -> List[Tuple[str, str]]:
        out: List[Tuple[str, str]] = []
        for p in self.ctrl.programs:
            if program and p.name.lower() != program.lower():
                continue
            out.extend((p.name, r.name) for r in p.routines)
        return out

    def is_safety(self, program: Optional[str] = None, routine: Optional[str] = None) -> bool:
        if program is None:
            return self.metrics.safety_programs > 0 or self.ctrl.safety_controller
        prog = self.ctrl.program(program)
        return bool(prog and prog.is_safety)

    # ---- findings -----------------------------------------------------------
    def findings_for(self, program: Optional[str] = None, routine: Optional[str] = None, min_severity: str = "info") -> List[Finding]:
        order = {s: i for i, s in enumerate(SEVERITIES)}
        limit = order.get(min_severity, 2)
        out = []
        for f in self.findings:
            if order.get(f.severity, 9) > limit:
                continue
            if program and (f.program or "").lower() != program.lower():
                # findings de escopo controlador (sem programa) entram só na visão geral
                if f.program is not None or routine is not None:
                    continue
            if routine and (f.routine or "").lower() != routine.lower():
                continue
            out.append(f)
        return out

    @staticmethod
    def findings_text(findings: Sequence[Finding], limit: int = 80) -> str:
        if not findings:
            return "(nenhum finding do parser para este escopo)"
        lines = []
        for f in findings[:limit]:
            safety = " [SAFETY]" if f.safety else ""
            ev = f" — evidência: {f.evidence}" if f.evidence else ""
            lines.append(f"- [{f.severity}] {f.rule} @ {f.location()}: {f.message}{safety}{ev}")
        if len(findings) > limit:
            lines.append(f"- … (+{len(findings) - limit} findings omitidos; priorizados por severidade)")
        return "\n".join(lines)

    # ---- deterministic tables for documentation ------------------------------
    def timer_table(self, program: Optional[str] = None) -> str:
        rows: List[str] = ["tag | instrução | preset | local"]
        for p in self.ctrl.programs:
            if program and p.name.lower() != program.lower():
                continue
            for r in p.routines:
                for rung in r.rungs:
                    for ins in rung.instructions:
                        if ins.name.upper() not in TIMER_COUNTER or not ins.operands:
                            continue
                        struct = ins.operands[0]
                        preset = "?"
                        if len(ins.operands) > 1:
                            op = ins.operands[1]
                            if op.kind == "literal":
                                preset = op.raw
                            elif op.kind == "tag":
                                preset = f"tag {op.tag_path}"
                            elif op.kind == "unused" and struct.kind == "tag":
                                tag = self.xref.resolve(p.name, struct.tag_base or "", struct.program)
                                if tag is not None and tag.values is not None:
                                    from l5x_core.analysis import _lookup_value

                                    rest = (struct.tag_path or "")[len(struct.tag_base or ""):]
                                    val = _lookup_value(tag.values, rest)
                                    if isinstance(val, dict) and val.get("PRE") is not None:
                                        preset = f"{val['PRE']} (valor da tag)"
                        rows.append(f"{struct.tag_path or struct.raw} | {ins.name} | {preset} | {p.name}/{r.name} rung {rung.number}")
        if len(rows) == 1:
            return "(nenhum timer/contador em ladder neste escopo)"
        return "\n".join(rows)

    def io_table(self, program: Optional[str] = None) -> str:
        modules = {m.name.lower(): m for m in self.ctrl.modules}
        rows: List[str] = ["módulo | catálogo | descrição | rotinas que usam"]
        for mod_key, usages in sorted(self.xref.io_usages.items()):
            where = sorted({f"{u.scope}/{u.routine}" for u in usages if not program or u.scope.lower() == program.lower()})
            if not where:
                continue
            m = modules.get(mod_key)
            name = m.name if m else mod_key
            rows.append(f"{name} | {(m.catalog if m else '?') or '?'} | {(m.description if m else '') or ''} | {', '.join(where[:6])}{' …' if len(where) > 6 else ''}")
        if len(rows) == 1:
            return "(nenhuma referência direta a I/O de módulo neste escopo)"
        return "\n".join(rows)

    # ---- retrieval support ------------------------------------------------
    def routine_word_index(self) -> Dict[Tuple[str, str], Set[str]]:
        if self._word_index is None:
            index: Dict[Tuple[str, str], Set[str]] = {}
            for p in self.ctrl.programs:
                for r in p.routines:
                    words: Set[str] = set()
                    desc_words: Set[str] = set()
                    if r.description:
                        desc_words = set(retrieval._WORD_RE.findall(r.description.lower()))
                        words |= desc_words
                    for rung in r.rungs:
                        if rung.comment:
                            words |= set(retrieval._WORD_RE.findall(rung.comment.lower()))
                    index[(p.name, r.name)] = words
                    self._desc_index[(p.name, r.name)] = desc_words
            self._word_index = index
        return self._word_index

    def routine_description_words(self, program: str, routine: str) -> Set[str]:
        self.routine_word_index()
        return self._desc_index.get((program, routine), set())


# ----------------------------------------------------------------------------
# Results
# ----------------------------------------------------------------------------


@dataclass
class Usage:
    input_tokens: int = 0
    output_tokens: int = 0
    cache_read_tokens: int = 0
    cache_write_tokens: int = 0

    def add(self, other: "Usage") -> None:
        self.input_tokens += other.input_tokens
        self.output_tokens += other.output_tokens
        self.cache_read_tokens += other.cache_read_tokens
        self.cache_write_tokens += other.cache_write_tokens


@dataclass
class CostEstimate:
    """Custo previsto de uma ação, calculado antes da chamada a partir do tamanho do contexto."""

    mode: str
    model: str
    calls: int
    input_tokens: int
    output_tokens: int
    usd: float
    credits: int

    def summary(self) -> str:
        def k(n: int) -> str:
            return f"{n:,}".replace(",", ".")

        return f"{self.model} · ~{k(self.input_tokens)} tokens de entrada, ~{k(self.output_tokens)} de saída, {self.calls} chamada(s) · ≈ US$ {self.usd:.2f}"


@dataclass
class LLMResult:
    text: str
    mode: str
    model: str
    usage: Usage = field(default_factory=Usage)
    latency_ms: int = 0
    calls: int = 1
    stop_reason: Optional[str] = None
    truncated: bool = False  # alguma chamada parou por max_tokens
    refusal: bool = False
    partial_results: List[str] = field(default_factory=list)  # map-reduce: saídas parciais
    context_tokens: int = 0  # estimativa do maior contexto enviado


# ----------------------------------------------------------------------------
# Prompt loading
# ----------------------------------------------------------------------------


def load_prompt(name: str, prompts_dir: Path = PROMPTS_DIR) -> str:
    path = prompts_dir / f"{name}.md"
    return path.read_text(encoding="utf-8")


def render(template: str, **values: str) -> str:
    """Substitui {chave} sem interpretar chaves ausentes (texto de rung pode ter chaves)."""

    def repl(m: re.Match) -> str:
        key = m.group(1)
        return values[key] if key in values else m.group(0)

    return re.sub(r"\{([a-z_]+)\}", repl, template)


# ----------------------------------------------------------------------------
# Client
# ----------------------------------------------------------------------------


OnText = Optional[Callable[[str], None]]


class LLMClient:
    def __init__(self, config: LLMConfig, client: Any = None, prompts_dir: Path = PROMPTS_DIR) -> None:
        self.config = config
        self.prompts_dir = prompts_dir
        self._client = client
        self._prompts: Dict[str, str] = {}

    # ---- lazy SDK client -----------------------------------------------------
    @property
    def client(self) -> Any:
        if self._client is None:
            try:
                import anthropic
            except ImportError as exc:  # pragma: no cover
                raise ConfigError("SDK 'anthropic' não instalado (pip install anthropic).") from exc
            try:
                self._client = anthropic.Anthropic(api_key=self.config.api_key) if self.config.api_key else anthropic.Anthropic()
            except Exception as exc:
                raise ConfigError(f"Não foi possível criar o cliente Anthropic: {exc}") from exc
        return self._client

    def prompt(self, name: str) -> str:
        if name not in self._prompts:
            self._prompts[name] = load_prompt(name, self.prompts_dir)
        return self._prompts[name]

    # ---- system block (fixed rules + cached overview) ------------------------
    def system_blocks(self, ctx: ProjectContext, language: Optional[str] = None) -> List[Dict[str, Any]]:
        lang = LANGUAGES.get(language or self.config.language, LANGUAGES["pt-BR"])
        rules = render(self.prompt("system"), language=lang)
        overview = "# Índice do projeto (gerado pelo parser; única fonte de verdade)\n\n" + ctx.overview.text
        return [
            {"type": "text", "text": rules},
            {"type": "text", "text": overview, "cache_control": {"type": "ephemeral"}},
        ]

    def system_tokens(self, ctx: ProjectContext, language: Optional[str] = None) -> int:
        return sum(estimate_tokens(b["text"]) for b in self.system_blocks(ctx, language))

    def budget_for_content(self, ctx: ProjectContext, language: Optional[str] = None, reserve: int = 1500) -> int:
        """Tokens disponíveis para rotinas/findings na mensagem do usuário."""
        return max(500, self.config.max_context_tokens - self.system_tokens(ctx, language) - reserve)

    # ---- modes ----------------------------------------------------------------
    def _explain_plan(self, ctx: ProjectContext, program: str, routine: Optional[str], language: Optional[str]) -> Tuple[str, List[str]]:
        if routine:
            units = [(program, routine)]
            target = f"{program} / {routine}"
        else:
            units = ctx.content_routines(program) or ctx.all_routines(program)
            target = f"programa {program} (main, fault e sub-rotinas na ordem dos JSR)"
            if not units:
                raise LLMError(f"Programa '{program}' não encontrado ou sem rotinas.", kind="api")
        return target, self._pack_units(ctx, units, self.budget_for_content(ctx, language))

    def explain(self, ctx: ProjectContext, program: str, routine: Optional[str] = None, language: Optional[str] = None, on_text: OnText = None, deep: bool = False) -> LLMResult:
        """Explica uma rotina, ou o programa inteiro (routine=None) rotina a rotina.
        ``deep=True`` usa o MODEL_DEEP (análise profunda) em vez do MODEL_FAST."""
        mode = "explain"
        target, chunks = self._explain_plan(ctx, program, routine, language)
        template = self.prompt("explain")
        if len(chunks) == 1:
            user = render(template, target=target, content=chunks[0])
            return self._call(mode, ctx, user, language, on_text, deep=deep)
        # map: explicações por parte, concatenadas (sem reduce)
        results: List[LLMResult] = []
        for i, chunk in enumerate(chunks, 1):
            user = render(template, target=f"{target} — parte {i}/{len(chunks)}", content=chunk)
            results.append(self._call(mode, ctx, user, language, on_text, deep=deep))
        return self._merge(results, mode, joiner="\n\n---\n\n")

    # ---- custo previsto ---------------------------------------------------------
    def credits_for_usd(self, usd: float) -> int:
        import math

        return max(1, int(math.ceil(usd * self.config.markup / self.config.credit_usd - 1e-9)))

    def estimate(self, mode: str, ctx: ProjectContext, program: Optional[str] = None, routine: Optional[str] = None, aoi: Optional[str] = None,
                 question: str = "", history: Sequence[Dict[str, str]] = (), language: Optional[str] = None, deep: bool = False) -> CostEstimate:
        """Custo previsto da ação (tokens, US$ e créditos) sem chamar o modelo. Reproduz a
        seleção de conteúdo de cada modo; a saída usa a média medida por modo, limitada ao teto."""
        sys_tokens = self.system_tokens(ctx, language)
        expected_out = min(EXPECTED_OUTPUT.get(mode, 4000), self.config.output_cap(mode))
        reduce_calls = 0
        if mode == "explain":
            if not program:
                raise LLMError("Informe o programa.", kind="api")
            _, chunks = self._explain_plan(ctx, program, routine, language)
            extra = 0
        elif mode == "review":
            if routine:
                units = [(program, routine)]  # type: ignore[list-item]
            else:
                units = ctx.content_routines(program)
            if not units:
                raise LLMError("Nenhuma rotina com conteúdo (ladder, ST ou FBD) no escopo pedido.", kind="api")
            findings_txt = ctx.findings_text(ctx.findings_for(program, routine))
            extra = estimate_tokens(findings_txt)
            chunks = self._pack_units(ctx, units, max(500, self.budget_for_content(ctx, language) - extra))
            reduce_calls = 1 if len(chunks) > 1 else 0
        elif mode == "document":
            units = ctx.content_routines(program)
            if program and not units and not ctx.all_routines(program):
                raise LLMError(f"Programa '{program}' não encontrado.", kind="api")
            extra = sum(estimate_tokens(x) for x in (ctx.findings_text(ctx.findings_for(program, min_severity="warning"), limit=40), ctx.io_table(program), ctx.timer_table(program)))
            chunks = self._pack_units(ctx, units, max(500, self.budget_for_content(ctx, language) - extra)) if units else [""]
            reduce_calls = 1 if len(chunks) > 1 else 0
        elif mode == "aoi":
            try:
                rep = ctx.aoi_report(aoi or "")
            except KeyError as exc:
                raise LLMError(str(exc), kind="api") from exc
            interface = aoi_interface_text(ctx.ctrl, rep.aoi.name, ctx.xref, rep).text
            extra = estimate_tokens(interface) + estimate_tokens(ctx.findings_text(rep.findings, limit=60))
            units = ctx.aoi_routines(rep.aoi.name)
            chunks = self._pack_units(ctx, units, max(500, self.budget_for_content(ctx, language) - extra)) if units else [""]
            reduce_calls = 1 if len(chunks) > 1 else 0
        elif mode == "chat":
            history_tokens = sum(estimate_tokens(m.get("content", "")) for m in history)
            budget = self.budget_for_content(ctx, language) - history_tokens - estimate_tokens(question)
            selected = retrieval.select_routines(ctx, question or "rotina principal", max(1500, int(budget * 0.8)), history=history)
            extra = history_tokens + estimate_tokens(question) + 400
            chunks = ["\n".join(sr.text for _, sr in selected)]
        else:
            raise LLMError(f"modo desconhecido: {mode}", kind="api")
        model = self.config.model_for(mode, deep)
        price_in, price_out, _ = self.config.price_for(model)
        calls = len(chunks) + reduce_calls
        input_tokens = sum(sys_tokens + extra + estimate_tokens(c) for c in chunks) + reduce_calls * (sys_tokens + len(chunks) * expected_out)
        output_tokens = calls * expected_out
        usd = (input_tokens * price_in + output_tokens * price_out) / 1e6
        return CostEstimate(mode=mode, model=model, calls=calls, input_tokens=input_tokens, output_tokens=output_tokens, usd=round(usd, 4), credits=self.credits_for_usd(usd))

    def review(self, ctx: ProjectContext, program: Optional[str] = None, routine: Optional[str] = None, language: Optional[str] = None, on_text: OnText = None) -> LLMResult:
        mode = "review"
        if routine and not program:
            raise LLMError("Informe o programa da rotina a revisar.", kind="api")
        if routine:
            units = [(program, routine)]  # type: ignore[list-item]
            target = f"{program} / {routine}"
        elif program:
            units = ctx.content_routines(program)
            target = f"programa {program} (main, fault e sub-rotinas na ordem dos JSR)"
        else:
            units = ctx.content_routines()
            target = f"projeto {ctx.ctrl.name} (todas as rotinas com conteúdo)"
        if not units:
            raise LLMError("Nenhuma rotina com conteúdo (ladder, ST ou FBD) no escopo pedido.", kind="api")
        findings = ctx.findings_for(program, routine)
        findings_txt = ctx.findings_text(findings)
        budget = self.budget_for_content(ctx, language) - estimate_tokens(findings_txt)
        chunks = self._pack_units(ctx, units, max(500, budget))
        template = self.prompt("review")
        if len(chunks) == 1:
            user = render(template, target=target, findings=findings_txt, content=chunks[0])
            return self._call(mode, ctx, user, language, on_text)
        partials: List[LLMResult] = []
        for i, chunk in enumerate(chunks, 1):
            scoped = ctx.findings_text(self._findings_for_chunk(findings, chunk))
            user = render(template, target=f"{target} — parte {i}/{len(chunks)}", findings=scoped, content=chunk)
            partials.append(self._call(mode, ctx, user, language, None))
        return self._reduce(mode, ctx, target, partials, language, on_text)

    def document(self, ctx: ProjectContext, program: Optional[str] = None, language: Optional[str] = None, on_text: OnText = None) -> LLMResult:
        mode = "document"
        units = ctx.content_routines(program)
        if program and not units and not ctx.all_routines(program):
            raise LLMError(f"Programa '{program}' não encontrado.", kind="api")
        target = f"programa {program}" if program else f"projeto {ctx.ctrl.name}"
        findings_txt = ctx.findings_text(ctx.findings_for(program, min_severity="warning"), limit=40)
        io_txt = ctx.io_table(program)
        timer_txt = ctx.timer_table(program)
        fixed = estimate_tokens(findings_txt) + estimate_tokens(io_txt) + estimate_tokens(timer_txt)
        budget = max(500, self.budget_for_content(ctx, language) - fixed)
        chunks = self._pack_units(ctx, units, budget) if units else ["(sem rotinas com conteúdo neste escopo; documente a estrutura a partir da visão geral)"]
        template = self.prompt("document")
        if len(chunks) == 1:
            user = render(template, target=target, io_table=io_txt, timer_table=timer_txt, findings=findings_txt, content=chunks[0])
            return self._call(mode, ctx, user, language, on_text)
        partials: List[LLMResult] = []
        for i, chunk in enumerate(chunks, 1):
            user = render(template, target=f"{target} — parte {i}/{len(chunks)}", io_table=io_txt, timer_table=timer_txt, findings=findings_txt, content=chunk)
            partials.append(self._call(mode, ctx, user, language, None))
        return self._reduce(mode, ctx, target, partials, language, on_text)

    def aoi(self, ctx: ProjectContext, name: str, language: Optional[str] = None, on_text: OnText = None) -> LLMResult:
        """Análise + documentação + roteiro de debug de uma AOI: interface e uso no projeto
        (parser), findings da AOI e rotinas internas; map-reduce se a AOI for grande."""
        mode = "aoi"
        try:
            rep = ctx.aoi_report(name)
        except KeyError as exc:
            raise LLMError(str(exc), kind="api") from exc
        aoi = rep.aoi
        if aoi.protected:
            raise LLMError(f"AOI '{aoi.name}' tem source protection: não há conteúdo para analisar (só instâncias e chamadas, veja o relatório determinístico).", kind="api")
        interface = aoi_interface_text(ctx.ctrl, aoi.name, ctx.xref, rep).text
        findings_txt = ctx.findings_text(rep.findings, limit=60)
        fixed = estimate_tokens(interface) + estimate_tokens(findings_txt)
        budget = max(500, self.budget_for_content(ctx, language) - fixed)
        units = ctx.aoi_routines(aoi.name)
        chunks = self._pack_units(ctx, units, budget) if units else ["(a AOI não tem rotinas ladder/ST analisáveis; documente a partir da interface e das chamadas)"]
        template = self.prompt("aoi")
        if len(chunks) == 1:
            user = render(template, target=aoi.name, interface=interface, findings=findings_txt, content=chunks[0])
            return self._call(mode, ctx, user, language, on_text)
        partials: List[LLMResult] = []
        for i, chunk in enumerate(chunks, 1):
            user = render(template, target=f"{aoi.name} — parte {i}/{len(chunks)}", interface=interface, findings=findings_txt, content=chunk)
            partials.append(self._call(mode, ctx, user, language, None))
        return self._reduce(mode, ctx, f"AOI {aoi.name}", partials, language, on_text)

    def chat(self, ctx: ProjectContext, history: Sequence[Dict[str, str]], question: str, language: Optional[str] = None, on_text: OnText = None, max_routines: int = 8, deep: bool = False) -> LLMResult:
        """Pergunta livre. ``history`` = [{"role": "user"|"assistant", "content": texto}, ...]
        com as perguntas e respostas anteriores (sem as rotinas recuperadas, que são
        injetadas só na pergunta atual). O chamador anexa a pergunta e a resposta ao histórico."""
        mode = "chat"
        history_tokens = sum(estimate_tokens(m.get("content", "")) for m in history)
        budget = self.budget_for_content(ctx, language) - history_tokens - estimate_tokens(question)
        selected = retrieval.select_routines(ctx, question, max(1500, int(budget * 0.8)), max_routines=max_routines, history=history)
        if selected:
            parts = []
            introduced: Set[str] = set()
            scoped_findings: List[Finding] = []
            for hit, sr in selected:
                if hit.program.startswith(AOI_PREFIX):
                    name = hit.program[len(AOI_PREFIX):]
                    if name.lower() not in introduced:
                        # interface, instâncias e findings da AOI acompanham a primeira rotina dela
                        introduced.add(name.lower())
                        try:
                            rep = ctx.aoi_report(name)
                            parts.append(aoi_interface_text(ctx.ctrl, name, ctx.xref, rep, max_instances=8, max_calls=8).text)
                            scoped_findings.extend(rep.findings)
                        except KeyError:  # pragma: no cover - nome vem do próprio índice
                            pass
                else:
                    scoped_findings.extend(ctx.findings_for(hit.program, hit.routine))
                parts.append(f"<!-- selecionada por: {', '.join(hit.reasons)} -->\n{sr.text}")
            content = "\n\n".join(parts)
            findings_txt = ctx.findings_text(scoped_findings, limit=40)
        else:
            content = "(nenhuma rotina selecionada; responda com base na visão geral e peça ao usuário para indicar a rotina ou a AOI)"
            findings_txt = "(nenhum)"
        user = render(self.prompt("chat"), question=question, content=content, findings=findings_txt)
        messages = [{"role": m["role"], "content": m["content"]} for m in history if m.get("role") in ("user", "assistant") and m.get("content")]
        messages.append({"role": "user", "content": user})
        result = self._call(mode, ctx, None, language, on_text, messages=messages, deep=deep)
        result.partial_results = [f"{h.program}/{h.routine}" for h, _ in selected]
        return result

    # ---- packing / map-reduce -------------------------------------------------
    def _pack_units(self, ctx: ProjectContext, units: Sequence[Tuple[str, str]], budget: int) -> List[str]:
        """Agrupa textos de rotina em pedaços que cabem no orçamento (greedy, em ordem)."""
        chunks: List[str] = []
        cur: List[str] = []
        cur_tokens = 0
        for program, routine in units:
            sr = ctx.routine_text(program, routine)
            if sr.tokens > budget:
                if cur:
                    chunks.append("\n\n".join(cur))
                    cur, cur_tokens = [], 0
                chunks.extend(self._split_text(sr.text, budget))
                continue
            if cur and cur_tokens + sr.tokens > budget:
                chunks.append("\n\n".join(cur))
                cur, cur_tokens = [], 0
            cur.append(sr.text)
            cur_tokens += sr.tokens
        if cur:
            chunks.append("\n\n".join(cur))
        return chunks

    @staticmethod
    def _split_text(text: str, budget: int) -> List[str]:
        """Divide o texto de uma rotina grande em partes nas fronteiras de rung/linha."""
        limit = int(budget * 3.5)
        lines = text.splitlines(keepends=True)
        header = lines[0] if lines else ""
        parts: List[str] = []
        cur: List[str] = []
        size = 0
        for ln in lines:
            boundary = (
                ln.startswith("Rung ")
                or re.match(r"L\d+:", ln) is not None
                or ln.startswith("## ")
                or ln.startswith("- ")  # lista de tags usadas
                or ln.startswith("  // ")  # comentário de rung
            )
            if cur and size + len(ln) > limit and (boundary or size > 2 * limit):
                parts.append("".join(cur))
                cur = [header.rstrip("\n") + " (continuação)\n"] if header else []
                size = len(cur[0]) if cur else 0
            cur.append(ln)
            size += len(ln)
        if cur:
            parts.append("".join(cur))
        return parts or [text[:limit]]

    @staticmethod
    def _findings_for_chunk(findings: Sequence[Finding], chunk: str) -> List[Finding]:
        """Findings cuja rotina aparece no pedaço (pelo cabeçalho '# Programa / Rotina')."""
        names = {m.group(1).lower() for m in re.finditer(r"^# [^/\n]+ / ([^\[\n]+?) \[", chunk, re.M)}
        if not names:
            return list(findings)
        return [f for f in findings if (f.routine or "").lower() in names or f.routine is None]

    def _reduce(self, mode: str, ctx: ProjectContext, target: str, partials: List[LLMResult], language: Optional[str], on_text: OnText) -> LLMResult:
        """Consolida resultados parciais; recursivo se não couberem de uma vez."""
        template = self.prompt("consolidate")
        budget = self.budget_for_content(ctx, language)
        texts = [p.text for p in partials]
        total_usage = Usage()
        total_calls = 0
        total_latency = 0
        truncated = any(p.truncated for p in partials)
        refusal = any(p.refusal for p in partials)
        for p in partials:
            total_usage.add(p.usage)
            total_calls += p.calls
            total_latency += p.latency_ms
        level_inputs = texts
        while True:
            groups: List[List[str]] = []
            cur: List[str] = []
            cur_tokens = 0
            for t in level_inputs:
                tk = estimate_tokens(t)
                if cur and cur_tokens + tk > budget:
                    groups.append(cur)
                    cur, cur_tokens = [], 0
                cur.append(t)
                cur_tokens += tk
            if cur:
                groups.append(cur)
            final = len(groups) == 1
            outputs: List[str] = []
            for g in groups:
                content = "\n\n".join(f"<!-- parte {i + 1}/{len(g)} -->\n{t}" for i, t in enumerate(g))
                user = render(template, count=str(len(g)), mode=mode, target=target, content=content)
                r = self._call(mode, ctx, user, language, on_text if final else None)
                total_usage.add(r.usage)
                total_calls += r.calls
                total_latency += r.latency_ms
                truncated = truncated or r.truncated
                refusal = refusal or r.refusal
                outputs.append(r.text)
            if final:
                return LLMResult(text=outputs[0], mode=mode, model=self.config.model_for(mode), usage=total_usage, latency_ms=total_latency, calls=total_calls, truncated=truncated, refusal=refusal, partial_results=texts)
            level_inputs = outputs

    @staticmethod
    def _merge(results: List[LLMResult], mode: str, joiner: str) -> LLMResult:
        usage = Usage()
        for r in results:
            usage.add(r.usage)
        return LLMResult(
            text=joiner.join(r.text for r in results), mode=mode, model=results[0].model, usage=usage,
            latency_ms=sum(r.latency_ms for r in results), calls=sum(r.calls for r in results),
            truncated=any(r.truncated for r in results), refusal=any(r.refusal for r in results),
            partial_results=[r.text for r in results], context_tokens=max(r.context_tokens for r in results),
        )

    # ---- the actual API call -------------------------------------------------
    def _call(self, mode: str, ctx: ProjectContext, user: Optional[str], language: Optional[str], on_text: OnText, messages: Optional[List[Dict[str, Any]]] = None, deep: bool = False) -> LLMResult:
        model = self.config.model_for(mode, deep)
        system = self.system_blocks(ctx, language)
        if messages is None:
            messages = [{"role": "user", "content": user or ""}]
        context_tokens = sum(estimate_tokens(b["text"]) for b in system) + sum(estimate_tokens(str(m["content"])) for m in messages)
        params: Dict[str, Any] = {
            "model": model,
            "max_tokens": self.config.output_cap(mode),
            "system": system,
            "messages": messages,
        }
        effort = self.config.effort_for(mode, deep)
        if effort:
            params["output_config"] = {"effort": effort}

        total_usage = Usage()
        total_latency = 0
        calls = 0
        parts: List[str] = []
        convo = list(messages)
        while True:
            params["messages"] = convo
            text, message, usage, latency_ms = self._stream_once(mode, model, params, context_tokens, on_text)
            calls += 1
            total_usage.add(usage)
            total_latency += latency_ms
            parts.append(text)
            stop_reason = getattr(message, "stop_reason", None)
            if stop_reason != "max_tokens" or calls > self.config.max_continuations:
                break
            # a resposta parou pelo limite de saída: pede para continuar de onde parou
            convo = convo + [
                {"role": "assistant", "content": text},
                {"role": "user", "content": self.CONTINUE_PROMPT},
            ]
        text = "".join(parts)
        refusal = stop_reason == "refusal"
        truncated = stop_reason == "max_tokens"
        if refusal:
            details = getattr(message, "stop_details", None)
            why = getattr(details, "explanation", None) if details else None
            text = (text + "\n\n" if text else "") + "O modelo recusou concluir esta resposta" + (f": {why}" if why else ".") + " Tente reformular o pedido ou reduzir o escopo."
        if truncated:
            text += f"\n\n*(resposta interrompida por limite de tokens de saída após {calls} chamada(s); aumente MAX_OUTPUT_TOKENS / LLM_MAX_CONTINUATIONS ou reduza o escopo)*"
        return LLMResult(text=text, mode=mode, model=model, usage=total_usage, latency_ms=total_latency, calls=calls, stop_reason=stop_reason, truncated=truncated, refusal=refusal, context_tokens=context_tokens)

    CONTINUE_PROMPT = "A resposta anterior foi cortada pelo limite de tokens de saída. Continue exatamente de onde parou, sem repetir o que já foi escrito e mantendo a mesma estrutura de seções."

    def _stream_once(self, mode: str, model: str, params: Dict[str, Any], context_tokens: int, on_text: OnText) -> Tuple[str, Any, Usage, int]:
        """Uma chamada em streaming; devolve (texto, mensagem final, uso, latência ms)."""
        t0 = time.perf_counter()
        text_parts: List[str] = []
        try:
            with self.client.messages.stream(**params) as stream:
                for delta in stream.text_stream:
                    text_parts.append(delta)
                    if on_text:
                        on_text(delta)
                message = stream.get_final_message()
        except Exception as exc:  # noqa: BLE001 - convertido em LLMError abaixo
            self._log(mode, model, Usage(), int((time.perf_counter() - t0) * 1000), context_tokens, status=type(exc).__name__)
            raise self._map_error(exc) from exc
        latency_ms = int((time.perf_counter() - t0) * 1000)
        usage = Usage(
            input_tokens=getattr(message.usage, "input_tokens", 0) or 0,
            output_tokens=getattr(message.usage, "output_tokens", 0) or 0,
            cache_read_tokens=getattr(message.usage, "cache_read_input_tokens", 0) or 0,
            cache_write_tokens=getattr(message.usage, "cache_creation_input_tokens", 0) or 0,
        )
        text = "".join(text_parts)
        if not text:
            text = "".join(getattr(b, "text", "") for b in getattr(message, "content", []) if getattr(b, "type", "") == "text")
        self._log(mode, model, usage, latency_ms, context_tokens, status=getattr(message, "stop_reason", None) or "ok")
        return text, message, usage, latency_ms

    # ---- errors ---------------------------------------------------------------
    @staticmethod
    def _map_error(exc: BaseException) -> LLMError:
        try:
            import anthropic
        except ImportError:  # pragma: no cover
            return LLMError(f"Falha na chamada ao modelo: {exc}", kind="api", cause=exc)
        if isinstance(exc, anthropic.AuthenticationError):
            return LLMError("Chave da API inválida ou ausente (ANTHROPIC_API_KEY).", kind="auth", cause=exc)
        if isinstance(exc, anthropic.PermissionDeniedError):
            return LLMError("A chave da API não tem permissão para este modelo.", kind="auth", cause=exc)
        if isinstance(exc, anthropic.NotFoundError):
            return LLMError("Modelo não encontrado: verifique MODEL_FAST / MODEL_DEEP (IDs em docs.claude.com/en/docs/about-claude/models).", kind="model", cause=exc)
        if isinstance(exc, anthropic.RateLimitError):
            retry_after = None
            try:
                retry_after = exc.response.headers.get("retry-after")
            except Exception:  # pragma: no cover
                pass
            extra = f" Tente novamente em {retry_after}s." if retry_after else " Tente novamente em instantes."
            return LLMError("Limite de requisições da API atingido." + extra, kind="rate_limit", cause=exc)
        if isinstance(exc, anthropic.BadRequestError):
            msg = str(getattr(exc, "message", exc)).lower()
            if "too long" in msg or "prompt is too long" in msg or "context" in msg and "exceed" in msg:
                return LLMError("O contexto enviado excede o limite do modelo. Reduza MAX_CONTEXT_TOKENS ou escolha uma rotina menor.", kind="context", cause=exc)
            return LLMError(f"Pedido rejeitado pela API: {getattr(exc, 'message', exc)}", kind="api", cause=exc)
        if isinstance(exc, anthropic.APIStatusError):
            if exc.status_code >= 500:
                return LLMError("Erro temporário no serviço do modelo. Tente novamente.", kind="api", cause=exc)
            return LLMError(f"Erro da API ({exc.status_code}): {getattr(exc, 'message', exc)}", kind="api", cause=exc)
        if isinstance(exc, anthropic.APIConnectionError):
            return LLMError("Sem conexão com a API (rede/timeout).", kind="connection", cause=exc)
        if isinstance(exc, LLMError):
            return exc
        return LLMError(f"Falha na chamada ao modelo: {exc}", kind="api", cause=exc)

    # ---- logging (metrics only, never program content) ------------------------
    _LOG_FIELDS = ["timestamp", "mode", "model", "input_tokens", "output_tokens", "cache_read_tokens", "cache_write_tokens", "context_tokens_est", "latency_ms", "status"]

    def _log(self, mode: str, model: str, usage: Usage, latency_ms: int, context_tokens: int, status: str) -> None:
        path = self.config.log_path
        if not path:
            return
        try:
            p = Path(path)
            new = not p.exists() or p.stat().st_size == 0
            with p.open("a", newline="", encoding="utf-8") as fh:
                w = csv.DictWriter(fh, fieldnames=self._LOG_FIELDS)
                if new:
                    w.writeheader()
                w.writerow({
                    "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    "mode": mode, "model": model,
                    "input_tokens": usage.input_tokens, "output_tokens": usage.output_tokens,
                    "cache_read_tokens": usage.cache_read_tokens, "cache_write_tokens": usage.cache_write_tokens,
                    "context_tokens_est": context_tokens, "latency_ms": latency_ms, "status": status,
                })
        except OSError:  # pragma: no cover - log nunca derruba a chamada
            pass


# ----------------------------------------------------------------------------
# CLI (verificação manual com chave real)
# ----------------------------------------------------------------------------


def main(argv: Optional[List[str]] = None) -> int:  # pragma: no cover - requer API
    import argparse

    from l5x_core.parser import parse_file, L5XError

    ap = argparse.ArgumentParser(prog="app.llm", description="Chama o modelo sobre um L5X (requer ANTHROPIC_API_KEY e MODEL_*).")
    ap.add_argument("file")
    ap.add_argument("--mode", choices=MODES, default="explain")
    ap.add_argument("--target", help="PROGRAMA ou PROGRAMA/ROTINA (modo aoi: NOME_DA_AOI)")
    ap.add_argument("--question", help="pergunta (modo chat)")
    ap.add_argument("--lang", choices=sorted(LANGUAGES), default=None)
    ap.add_argument("--deep", action="store_true", help="explain/chat com o MODEL_DEEP (análise profunda)")
    ap.add_argument("--estimate", action="store_true", help="só estima tokens, custo e créditos, sem chamar o modelo")
    args = ap.parse_args(argv)
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
    try:
        cfg = LLMConfig.from_env()
        ctrl = parse_file(args.file)
    except (LLMError, L5XError) as exc:
        print(f"erro: {exc}", file=sys.stderr)
        return 2
    ctx = ProjectContext(ctrl)
    llm = LLMClient(cfg)
    program = routine = None
    if args.target:
        program, _, routine = args.target.partition("/")
        routine = routine or None

    def echo(t: str) -> None:
        print(t, end="", flush=True)

    if args.estimate:
        try:
            est = llm.estimate(args.mode, ctx, program=program, routine=routine, aoi=args.target if args.mode == "aoi" else None, question=args.question or "", deep=args.deep)
        except LLMError as exc:
            print(f"erro: {exc.message}", file=sys.stderr)
            return 2
        print(f"{est.mode}: {est.summary()} -> {est.credits} crédito(s) (1 crédito = US$ {cfg.credit_usd}, margem {cfg.markup}x)")
        return 0
    try:
        if args.mode == "explain":
            if not program:
                print("erro: --target PROGRAMA[/ROTINA]", file=sys.stderr)
                return 2
            r = llm.explain(ctx, program, routine, args.lang, on_text=echo, deep=args.deep)
        elif args.mode == "review":
            r = llm.review(ctx, program, routine, args.lang, on_text=echo)
        elif args.mode == "document":
            r = llm.document(ctx, program, args.lang, on_text=echo)
        elif args.mode == "aoi":
            if not args.target:
                print("erro: --target NOME_DA_AOI", file=sys.stderr)
                return 2
            r = llm.aoi(ctx, args.target, args.lang, on_text=echo)
        else:
            if not args.question:
                print("erro: --question", file=sys.stderr)
                return 2
            r = llm.chat(ctx, [], args.question, args.lang, on_text=echo, deep=args.deep)
    except LLMError as exc:
        print(f"\nerro ({exc.kind}): {exc.message}", file=sys.stderr)
        return 1
    if r.calls > 1 or not r.text.endswith("\n"):
        print()
    print(f"\n[{r.mode} | {r.model} | chamadas={r.calls} | in={r.usage.input_tokens} out={r.usage.output_tokens} cache_read={r.usage.cache_read_tokens} cache_write={r.usage.cache_write_tokens} | {r.latency_ms} ms | stop={r.stop_reason}]", file=sys.stderr)
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())

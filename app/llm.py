"""Camada LLM: cliente Anthropic, prompts, cache, orçamento de contexto e log.

Uso típico:

    cfg = LLMConfig.from_env()
    ctx = ProjectContext(ctrl)                 # ctrl = l5x_core.parse_file(...)
    llm = LLMClient(cfg)
    result = llm.explain(ctx, "MainProgram", "MainRoutine")
    result = llm.review(ctx)                   # projeto inteiro (map-reduce)
    result = llm.document(ctx, program="MainProgram")
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
from l5x_core.model import Controller, Finding, Metrics, SEVERITIES
from l5x_core.rung_text import TIMER_COUNTER
from l5x_core.serialize import Serialized, estimate_tokens, project_overview, routine_text
from l5x_core.xref import CrossRef, build_xref

from app import retrieval

PROMPTS_DIR = Path(__file__).parent / "prompts"

LANGUAGES: Dict[str, str] = {
    "pt-BR": "português do Brasil",
    "en": "English",
    "es": "español",
}

MODES = ("explain", "review", "document", "chat")
DEEP_MODES = {"review", "document"}

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
    max_output_tokens: int = 16_000
    log_path: Optional[str] = "llm_calls.csv"
    language: str = "pt-BR"

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
        return cls(
            model_fast=model_fast,
            model_deep=model_deep or model_fast,
            api_key=get("ANTHROPIC_API_KEY"),
            effort_fast=get("LLM_EFFORT_FAST"),
            effort_deep=get("LLM_EFFORT_DEEP"),
            max_context_tokens=int(get("MAX_CONTEXT_TOKENS", "60000") or 60000),
            max_output_tokens=int(get("MAX_OUTPUT_TOKENS", "16000") or 16000),
            log_path=get("LLM_LOG_PATH", "llm_calls.csv"),
            language=language,
        )

    def model_for(self, mode: str) -> str:
        return self.model_deep if mode in DEEP_MODES else self.model_fast

    def effort_for(self, mode: str) -> Optional[str]:
        return self.effort_deep if mode in DEEP_MODES else self.effort_fast


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

    # ---- routines ----------------------------------------------------------
    def routine_text(self, program: str, routine: str) -> Serialized:
        key = (program.lower(), routine.lower())
        sr = self._routine_cache.get(key)
        if sr is None:
            sr = routine_text(self.ctrl, program, routine, self.xref)
            self._routine_cache[key] = sr
        return sr

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
    def explain(self, ctx: ProjectContext, program: str, routine: Optional[str] = None, language: Optional[str] = None, on_text: OnText = None) -> LLMResult:
        """Explica uma rotina, ou o programa inteiro (routine=None) rotina a rotina."""
        mode = "explain"
        if routine:
            units = [(program, routine)]
            target = f"{program} / {routine}"
        else:
            units = ctx.analyzed_routines(program) or ctx.all_routines(program)
            target = f"programa {program}"
            if not units:
                raise LLMError(f"Programa '{program}' não encontrado ou sem rotinas.", kind="api")
        chunks = self._pack_units(ctx, units, self.budget_for_content(ctx, language))
        template = self.prompt("explain")
        if len(chunks) == 1:
            user = render(template, target=target, content=chunks[0])
            return self._call(mode, ctx, user, language, on_text)
        # map: explicações por parte, concatenadas (sem reduce)
        results: List[LLMResult] = []
        for i, chunk in enumerate(chunks, 1):
            user = render(template, target=f"{target} — parte {i}/{len(chunks)}", content=chunk)
            results.append(self._call(mode, ctx, user, language, on_text))
        return self._merge(results, mode, joiner="\n\n---\n\n")

    def review(self, ctx: ProjectContext, program: Optional[str] = None, routine: Optional[str] = None, language: Optional[str] = None, on_text: OnText = None) -> LLMResult:
        mode = "review"
        if routine and not program:
            raise LLMError("Informe o programa da rotina a revisar.", kind="api")
        if routine:
            units = [(program, routine)]  # type: ignore[list-item]
            target = f"{program} / {routine}"
        elif program:
            units = ctx.analyzed_routines(program)
            target = f"programa {program}"
        else:
            units = ctx.analyzed_routines()
            target = f"projeto {ctx.ctrl.name} (todas as rotinas analisáveis)"
        if not units:
            raise LLMError("Nenhuma rotina analisável (ladder/ST) no escopo pedido.", kind="api")
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
        units = ctx.analyzed_routines(program)
        if program and not units and not ctx.all_routines(program):
            raise LLMError(f"Programa '{program}' não encontrado.", kind="api")
        target = f"programa {program}" if program else f"projeto {ctx.ctrl.name}"
        findings_txt = ctx.findings_text(ctx.findings_for(program, min_severity="warning"), limit=40)
        io_txt = ctx.io_table(program)
        timer_txt = ctx.timer_table(program)
        fixed = estimate_tokens(findings_txt) + estimate_tokens(io_txt) + estimate_tokens(timer_txt)
        budget = max(500, self.budget_for_content(ctx, language) - fixed)
        chunks = self._pack_units(ctx, units, budget) if units else ["(sem rotinas ladder/ST analisáveis neste escopo; documente a estrutura a partir da visão geral)"]
        template = self.prompt("document")
        if len(chunks) == 1:
            user = render(template, target=target, io_table=io_txt, timer_table=timer_txt, findings=findings_txt, content=chunks[0])
            return self._call(mode, ctx, user, language, on_text)
        partials: List[LLMResult] = []
        for i, chunk in enumerate(chunks, 1):
            user = render(template, target=f"{target} — parte {i}/{len(chunks)}", io_table=io_txt, timer_table=timer_txt, findings=findings_txt, content=chunk)
            partials.append(self._call(mode, ctx, user, language, None))
        return self._reduce(mode, ctx, target, partials, language, on_text)

    def chat(self, ctx: ProjectContext, history: Sequence[Dict[str, str]], question: str, language: Optional[str] = None, on_text: OnText = None, max_routines: int = 8) -> LLMResult:
        """Pergunta livre. ``history`` = [{"role": "user"|"assistant", "content": texto}, ...]
        com as perguntas e respostas anteriores (sem as rotinas recuperadas, que são
        injetadas só na pergunta atual). O chamador anexa a pergunta e a resposta ao histórico."""
        mode = "chat"
        history_tokens = sum(estimate_tokens(m.get("content", "")) for m in history)
        budget = self.budget_for_content(ctx, language) - history_tokens - estimate_tokens(question)
        selected = retrieval.select_routines(ctx, question, max(1500, budget), max_routines=max_routines)
        if selected:
            parts = []
            for hit, sr in selected:
                parts.append(f"<!-- selecionada por: {', '.join(hit.reasons)} -->\n{sr.text}")
            content = "\n\n".join(parts)
        else:
            content = "(nenhuma rotina selecionada; responda com base na visão geral e peça ao usuário para indicar a rotina)"
        user = render(self.prompt("chat"), question=question, content=content)
        messages = [{"role": m["role"], "content": m["content"]} for m in history if m.get("role") in ("user", "assistant") and m.get("content")]
        messages.append({"role": "user", "content": user})
        result = self._call(mode, ctx, None, language, on_text, messages=messages)
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
    def _call(self, mode: str, ctx: ProjectContext, user: Optional[str], language: Optional[str], on_text: OnText, messages: Optional[List[Dict[str, Any]]] = None) -> LLMResult:
        model = self.config.model_for(mode)
        system = self.system_blocks(ctx, language)
        if messages is None:
            messages = [{"role": "user", "content": user or ""}]
        context_tokens = sum(estimate_tokens(b["text"]) for b in system) + sum(estimate_tokens(str(m["content"])) for m in messages)
        params: Dict[str, Any] = {
            "model": model,
            "max_tokens": self.config.max_output_tokens,
            "system": system,
            "messages": messages,
        }
        effort = self.config.effort_for(mode)
        if effort:
            params["output_config"] = {"effort": effort}

        t0 = time.perf_counter()
        text_parts: List[str] = []
        error: Optional[str] = None
        try:
            with self.client.messages.stream(**params) as stream:
                for delta in stream.text_stream:
                    text_parts.append(delta)
                    if on_text:
                        on_text(delta)
                message = stream.get_final_message()
        except Exception as exc:  # noqa: BLE001 - convertido em LLMError abaixo
            error = type(exc).__name__
            self._log(mode, model, Usage(), int((time.perf_counter() - t0) * 1000), context_tokens, status=error)
            raise self._map_error(exc) from exc
        latency_ms = int((time.perf_counter() - t0) * 1000)

        usage = Usage(
            input_tokens=getattr(message.usage, "input_tokens", 0) or 0,
            output_tokens=getattr(message.usage, "output_tokens", 0) or 0,
            cache_read_tokens=getattr(message.usage, "cache_read_input_tokens", 0) or 0,
            cache_write_tokens=getattr(message.usage, "cache_creation_input_tokens", 0) or 0,
        )
        stop_reason = getattr(message, "stop_reason", None)
        text = "".join(text_parts)
        if not text:
            text = "".join(getattr(b, "text", "") for b in getattr(message, "content", []) if getattr(b, "type", "") == "text")
        refusal = stop_reason == "refusal"
        truncated = stop_reason == "max_tokens"
        if refusal:
            details = getattr(message, "stop_details", None)
            why = getattr(details, "explanation", None) if details else None
            text = (text + "\n\n" if text else "") + "O modelo recusou concluir esta resposta" + (f": {why}" if why else ".") + " Tente reformular o pedido ou reduzir o escopo."
        if truncated:
            text += "\n\n*(resposta interrompida por limite de tokens de saída; aumente MAX_OUTPUT_TOKENS ou reduza o escopo)*"
        self._log(mode, model, usage, latency_ms, context_tokens, status=stop_reason or "ok")
        return LLMResult(text=text, mode=mode, model=model, usage=usage, latency_ms=latency_ms, calls=1, stop_reason=stop_reason, truncated=truncated, refusal=refusal, context_tokens=context_tokens)

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
    ap.add_argument("--target", help="PROGRAMA ou PROGRAMA/ROTINA")
    ap.add_argument("--question", help="pergunta (modo chat)")
    ap.add_argument("--lang", choices=sorted(LANGUAGES), default=None)
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

    try:
        if args.mode == "explain":
            if not program:
                print("erro: --target PROGRAMA[/ROTINA]", file=sys.stderr)
                return 2
            r = llm.explain(ctx, program, routine, args.lang, on_text=echo)
        elif args.mode == "review":
            r = llm.review(ctx, program, routine, args.lang, on_text=echo)
        elif args.mode == "document":
            r = llm.document(ctx, program, args.lang, on_text=echo)
        else:
            if not args.question:
                print("erro: --question", file=sys.stderr)
                return 2
            r = llm.chat(ctx, [], args.question, args.lang, on_text=echo)
    except LLMError as exc:
        print(f"\nerro ({exc.kind}): {exc.message}", file=sys.stderr)
        return 1
    if r.calls > 1 or not r.text.endswith("\n"):
        print()
    print(f"\n[{r.mode} | {r.model} | chamadas={r.calls} | in={r.usage.input_tokens} out={r.usage.output_tokens} cache_read={r.usage.cache_read_tokens} cache_write={r.usage.cache_write_tokens} | {r.latency_ms} ms | stop={r.stop_reason}]", file=sys.stderr)
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())

"""Compact, LLM-oriented views of a parsed project.

Everything returned here is derived from the parsed model (never raw XML)
and comes with a token estimate (chars / 3.5).
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from typing import Dict, Iterable, List, Optional, Tuple

from .model import Controller, Finding, Metrics, Program, Routine, Tag, SEVERITIES
from .xref import CrossRef, build_xref

CHARS_PER_TOKEN = 3.5


def estimate_tokens(text: str) -> int:
    return int(len(text) / CHARS_PER_TOKEN) + 1


@dataclass
class Serialized:
    text: str
    tokens: int
    truncated: bool = False

    def __str__(self) -> str:  # pragma: no cover - convenience
        return self.text


def _s(value: Optional[str], limit: int = 120) -> str:
    if not value:
        return ""
    one = " ".join(value.split())
    return one if len(one) <= limit else one[: limit - 1] + "…"


def _routine_size(r: Routine) -> str:
    if r.protected:
        return "protegida"
    if r.type == "RLL":
        return f"{len(r.rungs)} rungs"
    if r.type == "ST":
        return f"{len(r.st_lines)} linhas"
    if r.type == "FBD":
        return f"{r.sheets} folha(s), não analisada em detalhe"
    if r.type == "SFC":
        return "não analisada em detalhe"
    return "?"


# --------------------------------------------------------------------------
# Project overview
# --------------------------------------------------------------------------
def project_overview(ctrl: Controller, findings: List[Finding], metrics: Metrics, max_tokens: int = 6000) -> Serialized:
    lines: List[str] = []
    add = lines.append
    add(f"# Projeto {ctrl.name}")
    add(f"Controlador: {ctrl.processor_type or '?'} rev. {ctrl.major_rev or '?'}.{ctrl.minor_rev or '?'} | Studio 5000 {ctrl.software_revision or '?'} | exportado em {ctrl.export_date or '?'}")
    if ctrl.description:
        add(f"Descrição: {_s(ctrl.description, 300)}")
    if ctrl.safety_controller or metrics.safety_programs:
        add("AVISO: projeto contém conteúdo de segurança (GuardLogix). Conteúdo safety é somente leitura.")
    add("")
    add(f"Métricas: {metrics.programs} programas, {metrics.routines} rotinas ({', '.join(f'{k}: {v}' for k, v in sorted(metrics.routines_by_type.items()))}), "
        f"{metrics.rungs} rungs, {metrics.instructions} instruções, {metrics.tags_controller} tags de controlador, {metrics.tags_program} tags de programa, "
        f"{metrics.aois} AOIs ({metrics.aois_protected} protegidas), {metrics.udts} UDTs, {metrics.modules} módulos. "
        f"Maior rotina: {metrics.largest_routine or '-'} ({metrics.largest_routine_rungs} rungs). Profundidade máx. de ramos: {metrics.max_branch_depth}. "
        f"Rungs comentados: {metrics.commented_rung_pct}%.")
    add("")
    add("## Tasks")
    for t in ctrl.tasks:
        rate = f" {t.rate} ms" if t.rate and t.type.upper() == "PERIODIC" else ""
        safety = " [SAFETY]" if t.is_safety else ""
        add(f"- {t.name} ({t.type}{rate}){safety}: {', '.join(t.programs) or '(sem programas)'}")
    add("")
    add("## Programas e rotinas")
    for p in ctrl.programs:
        flags = []
        if p.task is None:
            flags.append("NÃO AGENDADO")
        if p.disabled:
            flags.append("DESABILITADO")
        if p.is_safety:
            flags.append("SAFETY")
        flag_txt = f" [{', '.join(flags)}]" if flags else ""
        add(f"- Programa {p.name}{flag_txt} (main: {p.main_routine or '-'}; fault: {p.fault_routine or '-'}; {len(p.tags)} tags){(': ' + _s(p.description)) if p.description else ''}")
        for r in p.routines:
            desc = f" — {_s(r.description, 80)}" if r.description else ""
            add(f"  - {r.name} [{r.type}, {_routine_size(r)}]{desc}")
    add("")
    add("## AOIs")
    for a in ctrl.aois:
        prot = " [PROTEGIDA]" if a.protected else ""
        params = ", ".join(f"{pp.name}:{pp.usage[0]}" for pp in a.call_parameters)
        add(f"- {a.name} rev. {a.revision or '?'}{prot} ({len(a.routines)} rotinas; operandos: {params or 'só tag de instância'}){(' — ' + _s(a.description, 80)) if a.description else ''}")
    add("")
    add("## UDTs")
    for d in ctrl.data_types:
        members = ", ".join(f"{m.name}:{m.data_type}" + (f"[{m.dimension}]" if m.dimension else "") for m in d.members if not m.hidden)
        add(f"- {d.name}: {members}")
    add("")
    add("## Módulos")
    for m in ctrl.modules:
        add(f"- {m.name} ({m.catalog or '?'}, pai: {m.parent or '-'}){(' — ' + _s(m.description, 60)) if m.description else ''}")
    add("")
    add("## Resumo dos findings (análise determinística)")
    by_sev = metrics.findings_by_severity
    add(", ".join(f"{s}: {by_sev.get(s, 0)}" for s in SEVERITIES))
    by_rule: Dict[str, int] = {}
    for f in findings:
        by_rule[f.rule] = by_rule.get(f.rule, 0) + 1
    for rule, n in sorted(by_rule.items(), key=lambda kv: -kv[1]):
        add(f"- {rule}: {n}")
    add("Principais (error/warning):")
    shown = 0
    for f in findings:
        if f.severity == "info":
            continue
        add(f"- [{f.severity}] {f.location()}: {f.message}")
        shown += 1
        if shown >= 40:
            add("- …")
            break

    text = "\n".join(lines)
    truncated = False
    if estimate_tokens(text) > max_tokens:
        truncated = True
        text = _shrink_overview(lines, max_tokens)
    return Serialized(text=text, tokens=estimate_tokens(text), truncated=truncated)


def _shrink_overview(lines: List[str], max_tokens: int) -> str:
    """Drop routine detail lines, then UDT/module detail, until it fits."""
    limit = int(max_tokens * CHARS_PER_TOKEN)

    def drop(prefix_section: str, keep: int) -> None:
        nonlocal lines
        out: List[str] = []
        in_section = False
        count = 0
        for ln in lines:
            if ln.startswith("## "):
                in_section = ln == prefix_section
                count = 0
                out.append(ln)
                continue
            if in_section and ln.startswith("- Programa "):
                count = 0  # routine budget is per program
            if in_section and ln.startswith("  - "):
                count += 1
                if count == keep + 1:
                    out.append("  - … (ver rotinas via dump)")
                if count > keep:
                    continue
            elif in_section and ln.startswith("- ") and prefix_section in ("## UDTs", "## Módulos", "## AOIs"):
                count += 1
                if count == keep + 1:
                    out.append("- …")
                if count > keep:
                    continue
            out.append(ln)
        lines = out

    for section, keep in (("## Programas e rotinas", 12), ("## UDTs", 20), ("## Módulos", 20), ("## Programas e rotinas", 4), ("## AOIs", 20), ("## Programas e rotinas", 0), ("## UDTs", 0), ("## Módulos", 0)):
        text = "\n".join(lines)
        if len(text) <= limit:
            return text
        drop(section, keep)
    return "\n".join(lines)[:limit]


# --------------------------------------------------------------------------
# Routine text
# --------------------------------------------------------------------------
def routine_text(ctrl: Controller, program: str, routine: str, xref: Optional[CrossRef] = None) -> Serialized:
    prog = ctrl.program(program)
    if prog is None:
        raise KeyError(f"programa '{program}' não encontrado")
    rt = prog.routine(routine)
    if rt is None:
        raise KeyError(f"rotina '{routine}' não encontrada no programa '{program}'")
    xref = xref or build_xref(ctrl)
    lines: List[str] = []
    add = lines.append
    header = f"# {prog.name} / {rt.name} [{rt.type}]"
    if prog.is_safety:
        header += " [SAFETY — somente leitura]"
    add(header)
    if rt.description:
        add(f"Descrição: {_s(rt.description, 400)}")
    if rt.protected:
        add("Rotina protegida (source protection): conteúdo não disponível.")
        text = "\n".join(lines)
        return Serialized(text=text, tokens=estimate_tokens(text))
    if rt.calls:
        add(f"Chama: {', '.join(dict.fromkeys(rt.calls))}")
    callers = xref.callers_of(prog.name, rt.name)
    if callers:
        add(f"Chamada por: {', '.join(f'{p}/{r}' for p, r in callers)}")
    add("")
    used: Dict[str, Tag] = {}
    io_refs: List[str] = []

    def note_tag(base: Optional[str], program_ref: Optional[str]) -> None:
        if not base:
            return
        if ":" in base:
            io_refs.append(base)
            return
        tag = xref.resolve(prog.name, base, program_ref)
        if tag is not None:
            used.setdefault(f"{tag.scope}|{tag.name}".lower(), tag)

    if rt.type == "RLL":
        for rung in rt.rungs:
            if rung.comment:
                for c in rung.comment.strip().splitlines():
                    add(f"  // {c.rstrip()}")
            add(f"Rung {rung.number}: {' '.join(rung.text.split())}")
            if rung.parse_error:
                add(f"  (não interpretado: {rung.parse_error})")
            for ins in rung.instructions:
                for op in ins.operands:
                    if op.kind == "tag":
                        note_tag(op.tag_base, op.program)
                    for extra in op.extra_tags:
                        try:
                            from .rung_text import split_tag_path

                            p_, b_, _ = split_tag_path(extra)
                            note_tag(b_, p_)
                        except ValueError:
                            pass
    elif rt.type == "ST":
        for ln in rt.st_lines:
            add(f"L{ln.number}: {ln.text.rstrip()}")
        for ref in rt.tag_refs:
            note_tag(ref.tag_base, ref.program)
    else:
        add(f"Rotina {rt.type}: {rt.sheets} folha(s), {rt.size} elementos gráficos. Conteúdo não analisado em detalhe no MVP; tags referenciadas (aproximado):")
        for ref in rt.tag_refs:
            note_tag(ref.tag_base, ref.program)
    add("")
    add("## Tags usadas nesta rotina")
    for tag in sorted(used.values(), key=lambda t: t.name.lower()):
        scope = "ctrl" if tag.scope == "controller" else tag.scope
        alias = f" alias de {tag.alias_for}" if tag.alias_for else ""
        reads = len(xref.reads_of(tag))
        writes = len(xref.writes_of(tag))
        desc = f" — {_s(tag.description, 100)}" if tag.description else ""
        add(f"- {tag.name} ({tag.data_type}, {scope}{alias}; L{reads}/E{writes}){desc}")
    for io in sorted(set(io_refs)):
        add(f"- {io} (I/O de módulo)")
    text = "\n".join(lines)
    return Serialized(text=text, tokens=estimate_tokens(text))


# --------------------------------------------------------------------------
# Tag sheet
# --------------------------------------------------------------------------
def tag_sheet(ctrl: Controller, xref: Optional[CrossRef] = None, scope: Optional[str] = None, max_rows: Optional[int] = None) -> Serialized:
    """Table: tag | scope | type | reads | writes | description.

    ``scope``: None = all; "controller" or a program name.
    """
    xref = xref or build_xref(ctrl)
    rows: List[Tuple[str, str, str, int, int, str]] = []
    scopes: Iterable[Tuple[str, List[Tag]]] = [("controller", ctrl.tags)] + [(p.name, p.tags) for p in ctrl.programs]
    for sc, tags in scopes:
        if scope and sc.lower() != scope.lower():
            continue
        for t in tags:
            rows.append((t.name, sc, t.data_type + (f"[{t.dimensions}]" if t.dimensions else ""), len(xref.reads_of(t)), len(xref.writes_of(t)), _s(t.description, 80)))
    truncated = False
    if max_rows is not None and len(rows) > max_rows:
        rows = rows[:max_rows]
        truncated = True
    lines = ["tag | escopo | tipo | leituras | escritas | descrição", "--- | --- | --- | --- | --- | ---"]
    for r in rows:
        lines.append(" | ".join(str(c) for c in r))
    if truncated:
        lines.append("… (tabela truncada)")
    text = "\n".join(lines)
    return Serialized(text=text, tokens=estimate_tokens(text), truncated=truncated)


# --------------------------------------------------------------------------
# JSON
# --------------------------------------------------------------------------
def to_json(ctrl: Controller, findings: List[Finding], metrics: Metrics, include_rungs: bool = True) -> str:
    data = ctrl.to_dict()
    if not include_rungs:
        for p in data["programs"]:
            for r in p["routines"]:
                r["rungs"] = len(r["rungs"])
                r["st_lines"] = len(r["st_lines"])
                r["tag_refs"] = len(r["tag_refs"])
        for a in data["aois"]:
            for r in a["routines"]:
                r["rungs"] = len(r["rungs"])
                r["st_lines"] = len(r["st_lines"])
                r["tag_refs"] = len(r["tag_refs"])
    payload = {"controller": data, "findings": [asdict(f) for f in findings], "metrics": asdict(metrics)}
    return json.dumps(payload, ensure_ascii=False, indent=1)

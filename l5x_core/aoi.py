"""Add-On Instruction (AOI) analysis: deterministic rules, a structured report
per AOI and a Markdown document (documentation + debug guide) built only from
the parsed model and the cross references. Nothing here needs an LLM.

Rules (Finding.rule, all with program="aoi:<Name>"):
  aoi_unused, aoi_instance_never_called, aoi_call_arity, aoi_param_unused,
  aoi_output_never_written, aoi_input_written, aoi_local_unused,
  aoi_local_write_only, aoi_local_read_only, aoi_output_multiple_writes,
  aoi_no_enableinfalse, aoi_no_description, aoi_params_no_description,
  aoi_unresolved_tags, aoi_nested_protected.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .model import Aoi, Controller, Finding, Routine, Tag, SEVERITIES
from .rung_text import TIMER_COUNTER
from .xref import AOI_PREFIX, AoiCallSite, CrossRef, Usage, build_xref

INIT_ROUTINES = {"prescan", "postscan"}
_SEV_ORDER = {s: i for i, s in enumerate(SEVERITIES)}


def aoi_scope(name: str) -> str:
    return f"{AOI_PREFIX}{name}"


def _join(items: Iterable[str], limit: int = 8) -> str:
    items = list(dict.fromkeys(items))
    if len(items) <= limit:
        return ", ".join(items)
    return ", ".join(items[:limit]) + f" … (+{len(items) - limit})"


def _short_where(u: Usage) -> str:
    """'Logic rung 3' instead of 'aoi:X / Logic / rung 3' (used inside an AOI report)."""
    return f"{u.routine} {u.location}"


def format_value(value: Any, limit: int = 60) -> str:
    """Compact rendering of a decorated default value."""
    if value is None:
        return ""
    if isinstance(value, dict):
        inner = ", ".join(f"{k}={format_value(v, 20)}" for k, v in list(value.items())[:6])
        if len(value) > 6:
            inner += ", …"
        text = "{" + inner + "}"
    else:
        text = str(value)
    return text if len(text) <= limit else text[: limit - 1] + "…"


# --------------------------------------------------------------------------
# Report structures
# --------------------------------------------------------------------------
@dataclass
class TagStats:
    """How a parameter or local tag is used inside the AOI routines."""

    name: str
    reads: int = 0
    writes: int = 0
    where_reads: List[str] = field(default_factory=list)
    where_writes: List[str] = field(default_factory=list)
    write_instructions: List[str] = field(default_factory=list)

    @property
    def used(self) -> bool:
        return self.reads + self.writes > 0


@dataclass
class InstanceInfo:
    tag: Tag
    calls: List[AoiCallSite] = field(default_factory=list)
    member_reads: Dict[str, List[str]] = field(default_factory=dict)  # member path -> where (outside the AOI)
    member_writes: Dict[str, List[str]] = field(default_factory=dict)
    other_usages: int = 0  # references to the instance that are not calls nor member access

    @property
    def scope_label(self) -> str:
        return "controlador" if self.tag.scope == "controller" else self.tag.scope


@dataclass
class AoiReport:
    aoi: Aoi
    scope: str
    instances: List[InstanceInfo]
    calls: List[AoiCallSite]
    params: Dict[str, TagStats]  # lower-cased parameter name -> stats
    locals: Dict[str, TagStats]  # lower-cased local tag name -> stats
    nested: List[str]  # AOIs called inside this one
    used_by: List[str]  # AOIs that call this one
    findings: List[Finding]
    rungs: int = 0
    commented_rungs: int = 0
    st_lines: int = 0
    unresolved: List[str] = field(default_factory=list)

    @property
    def protected(self) -> bool:
        return self.aoi.protected

    def findings_by_severity(self) -> Dict[str, int]:
        out = {s: 0 for s in SEVERITIES}
        for f in self.findings:
            out[f.severity] = out.get(f.severity, 0) + 1
        return out

    def calls_by_source(self) -> Dict[str, int]:
        out: Dict[str, int] = defaultdict(int)
        for c in self.calls:
            out[c.call.source] += 1
        return dict(out)


# --------------------------------------------------------------------------
# Statistics
# --------------------------------------------------------------------------
def _tag_stats(aoi: Aoi, xref: CrossRef) -> Tuple[Dict[str, TagStats], Dict[str, TagStats]]:
    params: Dict[str, TagStats] = {}
    locals_: Dict[str, TagStats] = {}
    for tag in aoi.local_tags:
        st = TagStats(name=tag.name)
        for u in xref.usages_of(tag):
            if u.reads:
                st.reads += 1
                st.where_reads.append(_short_where(u))
            if u.writes:
                st.writes += 1
                st.where_writes.append(_short_where(u))
                st.write_instructions.append(u.instruction)
        if tag.usage:
            params[tag.name.lower()] = st
        else:
            locals_[tag.name.lower()] = st
    return params, locals_


def _instance_info(tag: Tag, aoi: Aoi, xref: CrossRef) -> InstanceInfo:
    info = InstanceInfo(tag=tag, calls=xref.calls_of_instance(tag))
    base = tag.name.lower()
    aoi_low = aoi.name.lower()
    for u in xref.usages_of(tag):
        if u.instruction.lower() == aoi_low:
            continue  # the call itself
        path = u.path
        if path.lower() == base or path.lower().startswith(base + "["):
            info.other_usages += 1
            continue
        member = path[len(tag.name):] if path.lower().startswith(base) else path
        member = member.lstrip(".")
        if u.writes:
            info.member_writes.setdefault(member, []).append(u.where())
        else:
            info.member_reads.setdefault(member, []).append(u.where())
    return info


def instance_counts(ctrl: Controller) -> Dict[str, int]:
    """Lower-cased AOI name -> number of instance tags (no xref needed; used by the overview)."""
    names = {a.name.lower() for a in ctrl.aois}
    counts: Dict[str, int] = defaultdict(int)
    if not names:
        return {}

    def scan(tags: Iterable[Tag], exclude_params: bool) -> None:
        for t in tags:
            dt = t.data_type.lower()
            if dt in names and not (exclude_params and (t.usage or "").lower() in ("input", "output", "inout")):
                counts[dt] += 1

    scan(ctrl.tags, False)
    for p in ctrl.programs:
        scan(p.tags, True)
    for a in ctrl.aois:
        scan(a.local_tags, True)
    return dict(counts)


# --------------------------------------------------------------------------
# Rules
# --------------------------------------------------------------------------
def rules_for_aoi(ctrl: Controller, aoi: Aoi, xref: CrossRef, params: Optional[Dict[str, TagStats]] = None, locals_: Optional[Dict[str, TagStats]] = None) -> List[Finding]:
    scope = aoi_scope(aoi.name)
    safety = aoi.is_safety
    out: List[Finding] = []
    instances = xref.instances_of(aoi.name)
    calls = xref.calls_of(aoi.name)
    label = f"AOI '{aoi.name}' (rev. {aoi.revision or '?'})"

    def add(rule: str, severity: str, message: str, **kw: Any) -> None:
        out.append(Finding(rule=rule, severity=severity, message=message, program=scope, safety=safety, **kw))

    # ---- usage of the AOI itself ------------------------------------------
    if not instances and not calls:
        add("aoi_unused", "info", f"{label} definida mas sem instâncias (nenhuma tag do tipo {aoi.name}) e sem chamadas.")
    for inst in instances:
        if xref.calls_of_instance(inst):
            continue
        scope_label = "controlador" if inst.scope == "controller" else inst.scope
        uses = [u for u in xref.usages_of(inst)]
        if uses:
            # an instance with no usage at all is already reported by tag_unused / aoi_local_unused
            add("aoi_instance_never_called", "info", f"Instância '{inst.name}' ({aoi.name}, escopo {scope_label}) é referenciada mas nunca chamada diretamente (passada como parâmetro ou só lida/escrita por membros).",
                tag=inst.name, evidence=_join([u.where() for u in uses]))

    # ---- call sites ----------------------------------------------------------
    if not aoi.protected:
        expected = len(aoi.call_parameters)
        for site in calls:
            c = site.call
            if c.source not in ("rll", "st"):
                continue
            if len(c.arguments) != expected:
                names = ", ".join(p.name for p in aoi.call_parameters) or "(só a tag de instância)"
                out.append(Finding(
                    rule="aoi_call_arity", severity="error", program=site.scope, routine=site.routine, rung=c.rung, tag=c.instance,
                    message=f"Chamada de {aoi.name} com {len(c.arguments)} operando(s) além da instância; a definição espera {expected} ({names}).",
                    evidence=f"{c.aoi}({', '.join([c.instance or '?'] + c.arguments)})", safety=safety,
                ))
    for r in aoi.routines:
        for c in r.aoi_calls:
            callee = ctrl.aoi(c.aoi)
            if callee is not None and callee.protected:
                add("aoi_nested_protected", "info", f"{label} chama a AOI protegida '{callee.name}' ({r.name} {c.location}): o comportamento interno dessa chamada não é analisável.", routine=r.name, rung=c.rung, tag=c.instance)

    if aoi.protected:
        return out

    # ---- parameters -------------------------------------------------------------
    if params is None or locals_ is None:
        params, locals_ = _tag_stats(aoi, xref)
    graphic = any(r.type in ("FBD", "SFC") for r in aoi.routines)
    approx = " (referências de FBD/SFC são aproximadas)" if graphic else ""
    if aoi.analyzable:
        for p in aoi.user_parameters:
            st = params.get(p.name.lower())
            if st is None:
                continue
            on_call = p.required or p.usage == "InOut"
            if not st.used:
                if on_call:
                    add("aoi_param_unused", "warning", f"Parâmetro '{p.name}' ({p.usage}, {p.data_type}) é operando da chamada mas não é referenciado em nenhuma rotina da {label}{approx}.", tag=p.name)
                else:
                    add("aoi_param_unused", "info", f"Parâmetro '{p.name}' ({p.usage}, {p.data_type}) não é referenciado em nenhuma rotina da {label}: só faz sentido se for lido/escrito pelo HMI na tag de instância{approx}.", tag=p.name)
                continue
            if p.usage == "Output" and st.writes == 0:
                add("aoi_output_never_written", "warning", f"Parâmetro de saída '{p.name}' nunca é escrito dentro da {label}: a saída fica sempre no valor default ({format_value(p.default) or '0'}){approx}.",
                    tag=p.name, evidence="lido em " + _join(st.where_reads))
            if p.usage == "Input" and p.required:
                # non-required inputs live in the instance tag and may legitimately be adjusted by the logic
                logic_writes = [w for w, inst in zip(st.where_writes, st.write_instructions) if w.split(" ", 1)[0].lower() not in INIT_ROUTINES]
                if logic_writes:
                    add("aoi_input_written", "warning", f"Parâmetro de entrada obrigatório '{p.name}' é escrito dentro da lógica da {label}: o valor é sobrescrito pelo argumento a cada chamada (escrita perdida ou efeito só dentro da varredura).",
                        tag=p.name, evidence=_join(logic_writes))
        for tag in aoi.local_tags:
            if tag.usage:
                continue
            st = locals_.get(tag.name.lower())
            if st is None:
                continue
            default = format_value(tag.values)
            if not st.used:
                hmi = " (exposta ao HMI: ExternalAccess=" + tag.external_access + ")" if tag.hmi_readable else ""
                add("aoi_local_unused", "info", f"Tag local '{tag.name}' ({tag.data_type}) nunca referenciada na {label}{hmi}{approx}.", tag=tag.name)
            elif st.writes and not st.reads and not tag.hmi_readable:
                # written-only locals visible to the HMI are status words, not dead writes
                add("aoi_local_write_only", "info", f"Tag local '{tag.name}' ({tag.data_type}) é escrita mas nunca lida na {label} e não é visível ao HMI (ExternalAccess=None){approx}.", tag=tag.name, evidence=_join(st.where_writes))
            elif st.reads and not st.writes and not tag.hmi_writable:
                hint = f" (default: {default})" if default else ""
                add("aoi_local_read_only", "info", f"Tag local '{tag.name}' ({tag.data_type}) é lida mas nunca escrita na {label} nem editável pelo HMI: constante de configuração ou valor esquecido?{hint}{approx}", tag=tag.name, evidence=_join(st.where_reads))

    # ---- OTE on the same operand in several rungs of the same routine -------------------
    for r in aoi.routines:
        ote: Dict[str, List[int]] = defaultdict(list)
        display: Dict[str, str] = {}
        for rung in r.rungs:
            for ins in rung.instructions:
                if ins.name.upper() == "OTE" and ins.operands and ins.operands[0].kind == "tag":
                    path = ins.operands[0].tag_path or ins.operands[0].raw
                    ote[path.lower()].append(rung.number)
                    display.setdefault(path.lower(), path)
        for key, rungs in ote.items():
            distinct = sorted(set(rungs))
            if len(distinct) > 1:
                add("aoi_output_multiple_writes", "warning", f"'{display[key]}' é escrita por OTE em {len(distinct)} rungs da rotina {r.name} da {label} (a última na varredura vence).",
                    routine=r.name, rung=distinct[0], tag=display[key], evidence="rungs " + ", ".join(str(n) for n in distinct))

    # ---- behaviour / documentation ---------------------------------------------
    outputs = [p for p in aoi.user_parameters if p.usage == "Output"]
    if outputs and not aoi.execute_enableinfalse:
        add("aoi_no_enableinfalse", "info", f"{label} tem {len(outputs)} saída(s) e não executa EnableInFalse: com o rung/EnableIn falso as saídas mantêm o último valor (verifique se isso é o esperado).",
            evidence=_join([p.name for p in outputs]))
    if not aoi.description:
        add("aoi_no_description", "info", f"{label} sem descrição.")
    missing = [p.name for p in aoi.user_parameters if not p.description]
    if missing:
        add("aoi_params_no_description", "info", f"{len(missing)} de {len(aoi.user_parameters)} parâmetro(s) da {label} sem descrição.", evidence=_join(missing, 12))
    for (sc, routine), names in xref.unresolved.items():
        if sc != scope:
            continue
        uniq = sorted(set(names), key=str.lower)
        add("aoi_unresolved_tags", "info", f"{len(uniq)} referência(s) na rotina {routine} da {label} a tags que não são parâmetro nem tag local (AOI só enxerga o próprio escopo).",
            routine=routine, evidence=_join(uniq, 10))
    return out


def aoi_findings(ctrl: Controller, xref: CrossRef) -> List[Finding]:
    out: List[Finding] = []
    for aoi in ctrl.aois:
        out += rules_for_aoi(ctrl, aoi, xref)
    return out


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------
def build_aoi_report(ctrl: Controller, name: str, xref: Optional[CrossRef] = None, findings: Optional[List[Finding]] = None) -> AoiReport:
    """Everything known about one AOI. ``findings``: pass the project findings to
    reuse them; otherwise the AOI rules run for this AOI only."""
    aoi = ctrl.aoi(name)
    if aoi is None:
        raise KeyError(f"AOI '{name}' não encontrada")
    xref = xref or build_xref(ctrl)
    scope = aoi_scope(aoi.name)
    params, locals_ = _tag_stats(aoi, xref)
    if findings is None:
        own = rules_for_aoi(ctrl, aoi, xref, params, locals_)
        own += [f for f in ctrl.parse_findings if f.program is None and any(f.routine == r.name for r in aoi.routines)]  # rung parse errors inside the AOI
    else:
        inst_names = {t.name.lower() for t in xref.instances_of(aoi.name)}
        own = [
            f for f in findings
            if f.program == scope
            or (f.rule == "aoi_protected" and f.tag == aoi.name)
            or (f.rule in ("aoi_call_arity", "tag_unused", "aoi_local_unused") and (f.tag or "").split("[")[0].lower() in inst_names)
        ]
    own = sorted(own, key=lambda f: (_SEV_ORDER.get(f.severity, 9), f.routine or "", f.rung if f.rung is not None else -1))
    instances = [_instance_info(t, aoi, xref) for t in xref.instances_of(aoi.name)]
    calls = xref.calls_of(aoi.name)
    nested: List[str] = []
    for r in aoi.routines:
        for c in r.aoi_calls:
            callee = ctrl.aoi(c.aoi)
            nested.append(callee.name if callee else c.aoi)
    used_by = [s.scope[len(AOI_PREFIX):] for s in calls if s.scope.startswith(AOI_PREFIX)]
    rungs = sum(len(r.rungs) for r in aoi.routines)
    commented = sum(1 for r in aoi.routines for g in r.rungs if g.comment)
    st_lines = sum(len(r.st_lines) for r in aoi.routines)
    unresolved = sorted({n for (sc, _), names in xref.unresolved.items() if sc == scope for n in names}, key=str.lower)
    return AoiReport(
        aoi=aoi, scope=scope, instances=instances, calls=calls, params=params, locals=locals_,
        nested=list(dict.fromkeys(nested)), used_by=list(dict.fromkeys(used_by)), findings=own,
        rungs=rungs, commented_rungs=commented, st_lines=st_lines, unresolved=unresolved,
    )


def aoi_summary_rows(ctrl: Controller, xref: Optional[CrossRef] = None, findings: Optional[List[Finding]] = None) -> List[Dict[str, Any]]:
    """One row per AOI for tables (CLI / UI)."""
    xref = xref or build_xref(ctrl)
    by_aoi: Dict[str, Dict[str, int]] = defaultdict(lambda: {s: 0 for s in SEVERITIES})
    if findings is not None:
        for f in findings:
            if f.program and f.program.startswith(AOI_PREFIX):
                by_aoi[f.program[len(AOI_PREFIX):].lower()][f.severity] += 1
    rows: List[Dict[str, Any]] = []
    for a in ctrl.aois:
        calls = xref.calls_of(a.name)
        rows.append({
            "aoi": a.name, "revision": a.revision or "", "vendor": a.vendor or "", "protected": a.protected,
            "parameters": len(a.user_parameters), "locals": sum(1 for t in a.local_tags if not t.usage),
            "routines": ", ".join(f"{r.name}[{r.type}]" for r in a.routines),
            "instances": len(xref.instances_of(a.name)), "calls": len(calls),
            "nested": sorted({c.aoi for r in a.routines for c in r.aoi_calls}, key=str.lower),
            "findings": dict(by_aoi.get(a.name.lower(), {s: 0 for s in SEVERITIES})),
            "description": a.description or "",
        })
    return rows


# --------------------------------------------------------------------------
# Markdown: documentation + debug guide (deterministic, no LLM)
# --------------------------------------------------------------------------
def _md_cell(text: Optional[str], limit: int = 80) -> str:
    if not text:
        return ""
    one = " ".join(str(text).split()).replace("|", "\\|")
    return one if len(one) <= limit else one[: limit - 1] + "…"


def _call_args_text(site: AoiCallSite, aoi: Aoi) -> str:
    c = site.call
    if c.source == "fbd":
        if not c.pins:
            return "(sem pinos fiados visíveis)"
        return "; ".join(f"{k}={v}" for k, v in c.pins.items())
    names = [p.name for p in aoi.call_parameters] if not aoi.protected else []
    parts = []
    for i, arg in enumerate(c.arguments):
        parts.append(f"{names[i]}={arg}" if i < len(names) else arg)
    return "; ".join(parts) if parts else "(só a instância)"


def _routine_md(r: Routine, max_rungs: int) -> List[str]:
    lines: List[str] = []
    if r.protected:
        lines.append("Rotina protegida (source protection): conteúdo não disponível.")
        return lines
    if r.type == "RLL":
        for i, rung in enumerate(r.rungs):
            if i >= max_rungs:
                lines.append(f"… (+{len(r.rungs) - max_rungs} rungs omitidos)")
                break
            if rung.comment:
                for c in rung.comment.strip().splitlines():
                    lines.append(f"    // {c.rstrip()}")
            lines.append(f"    Rung {rung.number}: {' '.join(rung.text.split())}")
            if rung.parse_error:
                lines.append(f"      (não interpretado: {rung.parse_error})")
    elif r.type == "ST":
        for i, ln in enumerate(r.st_lines):
            if i >= max_rungs:
                lines.append(f"… (+{len(r.st_lines) - max_rungs} linhas omitidas)")
                break
            lines.append(f"    L{ln.number}: {ln.text.rstrip()}")
    else:
        lines.append(f"Rotina {r.type}: {r.sheets} folha(s), {r.size} elementos gráficos; lógica interna não analisada. Tags referenciadas (aproximado): " + (_join(sorted({t.tag_path for t in r.tag_refs}, key=str.lower), 20) or "nenhuma"))
    return lines


def aoi_markdown(report: AoiReport, max_rows: int = 40, max_rungs: int = 80) -> str:
    """Technical document for one AOI: identification, interface, logic, instances,
    call sites, parser findings and a static debug guide. Portuguese, Markdown."""
    a = report.aoi
    L: List[str] = []
    add = L.append
    add(f"# AOI {a.name} — rev. {a.revision or '?'}" + (f" ({a.vendor})" if a.vendor else ""))
    add("")
    if a.is_safety:
        add("**Conteúdo SAFETY (Class=\"Safety\"): somente leitura; alterações exigem o processo de validação de segurança funcional.**")
        add("")
    if a.description:
        add(f"**Descrição:** {' '.join(a.description.split())}")
    if a.revision_note:
        add(f"**Nota de revisão:** {' '.join(a.revision_note.split())}")
    meta = []
    if a.created_date:
        meta.append(f"criada em {a.created_date[:10]}")
    if a.edited_date:
        meta.append(f"editada em {a.edited_date[:10]}")
    if a.software_revision:
        meta.append(f"Studio 5000 {a.software_revision}")
    if a.class_:
        meta.append(f"classe {a.class_}")
    if meta:
        add(f"**Metadados:** {'; '.join(meta)}")
    if a.protected:
        add("")
        add("**AOI protegida (source protection):** parâmetros, tags locais e rotinas não foram lidos. Só é possível listar instâncias e pontos de chamada.")
    else:
        flags = [f"Prescan: {'sim' if a.execute_prescan else 'não'}", f"Postscan: {'sim' if a.execute_postscan else 'não'}", f"EnableInFalse: {'sim' if a.execute_enableinfalse else 'não'}"]
        add(f"**Rotinas de sistema executadas:** {', '.join(flags)}")
    if a.help_text:
        add("")
        add("**Texto de ajuda (AdditionalHelpText):**")
        add("")
        add("> " + " ".join(a.help_text.split())[:800])
    add("")

    # ---- summary -------------------------------------------------------------------
    add("## Resumo")
    add("")
    user_params = a.user_parameters
    by_usage = {u: sum(1 for p in user_params if p.usage == u) for u in ("Input", "Output", "InOut")}
    n_locals = sum(1 for t in a.local_tags if not t.usage)
    by_scope: Dict[str, int] = defaultdict(int)
    for inst in report.instances:
        by_scope[inst.scope_label] += 1
    src = report.calls_by_source()
    sev = report.findings_by_severity()
    if not a.protected:
        add(f"- Interface: {len(user_params)} parâmetro(s) (Input {by_usage['Input']}, Output {by_usage['Output']}, InOut {by_usage['InOut']}); operandos na chamada: {', '.join(p.name for p in a.call_parameters) or 'só a tag de instância'}.")
        add(f"- Tags locais: {n_locals}. Rotinas: " + (", ".join(f"{r.name} [{r.type}, {_routine_size(r)}]" for r in a.routines) or "nenhuma") + ".")
        if report.rungs:
            pct = 100.0 * report.commented_rungs / report.rungs
            add(f"- Rungs: {report.rungs} ({pct:.0f}% comentados)" + (f"; linhas ST: {report.st_lines}" if report.st_lines else "") + ".")
    add(f"- Instâncias: {len(report.instances)}" + (" (" + ", ".join(f"{k}: {v}" for k, v in sorted(by_scope.items())) + ")" if by_scope else "") + ".")
    add(f"- Pontos de chamada: {len(report.calls)}" + (" (" + ", ".join(f"{k.upper()}: {v}" for k, v in sorted(src.items())) + ")" if src else "") + ".")
    if report.nested:
        add(f"- Chama outras AOIs: {', '.join(report.nested)}.")
    if report.used_by:
        add(f"- Usada dentro de outras AOIs: {', '.join(report.used_by)}.")
    add(f"- Findings do parser: " + ", ".join(f"{s}: {sev.get(s, 0)}" for s in SEVERITIES) + ".")
    add("")

    # ---- interface ---------------------------------------------------------------------
    if not a.protected:
        add("## Interface (parâmetros)")
        add("")
        add("| Parâmetro | Uso | Tipo | Req. | Vis. | Default | Uso interno (L/E) | Descrição |")
        add("| --- | --- | --- | --- | --- | --- | --- | --- |")
        for p in a.parameters:
            st = report.params.get(p.name.lower())
            use = f"L{st.reads}/E{st.writes}" if st else "-"
            add(f"| {p.name} | {p.usage} | {p.data_type or '-'} | {'sim' if p.required else 'não'} | {'sim' if p.visible else 'não'} | {_md_cell(format_value(p.default), 30)} | {use} | {_md_cell(p.description)} |")
        add("")
        add("Req. = aparece como operando na chamada (Required ou InOut). L/E = leituras/escritas dentro das rotinas da AOI (parser).")
        add("")
        add("## Tags locais")
        add("")
        locals_ = [t for t in a.local_tags if not t.usage]
        if locals_:
            add("| Tag | Tipo | Default | Uso interno (L/E) | Descrição |")
            add("| --- | --- | --- | --- | --- |")
            for i, t in enumerate(locals_):
                if i >= max_rows:
                    add(f"| … (+{len(locals_) - max_rows}) | | | | |")
                    break
                st = report.locals.get(t.name.lower())
                use = f"L{st.reads}/E{st.writes}" if st else "-"
                dims = f"[{t.dimensions}]" if t.dimensions else ""
                add(f"| {t.name} | {t.data_type}{dims} | {_md_cell(format_value(t.values), 40)} | {use} | {_md_cell(t.description)} |")
        else:
            add("Nenhuma tag local.")
        add("")

        # ---- logic -----------------------------------------------------------------------
        add("## Lógica interna")
        add("")
        for r in a.routines:
            add(f"### {r.name} [{r.type}, {_routine_size(r)}]")
            if r.description:
                add(f"_{' '.join(r.description.split())}_")
            add("")
            L.extend(_routine_md(r, max_rungs))
            add("")

    # ---- instances --------------------------------------------------------------------
    add("## Instâncias")
    add("")
    if report.instances:
        add("| Instância | Escopo | Tipo | Chamadas | Membros lidos fora | Membros escritos fora | Descrição |")
        add("| --- | --- | --- | --- | --- | --- | --- |")
        for i, inst in enumerate(report.instances):
            if i >= max_rows:
                add(f"| … (+{len(report.instances) - max_rows} instâncias) | | | | | | |")
                break
            t = inst.tag
            dims = f"[{t.dimensions}]" if t.dimensions else ""
            where = _join([f"{s.scope}/{s.routine} {s.call.location}" for s in inst.calls], 4) or "**nenhuma**"
            add(f"| {t.name} | {inst.scope_label} | {t.data_type}{dims} | {_md_cell(where, 120)} | {_md_cell(_join(sorted(inst.member_reads), 6), 100)} | {_md_cell(_join(sorted(inst.member_writes), 6), 100)} | {_md_cell(t.description)} |")
    else:
        add("Nenhuma tag do tipo desta AOI foi declarada no projeto.")
    add("")

    # ---- call sites ----------------------------------------------------------------------
    add("## Pontos de chamada")
    add("")
    if report.calls:
        add("| Onde | Linguagem | Instância | Argumentos / pinos |")
        add("| --- | --- | --- | --- |")
        for i, s in enumerate(report.calls):
            if i >= max_rows:
                add(f"| … (+{len(report.calls) - max_rows} chamadas) | | | |")
                break
            add(f"| {s.scope} / {s.routine} / {s.call.location} | {s.call.source.upper()} | {s.call.instance or '?'} | {_md_cell(_call_args_text(s, a), 200)} |")
        add("")
        add("RLL: operandos na ordem da definição (parâmetros Required/InOut). ST: aproximado (regex). FBD: pinos fiados = operando do IRef/ORef ou `bloco.pino` ligado; InOut = argumento declarado.")
    else:
        add("Nenhuma chamada encontrada em rotinas RLL, ST ou FBD.")
    add("")

    # ---- findings -------------------------------------------------------------------------
    add("## Findings do parser")
    add("")
    if report.findings:
        for f in report.findings:
            loc = " / ".join(x for x in (f.routine, f"rung {f.rung}" if f.rung is not None else None) if x)
            ev = f" — evidência: {f.evidence}" if f.evidence else ""
            add(f"- [{f.severity}] `{f.rule}`" + (f" ({loc})" if loc else "") + f": {f.message}{ev}")
    else:
        add("Nenhum finding para esta AOI.")
    add("")

    # ---- debug guide ---------------------------------------------------------------------
    add("## Roteiro de debug")
    add("")
    if a.protected:
        add("- AOI protegida: não há lógica para inspecionar. Verifique as instâncias e os pontos de chamada acima e compare os valores dos parâmetros em um instância online no Studio 5000.")
    else:
        outputs = [p for p in user_params if p.usage == "Output"]
        inputs = [p for p in user_params if p.usage == "Input"]
        inouts = [p for p in user_params if p.usage == "InOut"]
        if outputs:
            add("**Saídas: onde cada uma é decidida**")
            add("")
            for p in outputs:
                st = report.params.get(p.name.lower())
                if st and st.writes:
                    instr = _join(sorted(set(st.write_instructions)), 4)
                    add(f"- `{p.name}` escrita em {_join(st.where_writes, 6)} ({instr}). Para saber por que está em um estado, inspecione esses rungs na instância online.")
                else:
                    add(f"- `{p.name}` **nunca é escrita** na lógica: fica no default {format_value(p.default) or '0'}.")
            add("")
        if inputs or inouts:
            add("**Entradas: o que a lógica lê**")
            add("")
            for p in inputs + inouts:
                st = report.params.get(p.name.lower())
                where = _join(st.where_reads, 5) if st and st.reads else "nunca lida"
                add(f"- `{p.name}` ({p.usage}, {p.data_type or '?'}): {where}." + (" Argumento na chamada." if p.required or p.usage == "InOut" else " Valor vive na tag de instância (preencha pela chamada em FBD/ST ou pelo HMI)."))
            add("")
        timers = _timers(a)
        if timers:
            add("**Timers e contadores internos**")
            add("")
            for name, instr, preset, where in timers:
                add(f"- `{name}` ({instr}) preset {preset}, em {where}.")
            add("")
        add("**Comportamento de execução**")
        add("")
        if a.execute_enableinfalse:
            add("- Com EnableIn falso (rung de chamada falso) a rotina EnableInFalse executa: veja o que ela faz com as saídas acima.")
        else:
            add("- Com EnableIn falso (rung de chamada falso) a lógica NÃO executa e as saídas mantêm o último valor.")
        if a.execute_prescan:
            add("- Prescan executa na transição para Run: valores iniciais vêm dessa rotina.")
        else:
            add("- Sem Prescan: valores iniciais são os defaults da definição ou o que estava na tag de instância.")
        if report.nested:
            add(f"- AOIs aninhadas ({', '.join(report.nested)}): o estado dessas instâncias internas fica nas tags locais correspondentes.")
        if report.unresolved:
            add(f"- Referências não resolvidas dentro da AOI: {_join(report.unresolved, 10)} (verifique se são parâmetros removidos).")
        add("")
        add("**Checklist por instância**")
        add("")
        add("1. Localize a instância na tabela acima e abra o ponto de chamada (programa/rotina/rung ou folha).")
        add("2. Confira os argumentos/pinos: cada parâmetro Required/InOut precisa de uma tag válida; pinos de entrada não fiados em FBD ficam no último valor escrito.")
        add("3. Compare os membros lidos fora da AOI (coluna 'Membros lidos fora') com as saídas escritas na lógica: um membro lido fora e nunca escrito dentro indica saída morta.")
        add("4. Se a saída não muda, verifique primeiro EnableIn/condição do rung e depois os rungs listados em 'Saídas'.")
    add("")
    add("---")
    add("_Relatório determinístico gerado pelo parser (sem LLM). Validar com o engenheiro responsável antes de qualquer alteração._")
    return "\n".join(L)


def _routine_size(r: Routine) -> str:
    if r.protected:
        return "protegida"
    if r.type == "RLL":
        return f"{len(r.rungs)} rungs"
    if r.type == "ST":
        return f"{len(r.st_lines)} linhas"
    if r.type == "FBD":
        return f"{r.sheets} folha(s)"
    return "não analisada"


def _timers(aoi: Aoi) -> List[Tuple[str, str, str, str]]:
    """(tag, instruction, preset, where) for TON/TOF/RTO/CTU/CTD inside the AOI routines."""
    locals_ = {t.name.lower(): t for t in aoi.local_tags}
    out: List[Tuple[str, str, str, str]] = []
    seen = set()
    for r in aoi.routines:
        for rung in r.rungs:
            for ins in rung.instructions:
                if ins.name.upper() not in TIMER_COUNTER or not ins.operands:
                    continue
                struct = ins.operands[0]
                name = struct.tag_path or struct.raw
                key = (name.lower(), ins.name.upper())
                preset = "?"
                if len(ins.operands) > 1:
                    op = ins.operands[1]
                    if op.kind == "literal":
                        preset = op.raw
                    elif op.kind == "tag":
                        preset = f"tag {op.tag_path}"
                    elif op.kind == "unused":
                        tag = locals_.get((struct.tag_base or "").lower())
                        if tag is not None and isinstance(tag.values, dict) and tag.values.get("PRE") is not None:
                            preset = f"{tag.values['PRE']} (default da tag local)"
                        else:
                            preset = "? (no valor da tag)"
                where = f"{r.name} rung {rung.number}"
                if key in seen:
                    for i, row in enumerate(out):
                        if (row[0].lower(), row[1]) == key:
                            out[i] = (row[0], row[1], row[2], f"{row[3]}, {where}")
                    continue
                seen.add(key)
                out.append((name, ins.name.upper(), preset, where))
    return out

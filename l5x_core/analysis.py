"""Deterministic analysis rules -> list of Finding, plus project Metrics.

Rules (ids used in Finding.rule):
  routine_never_called, program_not_scheduled, program_disabled,
  tag_unused, tag_write_only, tag_read_only, output_multiple_writes,
  output_mixed_latch, timer_no_preset, jsr_missing_routine,
  rungs_without_comment, routine_no_description, routine_protected,
  aoi_protected, safety_content, unknown_instruction, unresolved_tags,
  rung_parse_error (emitted by the parser).
"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Dict, List, Optional, Tuple

from .model import Controller, Finding, Metrics, Program, Routine, Tag, SEVERITIES
from .rung_text import TIMER_COUNTER
from .xref import CrossRef, Usage, CONTROLLER

AOI_ROUTINE_NAMES = {"logic", "prescan", "postscan", "enableinfalse"}
_SEV_ORDER = {s: i for i, s in enumerate(SEVERITIES)}


def _join(items: List[str], limit: int = 8) -> str:
    items = list(dict.fromkeys(items))  # de-duplicate, keep order
    if len(items) <= limit:
        return ", ".join(items)
    return ", ".join(items[:limit]) + f" … (+{len(items) - limit})"


# --------------------------------------------------------------------------
# Rules
# --------------------------------------------------------------------------
def rule_routine_never_called(ctrl: Controller, xref: CrossRef) -> List[Finding]:
    out: List[Finding] = []
    for prog in ctrl.programs:
        roots = {(prog.main_routine or "").lower(), (prog.fault_routine or "").lower()}
        for r in prog.routines:
            if r.name.lower() in roots:
                continue
            if xref.callers_of(prog.name, r.name):
                continue
            out.append(Finding(
                rule="routine_never_called", severity="warning", program=prog.name, routine=r.name,
                message=f"Rotina '{r.name}' não é main/fault e não é chamada por nenhum JSR no programa '{prog.name}'.",
                safety=prog.is_safety,
            ))
    return out


def rule_program_scheduling(ctrl: Controller) -> List[Finding]:
    out: List[Finding] = []
    for prog in ctrl.programs:
        if prog.task is None:
            out.append(Finding(rule="program_not_scheduled", severity="warning", program=prog.name,
                               message=f"Programa '{prog.name}' não está agendado em nenhuma task (nunca executa).", safety=prog.is_safety))
        if prog.disabled:
            out.append(Finding(rule="program_disabled", severity="warning", program=prog.name,
                               message=f"Programa '{prog.name}' está desabilitado (Disabled=\"true\").", safety=prog.is_safety))
    return out


def _iter_scoped_tags(ctrl: Controller):
    for t in ctrl.tags:
        yield None, t
    for p in ctrl.programs:
        for t in p.tags:
            yield p, t


def rule_tag_usage(ctrl: Controller, xref: CrossRef) -> List[Finding]:
    out: List[Finding] = []
    for prog, tag in _iter_scoped_tags(ctrl):
        if tag.is_io or tag.tag_type in ("Produced", "Consumed"):
            continue
        usages = xref.usages_of(tag)
        pname = prog.name if prog else None
        safety = tag.is_safety or (prog.is_safety if prog else False)
        scope_label = f"programa '{prog.name}'" if prog else "controlador"
        if not usages:
            out.append(Finding(rule="tag_unused", severity="info", program=pname, tag=tag.name,
                               message=f"Tag '{tag.name}' ({tag.data_type}) declarada no escopo {scope_label} e nunca referenciada.", safety=safety))
            continue
        if tag.tag_type == "Alias" or tag.usage:  # aliases and program parameters are excluded
            continue
        reads = [u for u in usages if u.reads]
        writes = [u for u in usages if u.writes]
        if writes and not reads:
            out.append(Finding(rule="tag_write_only", severity="warning", program=pname, tag=tag.name,
                               message=f"Tag '{tag.name}' é escrita mas nunca lida.",
                               evidence=_join([u.where() for u in writes]), safety=safety))
        elif reads and not writes and not tag.constant:
            out.append(Finding(rule="tag_read_only", severity="info", program=pname, tag=tag.name,
                               message=f"Tag '{tag.name}' é lida mas nunca escrita no programa (valor fixo, HMI ou externo?).",
                               evidence=_join([u.where() for u in reads]), safety=safety))
    return out


def rule_output_writes(ctrl: Controller, xref: CrossRef) -> List[Finding]:
    """OTE on the same bit in more than one rung; OTE mixed with OTL/OTU."""
    out: List[Finding] = []
    ote: Dict[Tuple[str, str], List[Tuple[str, str, int]]] = defaultdict(list)  # (tagkey, path) -> [(prog, routine, rung)]
    latch: Dict[Tuple[str, str], List[Tuple[str, str, int, str]]] = defaultdict(list)
    display: Dict[Tuple[str, str], str] = {}
    safety_tags: Dict[Tuple[str, str], bool] = {}
    for prog in ctrl.programs:
        for routine in prog.routines:
            for rung in routine.rungs:
                for ins in rung.instructions:
                    name = ins.name.upper()
                    if name not in ("OTE", "OTL", "OTU") or not ins.operands:
                        continue
                    op = ins.operands[0]
                    if op.kind != "tag":
                        continue
                    tag = xref.resolve(prog.name, op.tag_base or "", op.program)
                    ident = f"{tag.scope}|{tag.name}".lower() if tag else f"?|{op.tag_base}".lower()
                    key = (ident, (op.tag_path or op.raw).lower())
                    display.setdefault(key, op.tag_path or op.raw)
                    safety_tags[key] = safety_tags.get(key, False) or prog.is_safety or (tag.is_safety if tag else False)
                    if name == "OTE":
                        ote[key].append((prog.name, routine.name, rung.number))
                    else:
                        latch[key].append((prog.name, routine.name, rung.number, name))
    for key, places in ote.items():
        distinct = sorted(set(places))
        if len(distinct) > 1:
            p0 = distinct[0]
            out.append(Finding(
                rule="output_multiple_writes", severity="warning", program=p0[0], routine=p0[1], rung=p0[2], tag=display[key],
                message=f"Saída '{display[key]}' é escrita por OTE em {len(distinct)} rungs diferentes (a última na varredura vence).",
                evidence=_join([f"{p}/{r} rung {n}" for p, r, n in distinct]), safety=safety_tags.get(key, False),
            ))
    for key, places in latch.items():
        if key in ote:
            p0 = ote[key][0]
            out.append(Finding(
                rule="output_mixed_latch", severity="warning", program=p0[0], routine=p0[1], rung=p0[2], tag=display[key],
                message=f"Saída '{display[key]}' é escrita por OTE e também por OTL/OTU (comportamento imprevisível).",
                evidence=_join([f"{p}/{r} rung {n} (OTE)" for p, r, n in ote[key]] + [f"{p}/{r} rung {n} ({i})" for p, r, n, i in places]),
                safety=safety_tags.get(key, False),
            ))
    return out


def _lookup_value(values, path_rest: str):
    """Walk a decorated-values dict following '.Member' / '[i]' steps."""
    import re

    cur = values
    for step in re.findall(r"\.(\w+)|(\[[^\]]*\])", path_rest):
        member, index = step
        if not isinstance(cur, dict):
            return None
        if member:
            cur = cur.get(member)
        else:
            cur = cur.get(index.replace(" ", ""))
        if cur is None:
            return None
    return cur


def rule_timer_preset(ctrl: Controller, xref: CrossRef) -> List[Finding]:
    out: List[Finding] = []
    for prog in ctrl.programs:
        for routine in prog.routines:
            for rung in routine.rungs:
                for ins in rung.instructions:
                    if ins.name.upper() not in TIMER_COUNTER or len(ins.operands) < 2:
                        continue
                    struct, preset = ins.operands[0], ins.operands[1]
                    kind = "Timer" if ins.name.upper() in ("TON", "TOF", "RTO") else "Contador"
                    name = struct.tag_path or struct.raw
                    msg: Optional[str] = None
                    if preset.kind == "literal":
                        try:
                            if float(preset.raw) == 0:
                                msg = f"{kind} '{name}' ({ins.name}) com preset literal 0."
                        except ValueError:
                            pass
                    elif preset.kind == "unused" and struct.kind == "tag":
                        tag = xref.resolve(prog.name, struct.tag_base or "", struct.program)
                        if tag is not None and tag.values is not None:
                            rest = (struct.tag_path or "")[len(struct.tag_base or ""):]
                            val = _lookup_value(tag.values, rest)
                            pre = val.get("PRE") if isinstance(val, dict) else None
                            if pre is not None and pre == 0:
                                msg = f"{kind} '{name}' ({ins.name}) com preset 0 no valor da tag (PRE=0)."
                    if msg:
                        out.append(Finding(rule="timer_no_preset", severity="warning", program=prog.name, routine=routine.name, rung=rung.number, tag=name, message=msg, evidence=rung.text.strip()[:200], safety=prog.is_safety))
    return out


def rule_jsr_missing(ctrl: Controller, xref: CrossRef) -> List[Finding]:
    out: List[Finding] = []
    seen = set()
    for prog_name, routine_name, target in xref.missing_calls:
        key = (prog_name, routine_name, target.lower())
        if key in seen:
            continue
        seen.add(key)
        prog = ctrl.program(prog_name)
        rung_no = None
        routine = prog.routine(routine_name) if prog else None
        if routine:
            for rung in routine.rungs:
                if any(i.name.upper() in ("JSR", "FOR") and i.operands and i.operands[0].raw.lower() == target.lower() for i in rung.instructions):
                    rung_no = rung.number
                    break
        out.append(Finding(rule="jsr_missing_routine", severity="error", program=prog_name, routine=routine_name, rung=rung_no,
                           message=f"JSR para rotina inexistente '{target}' no programa '{prog_name}'.", safety=prog.is_safety if prog else False))
    return out


def rule_documentation(ctrl: Controller) -> List[Finding]:
    out: List[Finding] = []
    for prog in ctrl.programs:
        for routine in prog.routines:
            if routine.protected:
                continue
            if not routine.description:
                out.append(Finding(rule="routine_no_description", severity="info", program=prog.name, routine=routine.name,
                                   message=f"Rotina '{routine.name}' sem descrição.", safety=prog.is_safety))
            if routine.type == "RLL" and routine.rungs:
                missing = [r.number for r in routine.rungs if not r.comment]
                if missing:
                    pct = 100.0 * (len(routine.rungs) - len(missing)) / len(routine.rungs)
                    out.append(Finding(rule="rungs_without_comment", severity="info", program=prog.name, routine=routine.name, rung=missing[0],
                                       message=f"{len(missing)} de {len(routine.rungs)} rungs sem comentário ({pct:.0f}% comentados).",
                                       evidence="rungs " + _join([str(n) for n in missing], 15), safety=prog.is_safety))
    return out


def rule_protected(ctrl: Controller) -> List[Finding]:
    out: List[Finding] = []
    for prog in ctrl.programs:
        for routine in prog.routines:
            if routine.protected:
                out.append(Finding(rule="routine_protected", severity="info", program=prog.name, routine=routine.name,
                                   message=f"Rotina '{routine.name}' tem source protection (EncodedData): não analisável, conteúdo não lido."))
    for aoi in ctrl.aois:
        if aoi.protected:
            out.append(Finding(rule="aoi_protected", severity="info", tag=aoi.name,
                               message=f"AOI '{aoi.name}' (rev. {aoi.revision or '?'}) tem source protection: não analisável, conteúdo não lido."))
    return out


def rule_safety(ctrl: Controller) -> List[Finding]:
    out: List[Finding] = []
    for task in ctrl.tasks:
        if task.is_safety:
            out.append(Finding(rule="safety_content", severity="info", message=f"Task '{task.name}' é de segurança (Class=\"Safety\"): somente leitura, sem sugestões de alteração.", safety=True))
    for prog in ctrl.programs:
        if prog.is_safety:
            out.append(Finding(rule="safety_content", severity="info", program=prog.name, message=f"Programa '{prog.name}' é de segurança (Class=\"Safety\"): somente leitura, sem sugestões de alteração.", safety=True))
    for prog, tag in _iter_scoped_tags(ctrl):
        if tag.is_safety:
            out.append(Finding(rule="safety_content", severity="info", program=prog.name if prog else None, tag=tag.name, message=f"Tag '{tag.name}' é de segurança (Class=\"Safety\").", safety=True))
    return out


def rule_unknown_and_unresolved(ctrl: Controller, xref: CrossRef) -> Tuple[List[Finding], Dict[str, int]]:
    out: List[Finding] = []
    unknown: Counter = Counter()
    for prog in ctrl.programs:
        for routine in prog.routines:
            for rung in routine.rungs:
                for ins in rung.instructions:
                    if not ins.known:
                        unknown[ins.name] += 1
    for name, count in unknown.most_common():
        out.append(Finding(rule="unknown_instruction", severity="info", tag=name,
                           message=f"Instrução '{name}' não está na tabela do analisador ({count} ocorrências): operandos tratados como leitura."))
    for (scope, routine), names in xref.unresolved.items():
        if scope.startswith("aoi:"):
            continue
        uniq = sorted(set(names), key=str.lower)
        out.append(Finding(rule="unresolved_tags", severity="info", program=scope, routine=routine,
                           message=f"{len(uniq)} referência(s) a tags não encontradas em nenhum escopo (módulo removido, tag apagada ou parâmetro?).",
                           evidence=_join(uniq, 10)))
    return out, dict(unknown)


# --------------------------------------------------------------------------
# Metrics
# --------------------------------------------------------------------------
def compute_metrics(ctrl: Controller, findings: List[Finding], unknown: Dict[str, int]) -> Metrics:
    m = Metrics()
    m.programs = len(ctrl.programs)
    m.tasks = len(ctrl.tasks)
    m.tags_controller = len(ctrl.tags)
    m.tags_program = sum(len(p.tags) for p in ctrl.programs)
    m.aois = len(ctrl.aois)
    m.aois_protected = sum(1 for a in ctrl.aois if a.protected)
    m.udts = len(ctrl.data_types)
    m.modules = len(ctrl.modules)
    m.safety_programs = sum(1 for p in ctrl.programs if p.is_safety)
    by_type: Counter = Counter()
    commented = total_rungs = 0
    for prog in ctrl.programs:
        for r in prog.routines:
            m.routines += 1
            by_type[r.type] += 1
            if r.protected:
                m.protected_routines += 1
            m.st_lines += len(r.st_lines)
            if r.rungs:
                m.rungs += len(r.rungs)
                total_rungs += len(r.rungs)
                commented += sum(1 for g in r.rungs if g.comment)
                if len(r.rungs) > m.largest_routine_rungs:
                    m.largest_routine_rungs = len(r.rungs)
                    m.largest_routine = f"{prog.name}/{r.name}"
                for g in r.rungs:
                    m.instructions += len(g.instructions)
                    m.max_branch_depth = max(m.max_branch_depth, g.max_branch_depth)
    m.routines_by_type = dict(by_type)
    m.commented_rung_pct = round(100.0 * commented / total_rungs, 1) if total_rungs else 0.0
    m.unknown_instructions = unknown
    sev: Counter = Counter(f.severity for f in findings)
    m.findings_by_severity = {s: sev.get(s, 0) for s in SEVERITIES}
    return m


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def analyze(ctrl: Controller, xref: Optional[CrossRef] = None) -> Tuple[List[Finding], Metrics]:
    if xref is None:
        from .xref import build_xref

        xref = build_xref(ctrl)
    findings: List[Finding] = []
    findings += rule_jsr_missing(ctrl, xref)
    findings += rule_routine_never_called(ctrl, xref)
    findings += rule_program_scheduling(ctrl)
    findings += rule_output_writes(ctrl, xref)
    findings += rule_timer_preset(ctrl, xref)
    findings += rule_tag_usage(ctrl, xref)
    findings += rule_protected(ctrl)
    findings += rule_safety(ctrl)
    findings += rule_documentation(ctrl)
    extra, unknown = rule_unknown_and_unresolved(ctrl, xref)
    findings += extra
    findings += list(ctrl.parse_findings)
    findings.sort(key=lambda f: (_SEV_ORDER.get(f.severity, 9), f.program or "", f.routine or "", f.rung if f.rung is not None else -1))
    metrics = compute_metrics(ctrl, findings, unknown)
    return findings, metrics

"""Semantic diff between two L5X exports of the same project (old -> new).

Reports, by category: programs and routines added/removed, rung / ST line /
FBD node changes, tags added/removed/retyped/re-described, AOI and UDT
definition changes, modules and tasks, configuration values stored in tags
(AOI instances: config members only; timers: PRE; alarms: config members)
and findings that appeared or disappeared. Runtime members (ACC, DN, EN...)
are ignored so that two exports taken minutes apart do not look different.
"""

from __future__ import annotations

import difflib
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .aoi_config import flatten, members_for
from .model import Aoi, Controller, Finding, Routine, Tag
from .xref import AOI_PREFIX

CATEGORIES = ("programa", "rotina", "rung", "tag", "aoi", "udt", "modulo", "task", "config", "finding")
_RUNTIME = {"acc", "dn", "tt", "en", "cu", "cd", "ov", "un", "inalarm", "acked", "intime", "shelved", "suppressed", "disabled_sts"}


@dataclass
class Change:
    category: str  # one of CATEGORIES
    kind: str  # adicionado | removido | alterado
    item: str  # "Program/Routine", "scope/tag", AOI name...
    detail: str = ""  # member, rung number...
    old: Any = ""
    new: Any = ""

    def row(self) -> Dict[str, Any]:
        return {"Categoria": self.category, "Tipo": self.kind, "Item": self.item, "Detalhe": self.detail, "Antes": _short(self.old), "Depois": _short(self.new)}


@dataclass
class ProjectDiff:
    old_name: str
    new_name: str
    old_date: Optional[str]
    new_date: Optional[str]
    changes: List[Change] = field(default_factory=list)

    def by_category(self) -> Dict[str, List[Change]]:
        out: Dict[str, List[Change]] = {c: [] for c in CATEGORIES}
        for ch in self.changes:
            out.setdefault(ch.category, []).append(ch)
        return out

    def counts(self) -> Dict[str, int]:
        return {c: len(v) for c, v in self.by_category().items() if v}

    @property
    def empty(self) -> bool:
        return not self.changes


DIFF_HEADER = ["Categoria", "Tipo", "Item", "Detalhe", "Antes", "Depois"]


def _short(v: Any, limit: int = 160) -> str:
    s = " ".join(str(v).split()) if v is not None else ""
    return s if len(s) <= limit else s[: limit - 1] + "…"


def _norm(text: str) -> str:
    return " ".join(text.split())


# --------------------------------------------------------------------------
# Routines
# --------------------------------------------------------------------------
def _diff_sequences(old: List[str], new: List[str], item: str, unit: str, out: List[Change], old_numbers: Optional[List[int]] = None, new_numbers: Optional[List[int]] = None) -> None:
    sm = difflib.SequenceMatcher(None, old, new, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        if op == "replace":
            for k in range(max(i2 - i1, j2 - j1)):
                o = old[i1 + k] if i1 + k < i2 else None
                n = new[j1 + k] if j1 + k < j2 else None
                num = (new_numbers[j1 + k] if new_numbers and j1 + k < j2 else (old_numbers[i1 + k] if old_numbers and i1 + k < i2 else None))
                kind = "alterado" if o is not None and n is not None else ("adicionado" if o is None else "removido")
                out.append(Change("rung", kind, item, f"{unit} {num}" if num is not None else unit, o or "", n or ""))
        elif op == "insert":
            for k in range(j1, j2):
                out.append(Change("rung", "adicionado", item, f"{unit} {new_numbers[k] if new_numbers else k}", "", new[k]))
        elif op == "delete":
            for k in range(i1, i2):
                out.append(Change("rung", "removido", item, f"{unit} {old_numbers[k] if old_numbers else k}", old[k], ""))


def diff_routine(old: Routine, new: Routine, item: str, out: List[Change]) -> None:
    if (old.description or "") != (new.description or ""):
        out.append(Change("rotina", "alterado", item, "descrição", old.description or "", new.description or ""))
    if old.type != new.type:
        out.append(Change("rotina", "alterado", item, "tipo", old.type, new.type))
        return
    if old.protected != new.protected:
        out.append(Change("rotina", "alterado", item, "proteção", "protegida" if old.protected else "aberta", "protegida" if new.protected else "aberta"))
        return
    if new.type == "RLL":
        _diff_sequences([_norm(r.text) for r in old.rungs], [_norm(r.text) for r in new.rungs], item, "rung", out, [r.number for r in old.rungs], [r.number for r in new.rungs])
    elif new.type == "ST":
        _diff_sequences([_norm(l.text) for l in old.st_lines], [_norm(l.text) for l in new.st_lines], item, "linha", out, [l.number for l in old.st_lines], [l.number for l in new.st_lines])
    elif new.type == "FBD":
        def sig(n: Any) -> str:
            ins = "; ".join(f"{k}={v}" for k, v in n.inputs.items())
            outs = "; ".join(f"{k}->{v}" for k, v in n.outputs.items())
            return f"folha {n.sheet}: {n.kind} {n.label}: {ins} {outs}".strip()

        old_set = {sig(n) for n in old.fbd_nodes}
        new_set = {sig(n) for n in new.fbd_nodes}
        for s in sorted(old_set - new_set):
            out.append(Change("rung", "removido", item, "bloco", s, ""))
        for s in sorted(new_set - old_set):
            out.append(Change("rung", "adicionado", item, "bloco", "", s))
        old_d = {f"{d.kind} {d.operand} (folha {d.sheet})" for d in old.dangling}
        new_d = {f"{d.kind} {d.operand} (folha {d.sheet})" for d in new.dangling}
        for s in sorted(new_d - old_d):
            out.append(Change("rung", "adicionado", item, "sem fio", "", s))
        for s in sorted(old_d - new_d):
            out.append(Change("rung", "removido", item, "sem fio", s, ""))


def _diff_routines(old_rts: Sequence[Routine], new_rts: Sequence[Routine], prefix: str, out: List[Change]) -> None:
    o = {r.name.lower(): r for r in old_rts}
    n = {r.name.lower(): r for r in new_rts}
    for key in sorted(set(o) - set(n)):
        out.append(Change("rotina", "removido", f"{prefix}/{o[key].name}", o[key].type, f"{len(o[key].rungs)} rungs" if o[key].type == "RLL" else "", ""))
    for key in sorted(set(n) - set(o)):
        out.append(Change("rotina", "adicionado", f"{prefix}/{n[key].name}", n[key].type, "", f"{len(n[key].rungs)} rungs" if n[key].type == "RLL" else ""))
    for key in sorted(set(o) & set(n)):
        diff_routine(o[key], n[key], f"{prefix}/{n[key].name}", out)


# --------------------------------------------------------------------------
# Tags and stored configuration
# --------------------------------------------------------------------------
def _config_values(tag: Tag, ctrl: Controller) -> Dict[str, Any]:
    """Members of a tag worth comparing: config members of AOI instances, PRE of timers,
    config of alarms; runtime members dropped."""
    if not isinstance(tag.values, dict):
        return {}
    aoi = ctrl.aoi(tag.data_type)
    flat = flatten(tag.values)
    if aoi is not None and not aoi.protected:
        keep = {m.name.lower() for m in members_for(aoi, "config")}
        return {k: v for k, v in flat.items() if k.split(".")[0].split("[")[-1].lower() in keep or k.split(".")[0].lower() in keep}
    out = {}
    for k, v in flat.items():
        leaf = k.rsplit(".", 1)[-1].lower()
        if leaf in _RUNTIME:
            continue
        if tag.data_type.upper() in ("TIMER", "COUNTER") and leaf != "pre":
            continue
        out[k] = v
    return out


def _diff_tags(old_tags: Sequence[Tag], new_tags: Sequence[Tag], scope: str, old_ctrl: Controller, new_ctrl: Controller, out: List[Change]) -> None:
    o = {t.name.lower(): t for t in old_tags}
    n = {t.name.lower(): t for t in new_tags}
    for key in sorted(set(o) - set(n)):
        out.append(Change("tag", "removido", f"{scope}/{o[key].name}", o[key].data_type, "", ""))
    for key in sorted(set(n) - set(o)):
        out.append(Change("tag", "adicionado", f"{scope}/{n[key].name}", n[key].data_type, "", n[key].description or ""))
    for key in sorted(set(o) & set(n)):
        a, b = o[key], n[key]
        item = f"{scope}/{b.name}"
        for attr, label in (("data_type", "tipo"), ("dimensions", "dimensões"), ("alias_for", "alias"), ("description", "descrição"), ("external_access", "ExternalAccess"), ("constant", "constante")):
            if (getattr(a, attr) or "") != (getattr(b, attr) or ""):
                out.append(Change("tag", "alterado", item, label, getattr(a, attr) or "", getattr(b, attr) or ""))
        va, vb = _config_values(a, old_ctrl), _config_values(b, new_ctrl)
        for member in sorted(set(va) | set(vb), key=str.lower):
            if va.get(member, "") != vb.get(member, ""):
                out.append(Change("config", "alterado", item, member, va.get(member, ""), vb.get(member, "")))


# --------------------------------------------------------------------------
# Definitions
# --------------------------------------------------------------------------
def _diff_aois(old_ctrl: Controller, new_ctrl: Controller, out: List[Change]) -> None:
    o = {a.name.lower(): a for a in old_ctrl.aois}
    n = {a.name.lower(): a for a in new_ctrl.aois}
    for key in sorted(set(o) - set(n)):
        out.append(Change("aoi", "removido", o[key].name, f"rev. {o[key].revision or '?'}"))
    for key in sorted(set(n) - set(o)):
        out.append(Change("aoi", "adicionado", n[key].name, f"rev. {n[key].revision or '?'}"))
    for key in sorted(set(o) & set(n)):
        a, b = o[key], n[key]
        if (a.revision or "") != (b.revision or ""):
            out.append(Change("aoi", "alterado", b.name, "revisão", a.revision or "?", b.revision or "?"))
        if a.protected != b.protected:
            out.append(Change("aoi", "alterado", b.name, "proteção", "protegida" if a.protected else "aberta", "protegida" if b.protected else "aberta"))
            continue
        if a.protected:
            continue
        pa = {p.name.lower(): p for p in a.parameters}
        pb = {p.name.lower(): p for p in b.parameters}
        for k in sorted(set(pa) - set(pb)):
            out.append(Change("aoi", "removido", b.name, f"parâmetro {pa[k].name}", f"{pa[k].usage} {pa[k].data_type}", ""))
        for k in sorted(set(pb) - set(pa)):
            out.append(Change("aoi", "adicionado", b.name, f"parâmetro {pb[k].name}", "", f"{pb[k].usage} {pb[k].data_type}"))
        for k in sorted(set(pa) & set(pb)):
            x, y = pa[k], pb[k]
            for attr, label in (("usage", "uso"), ("data_type", "tipo"), ("required", "obrigatório"), ("visible", "visível"), ("default", "default"), ("description", "descrição")):
                if (getattr(x, attr) if getattr(x, attr) is not None else "") != (getattr(y, attr) if getattr(y, attr) is not None else ""):
                    out.append(Change("aoi", "alterado", b.name, f"parâmetro {y.name}: {label}", getattr(x, attr), getattr(y, attr)))
        la = {t.name.lower(): t for t in a.local_tags if not t.usage}
        lb = {t.name.lower(): t for t in b.local_tags if not t.usage}
        for k in sorted(set(la) - set(lb)):
            out.append(Change("aoi", "removido", b.name, f"tag local {la[k].name}", la[k].data_type, ""))
        for k in sorted(set(lb) - set(la)):
            out.append(Change("aoi", "adicionado", b.name, f"tag local {lb[k].name}", "", lb[k].data_type))
        for k in sorted(set(la) & set(lb)):
            if la[k].data_type != lb[k].data_type:
                out.append(Change("aoi", "alterado", b.name, f"tag local {lb[k].name}: tipo", la[k].data_type, lb[k].data_type))
            if (la[k].values if la[k].values is not None else "") != (lb[k].values if lb[k].values is not None else ""):
                out.append(Change("aoi", "alterado", b.name, f"tag local {lb[k].name}: default", la[k].values, lb[k].values))
        for flag in ("execute_prescan", "execute_postscan", "execute_enableinfalse"):
            if getattr(a, flag) != getattr(b, flag):
                out.append(Change("aoi", "alterado", b.name, flag, getattr(a, flag), getattr(b, flag)))
        _diff_routines(a.routines, b.routines, f"{AOI_PREFIX}{b.name}", out)


def _diff_udts(old_ctrl: Controller, new_ctrl: Controller, out: List[Change]) -> None:
    o = {d.name.lower(): d for d in old_ctrl.data_types}
    n = {d.name.lower(): d for d in new_ctrl.data_types}
    for key in sorted(set(o) - set(n)):
        out.append(Change("udt", "removido", o[key].name))
    for key in sorted(set(n) - set(o)):
        out.append(Change("udt", "adicionado", n[key].name))
    for key in sorted(set(o) & set(n)):
        ma = {m.name.lower(): m for m in o[key].members}
        mb = {m.name.lower(): m for m in n[key].members}
        for k in sorted(set(ma) - set(mb)):
            out.append(Change("udt", "removido", n[key].name, f"membro {ma[k].name}", ma[k].data_type, ""))
        for k in sorted(set(mb) - set(ma)):
            out.append(Change("udt", "adicionado", n[key].name, f"membro {mb[k].name}", "", mb[k].data_type))
        for k in sorted(set(ma) & set(mb)):
            if (ma[k].data_type, ma[k].dimension) != (mb[k].data_type, mb[k].dimension):
                out.append(Change("udt", "alterado", n[key].name, f"membro {mb[k].name}", f"{ma[k].data_type}[{ma[k].dimension}]", f"{mb[k].data_type}[{mb[k].dimension}]"))


def _diff_modules_tasks(old_ctrl: Controller, new_ctrl: Controller, out: List[Change]) -> None:
    o = {m.name.lower(): m for m in old_ctrl.modules}
    n = {m.name.lower(): m for m in new_ctrl.modules}
    for key in sorted(set(o) - set(n)):
        out.append(Change("modulo", "removido", o[key].name, o[key].catalog or ""))
    for key in sorted(set(n) - set(o)):
        out.append(Change("modulo", "adicionado", n[key].name, n[key].catalog or ""))
    for key in sorted(set(o) & set(n)):
        a, b = o[key], n[key]
        for attr, label in (("catalog", "catálogo"), ("parent", "pai"), ("slot", "slot"), ("major", "firmware major"), ("minor", "firmware minor"), ("inhibited", "inibido")):
            if (getattr(a, attr) or "") != (getattr(b, attr) or ""):
                out.append(Change("modulo", "alterado", b.name, label, getattr(a, attr) or "", getattr(b, attr) or ""))
        ca = {c.path.lower(): c for c in a.channels}
        cb = {c.path.lower(): c for c in b.channels}
        for k in sorted(set(ca) | set(cb)):
            x, y = ca.get(k), cb.get(k)
            if x is None or y is None or (x.description or "") != (y.description or ""):
                out.append(Change("modulo", "alterado" if x and y else ("adicionado" if y else "removido"), b.name, f"canal {(y or x).path}", x.description if x else "", y.description if y else ""))
    o = {t.name.lower(): t for t in old_ctrl.tasks}
    n = {t.name.lower(): t for t in new_ctrl.tasks}
    for key in sorted(set(o) - set(n)):
        out.append(Change("task", "removido", o[key].name, o[key].type))
    for key in sorted(set(n) - set(o)):
        out.append(Change("task", "adicionado", n[key].name, n[key].type))
    for key in sorted(set(o) & set(n)):
        a, b = o[key], n[key]
        for attr, label in (("type", "tipo"), ("rate", "período"), ("priority", "prioridade")):
            if (getattr(a, attr) or "") != (getattr(b, attr) or ""):
                out.append(Change("task", "alterado", b.name, label, getattr(a, attr) or "", getattr(b, attr) or ""))
        if [p.lower() for p in a.programs] != [p.lower() for p in b.programs]:
            out.append(Change("task", "alterado", b.name, "programas", ", ".join(a.programs), ", ".join(b.programs)))


def _finding_key(f: Finding) -> Tuple[str, str, str, str]:
    return (f.rule, (f.program or "").lower(), (f.routine or "").lower(), (f.tag or "").lower())


def _diff_findings(old: Sequence[Finding], new: Sequence[Finding], out: List[Change]) -> None:
    o = {_finding_key(f): f for f in old}
    n = {_finding_key(f): f for f in new}
    for key in sorted(set(n) - set(o)):
        f = n[key]
        out.append(Change("finding", "adicionado", f.location(), f"[{f.severity}] {f.rule}", "", f.message))
    for key in sorted(set(o) - set(n)):
        f = o[key]
        out.append(Change("finding", "removido", f.location(), f"[{f.severity}] {f.rule}", f.message, ""))


# --------------------------------------------------------------------------
# Entry point and renderers
# --------------------------------------------------------------------------
def diff_projects(old: Controller, new: Controller, old_findings: Optional[Sequence[Finding]] = None, new_findings: Optional[Sequence[Finding]] = None) -> ProjectDiff:
    d = ProjectDiff(old_name=old.name, new_name=new.name, old_date=old.export_date, new_date=new.export_date)
    out = d.changes
    op = {p.name.lower(): p for p in old.programs}
    np_ = {p.name.lower(): p for p in new.programs}
    for key in sorted(set(op) - set(np_)):
        out.append(Change("programa", "removido", op[key].name, f"{len(op[key].routines)} rotinas"))
    for key in sorted(set(np_) - set(op)):
        out.append(Change("programa", "adicionado", np_[key].name, f"{len(np_[key].routines)} rotinas"))
    for key in sorted(set(op) & set(np_)):
        a, b = op[key], np_[key]
        for attr, label in (("main_routine", "main"), ("fault_routine", "fault"), ("disabled", "desabilitado"), ("task", "task"), ("description", "descrição")):
            if (getattr(a, attr) or "") != (getattr(b, attr) or ""):
                out.append(Change("programa", "alterado", b.name, label, getattr(a, attr) or "", getattr(b, attr) or ""))
        _diff_routines(a.routines, b.routines, b.name, out)
        _diff_tags(a.tags, b.tags, b.name, old, new, out)
    _diff_tags(old.tags, new.tags, "controlador", old, new, out)
    _diff_aois(old, new, out)
    _diff_udts(old, new, out)
    _diff_modules_tasks(old, new, out)
    if old_findings is not None and new_findings is not None:
        _diff_findings(old_findings, new_findings, out)
    return d


def diff_rows(d: ProjectDiff) -> List[Dict[str, Any]]:
    return [c.row() for c in d.changes]


def diff_markdown(d: ProjectDiff, limit_per_category: int = 80) -> str:
    L = [f"# Diff: {d.old_name} ({d.old_date or '?'}) → {d.new_name} ({d.new_date or '?'})", ""]
    counts = d.counts()
    if not counts:
        L.append("Nenhuma diferença semântica encontrada.")
        return "\n".join(L)
    L.append("Resumo: " + ", ".join(f"{k}: {v}" for k, v in counts.items()))
    L.append("")
    titles = {"programa": "Programas", "rotina": "Rotinas", "rung": "Lógica (rungs, linhas ST, blocos FBD)", "tag": "Tags", "aoi": "AOIs", "udt": "UDTs", "modulo": "Módulos e canais", "task": "Tasks", "config": "Configuração armazenada nas tags", "finding": "Findings"}
    for cat, changes in d.by_category().items():
        if not changes:
            continue
        L.append(f"## {titles.get(cat, cat)} ({len(changes)})")
        L.append("")
        L.append("| Tipo | Item | Detalhe | Antes | Depois |")
        L.append("| --- | --- | --- | --- | --- |")
        for c in changes[:limit_per_category]:
            L.append(f"| {c.kind} | {c.item} | {c.detail} | {_short(c.old, 120).replace('|', '\\|')} | {_short(c.new, 120).replace('|', '\\|')} |")
        if len(changes) > limit_per_category:
            L.append(f"| … (+{len(changes) - limit_per_category}) | | | | |")
        L.append("")
    return "\n".join(L)

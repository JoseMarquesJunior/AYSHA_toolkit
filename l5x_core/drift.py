"""Configuration drift.

* ``instance_drift``: inside one project, instances of the same AOI whose
  configuration member differs from the value shared by the large majority
  (e.g. one valve with travel time 7000 when all others use 5000).
* ``library_drift``: across several projects (PLCs), AOIs and UDTs present
  with different revisions or interfaces, or missing in some of them.
"""

from __future__ import annotations

from collections import Counter
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .aoi_config import build_sheets
from .model import Aoi, Controller, DataType
from .xref import CrossRef, build_xref

DRIFT_HEADER = ["AOI", "Membro", "ValorComum", "Participacao", "Instancias", "Divergentes", "Detalhe"]
LIBRARY_HEADER = ["Item", "Tipo", "Status", "Detalhe"]


def instance_drift(ctrl: Controller, xref: Optional[CrossRef] = None, mode: str = "config", min_instances: int = 4, majority: float = 0.75, max_outliers: int = 10) -> List[Dict[str, Any]]:
    """One row per (AOI, member) where >= ``majority`` of the instances share one value
    and the others deviate. Members whose values are all different are not drift."""
    xref = xref or build_xref(ctrl)
    rows: List[Dict[str, Any]] = []
    for sheet in build_sheets(ctrl, xref, mode=mode):
        if len(sheet.rows) < min_instances:
            continue
        for col in sheet.columns:
            pairs = [(r["Instancia"], r[col]) for r in sheet.rows if r.get(col, "") != ""]
            if len(pairs) < min_instances:
                continue
            counts = Counter(str(v) for _, v in pairs)
            common, n = counts.most_common(1)[0]
            share = n / len(pairs)
            if share < majority or n == len(pairs):
                continue
            outliers = [(inst, v) for inst, v in pairs if str(v) != common]
            rows.append({
                "AOI": sheet.aoi, "Membro": col, "ValorComum": common, "Participacao": f"{100 * share:.0f}%", "Instancias": len(pairs),
                "Divergentes": len(outliers),
                "Detalhe": "; ".join(f"{inst}={v}" for inst, v in outliers[:max_outliers]) + (f" … (+{len(outliers) - max_outliers})" if len(outliers) > max_outliers else ""),
            })
    rows.sort(key=lambda r: (-r["Divergentes"] / max(1, r["Instancias"]), r["AOI"].lower(), r["Membro"].lower()))
    return rows


def _aoi_signature(aoi: Aoi) -> Tuple[Tuple[str, str, str, bool], ...]:
    return tuple((p.name.lower(), p.usage, p.data_type.upper(), p.required) for p in aoi.parameters if not p.is_system)


def _udt_signature(udt: DataType) -> Tuple[Tuple[str, str, int], ...]:
    return tuple((m.name.lower(), m.data_type.upper(), m.dimension) for m in udt.members if not m.hidden)


def library_drift(ctrls: Sequence[Controller]) -> List[Dict[str, Any]]:
    """AOIs and UDTs compared across controllers: same, revision differs, interface differs, missing."""
    rows: List[Dict[str, Any]] = []
    names = [c.name for c in ctrls]
    aoi_names = sorted({a.name.lower(): a.name for c in ctrls for a in c.aois}.values(), key=str.lower)
    for name in aoi_names:
        present = [(c.name, c.aoi(name)) for c in ctrls]
        found = [(cn, a) for cn, a in present if a is not None]
        missing = [cn for cn, a in present if a is None]
        revs = {a.revision or "?" for _, a in found}
        sigs = {_aoi_signature(a) for _, a in found if not a.protected}
        if missing:
            status = "ausente em " + ", ".join(missing)
        elif len(sigs) > 1:
            status = "interface diferente"
        elif len(revs) > 1:
            status = "revisão diferente"
        else:
            status = "igual"
        detail = "; ".join(f"{cn}: rev. {a.revision or '?'}" + (" (protegida)" if a.protected else f", {len(a.user_parameters)} par.") for cn, a in found)
        if len(sigs) > 1:
            params = [set(p[0] for p in _aoi_signature(a)) for _, a in found if not a.protected]
            union, inter = set.union(*params), set.intersection(*params)
            if union - inter:
                detail += " | parâmetros só em algumas: " + ", ".join(sorted(union - inter)[:10])
        rows.append({"Item": name, "Tipo": "AOI", "Status": status, "Detalhe": detail})
    udt_names = sorted({d.name.lower(): d.name for c in ctrls for d in c.data_types}.values(), key=str.lower)
    for name in udt_names:
        present = [(c.name, next((d for d in c.data_types if d.name.lower() == name.lower()), None)) for c in ctrls]
        found = [(cn, d) for cn, d in present if d is not None]
        missing = [cn for cn, d in present if d is None]
        sigs = {_udt_signature(d) for _, d in found}
        if missing:
            status = "ausente em " + ", ".join(missing)
        elif len(sigs) > 1:
            status = "membros diferentes"
        else:
            status = "igual"
        rows.append({"Item": name, "Tipo": "UDT", "Status": status, "Detalhe": "; ".join(f"{cn}: {len(d.members)} membros" for cn, d in found)})
    rows.sort(key=lambda r: (r["Status"] == "igual", r["Tipo"], r["Item"].lower()))
    return rows


def drift_markdown(rows: List[Dict[str, Any]], limit: int = 60) -> str:
    L = ["# Drift entre instâncias", "", f"{len(rows)} membro(s) com valor fora do padrão da maioria.", ""]
    if rows:
        L.append("| AOI | Membro | Valor comum | Participação | Divergentes | Detalhe |")
        L.append("| --- | --- | --- | --- | --- | --- |")
        for r in rows[:limit]:
            L.append(f"| {r['AOI']} | {r['Membro']} | {r['ValorComum']} | {r['Participacao']} | {r['Divergentes']}/{r['Instancias']} | {str(r['Detalhe']).replace('|', '/')} |")
        if len(rows) > limit:
            L.append(f"| … (+{len(rows) - limit}) | | | | | |")
    return "\n".join(L)


def library_markdown(rows: List[Dict[str, Any]], names: Sequence[str]) -> str:
    diff = [r for r in rows if r["Status"] != "igual"]
    L = ["# Drift de biblioteca entre PLCs", "", f"Controladores: {', '.join(names)}. {len(diff)} item(ns) divergente(s) de {len(rows)}.", ""]
    L.append("| Item | Tipo | Status | Detalhe |")
    L.append("| --- | --- | --- | --- |")
    for r in diff:
        L.append(f"| {r['Item']} | {r['Tipo']} | {r['Status']} | {str(r['Detalhe']).replace('|', '/')} |")
    return "\n".join(L)

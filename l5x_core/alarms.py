"""Alarm inventory for rationalization (ISA 18.2 style).

Alarms come from three places:
* native ALMD / ALMA instructions (tags of type ALARM_DIGITAL / ALARM_ANALOG);
* instances of alarm-like AOIs (name contains "alarm"/"alm", or a parameter
  named Severity / Cfg_Severity), e.g. PlantPAx P_Alarm;
* alarm AOIs nested inside other AOIs (AI_TYPE5 holding several P_Alarm locals):
  reached through the parent instance tag (``LIT_100.AHH_Alm``).

Each row carries the configuration members (Cfg_*, Severity, limits, ack /
latch / shelve flags) read from the instance data, plus where it is used.
"""

from __future__ import annotations

from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

from .aoi_config import flatten
from .model import Aoi, Controller, Tag
from .xref import AOI_PREFIX, CrossRef, build_xref

BASE_HEADER = ["Alarme", "Tipo", "Escopo", "Instancia", "InstanciaTipo", "Descricao", "DescricaoMembro", "Severidade", "Habilitado", "Uso"]
KNOWN_CONFIG = {
    "severity", "minduration", "mindurationpre", "shelveduration", "ackrequired", "latched", "disabled", "useprogtime",
    "hhlimit", "hlimit", "llimit", "lllimit", "deadband", "rocposlimit", "rocneglimit", "rocperiod", "hhminduration", "hminduration", "lminduration", "llminduration",
}
NATIVE_TYPES = {"ALARM_DIGITAL": "ALMD", "ALARM_ANALOG": "ALMA"}


def is_alarm_aoi(aoi: Aoi) -> bool:
    low = aoi.name.lower()
    if "alarm" in low or "alm" in low:
        return True
    return any(p.name.lower() in ("severity", "cfg_severity") for p in aoi.parameters)


def _config_member(name: str) -> bool:
    low = name.lower().replace("_", "")
    return name.lower().startswith("cfg_") or low in KNOWN_CONFIG


def alarm_paths(ctrl: Controller, aoi: Aoi, depth: int = 3) -> List[Tuple[str, Aoi, Optional[str]]]:
    """(member path, alarm AOI, member description) for alarm instances nested in ``aoi``."""
    out: List[Tuple[str, Aoi, Optional[str]]] = []
    if depth <= 0:
        return out
    for t in aoi.local_tags:
        if t.usage in ("InOut",):
            continue
        child = ctrl.aoi(t.data_type)
        if child is None or child.protected:
            continue
        if is_alarm_aoi(child):
            out.append((t.name, child, t.description))
        else:
            for path, alm, desc in alarm_paths(ctrl, child, depth - 1):
                out.append((f"{t.name}.{path}", alm, desc or t.description))
    return out


def _severity(values: Dict[str, Any]) -> Any:
    for key in ("Cfg_Severity", "Severity"):
        for k, v in values.items():
            if k.lower() == key.lower():
                return v
    return ""


def _enabled(values: Dict[str, Any]) -> Any:
    for k, v in values.items():
        if k.lower() in ("cfg_exists", "exists"):
            return v
        if k.lower() == "disabled":
            return 0 if v else 1
    return ""


def alarm_rows(ctrl: Controller, xref: Optional[CrossRef] = None) -> Tuple[List[str], List[Dict[str, Any]]]:
    """(header, rows). Config columns are discovered from the data, in first-seen order."""
    xref = xref or build_xref(ctrl)
    rows: List[Dict[str, Any]] = []
    config_cols: List[str] = []

    def add(values: Dict[str, Any], **fixed: Any) -> None:
        row = dict(fixed)
        row["Severidade"] = _severity(values)
        row["Habilitado"] = _enabled(values)
        for k, v in values.items():
            if k.lower() in ("cfg_severity", "severity", "cfg_exists", "exists"):
                continue
            if _config_member(k) and not isinstance(v, dict):
                if k not in config_cols:
                    config_cols.append(k)
                row[k] = v
        rows.append(row)

    def where(tag: Tag) -> str:
        sites = xref.calls_of_instance(tag)
        if sites:
            return ", ".join(f"{s.scope}/{s.routine} {s.call.location}" for s in sites[:3])
        uses = [u for u in xref.usages_of(tag)]
        return ", ".join(sorted({f"{u.scope}/{u.routine}" for u in uses})[:3])

    scope_label = lambda t: "controlador" if t.scope == "controller" else t.scope  # noqa: E731

    # native ALMD / ALMA
    for tag in list(ctrl.tags) + [t for p in ctrl.programs for t in p.tags]:
        kind = NATIVE_TYPES.get(tag.data_type.upper())
        if kind is None:
            continue
        values = flatten(tag.values) if isinstance(tag.values, dict) else {}
        add(values, Alarme=tag.name, Tipo=kind, Escopo=scope_label(tag), Instancia=tag.name, InstanciaTipo=tag.data_type,
            Descricao=" ".join((tag.description or "").split()), DescricaoMembro="", Uso=where(tag))

    # AOI based (direct instances and nested)
    for aoi in ctrl.aois:
        direct = is_alarm_aoi(aoi) and not aoi.protected
        nested = [] if aoi.protected else alarm_paths(ctrl, aoi)
        if not direct and not nested:
            continue
        for tag in xref.instances_of(aoi.name):
            if tag.scope.startswith(AOI_PREFIX):
                continue
            values = tag.values if isinstance(tag.values, dict) else {}
            desc = " ".join((tag.description or "").split())
            if direct:
                add(flatten(values), Alarme=tag.name, Tipo=aoi.name, Escopo=scope_label(tag), Instancia=tag.name, InstanciaTipo=aoi.name,
                    Descricao=desc, DescricaoMembro="", Uso=where(tag))
            for path, alm, mdesc in nested:
                sub: Any = values
                for part in path.split("."):
                    sub = sub.get(part) if isinstance(sub, dict) else None
                add(flatten(sub) if isinstance(sub, dict) else {}, Alarme=f"{tag.name}.{path}", Tipo=alm.name, Escopo=scope_label(tag), Instancia=tag.name,
                    InstanciaTipo=aoi.name, Descricao=desc, DescricaoMembro=" ".join((mdesc or "").split()), Uso=where(tag))
    header = BASE_HEADER + config_cols
    for r in rows:
        for c in config_cols:
            r.setdefault(c, "")
    return header, rows


def alarm_summary(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Counts used by the rationalization checks."""
    sev = Counter(str(r["Severidade"]) if r["Severidade"] != "" else "(sem)" for r in rows)
    by_type = Counter(r["Tipo"] for r in rows)
    disabled = sum(1 for r in rows if r["Habilitado"] in (0, "0", False))
    no_desc = sum(1 for r in rows if not r["Descricao"] and not r["DescricaoMembro"])
    unused = sum(1 for r in rows if not r["Uso"])
    dup = Counter((r["Descricao"], r["DescricaoMembro"]) for r in rows if r["Descricao"])
    duplicates = sum(1 for k, n in dup.items() if n > 1)
    return {"total": len(rows), "por_tipo": dict(by_type), "por_severidade": dict(sev), "desabilitados": disabled, "sem_descricao": no_desc, "sem_uso": unused, "descricoes_duplicadas": duplicates}


def alarm_markdown(header: List[str], rows: List[Dict[str, Any]], limit: int = 50) -> str:
    s = alarm_summary(rows)
    L = ["# Lista de alarmes", "",
         f"- Total: {s['total']} (" + ", ".join(f"{k}: {v}" for k, v in sorted(s["por_tipo"].items())) + ")",
         "- Por severidade: " + ", ".join(f"{k}: {v}" for k, v in sorted(s["por_severidade"].items(), key=lambda kv: kv[0])),
         f"- Desabilitados (Cfg_Exists=0): {s['desabilitados']} · sem descrição: {s['sem_descricao']} · sem uso na lógica: {s['sem_uso']} · descrições duplicadas: {s['descricoes_duplicadas']}", ""]
    cols = ["Alarme", "Tipo", "Descricao", "DescricaoMembro", "Severidade", "Habilitado", "Uso"]
    L.append("| " + " | ".join(cols) + " |")
    L.append("|" + " --- |" * len(cols))
    for r in rows[:limit]:
        L.append("| " + " | ".join(str(r.get(c, "")).replace("|", "\\|") for c in cols) + " |")
    if len(rows) > limit:
        L.append(f"| … (+{len(rows) - limit}) |" + " |" * (len(cols) - 1))
    return "\n".join(L)

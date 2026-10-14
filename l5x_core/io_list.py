"""Instrument list and I/O channel list derived from the program.

Studio 5000 keeps channel comments inside each module and instrument data
inside AOI instance tags; nothing joins them. This module produces:

* ``io_channels``: one row per commented module point (module, slot, channel,
  description, unit, limits, spare?, routines that use it).
* ``instruments``: one row per AOI instance that looks like a field device
  (has an engineering unit, a range or a linked I/O channel): tag, type,
  description, unit, area, range, linked input/output channels, call site.

The channel linking is approximate: direct I/O operands on the call, plus one
hop through a tag that is copied from / to an I/O point (MOV/COP in ladder,
XIC->OTE in the same rung, direct wires in FBD).
"""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from .aoi_config import flatten
from .model import Controller, Tag
from .xref import AOI_PREFIX, CrossRef, build_xref

CHANNEL_HEADER = ["Modulo", "Catalogo", "Slot", "Direcao", "Canal", "Descricao", "Unidade", "Min", "Max", "Spare", "Usos", "Rotinas"]
INSTRUMENT_HEADER = ["Tag", "Tipo", "Descricao", "Unidade", "Area", "Label", "Escopo", "RangeMin", "RangeMax", "EntradaIO", "EntradaIODescricao", "SaidaIO", "Chamada"]


def _is_io(path: Optional[str]) -> bool:
    return bool(path) and ":" in path


def range_kind(member: str) -> Optional[str]:
    """'min' / 'max' when a member name looks like a range limit (rRangeMin, EU_Max, Cfg_RangeHi...)."""
    n = member.lower().rsplit(".", 1)[-1]
    if not any(k in n for k in ("range", "eu", "scale", "span", "eng")):
        return None
    if "min" in n or n.endswith("lo") or "low" in n:
        return "min"
    if "max" in n or n.endswith("hi") or "high" in n:
        return "max"
    return None


def io_source_map(ctrl: Controller) -> Tuple[Dict[str, str], Dict[str, str]]:
    """(tag -> I/O input path it is copied from, tag -> I/O output path it feeds), lower-cased tags."""
    source: Dict[str, str] = {}
    sink: Dict[str, str] = {}
    routines = [r for p in ctrl.programs for r in p.routines] + [r for a in ctrl.aois for r in a.routines]
    for r in routines:
        for rung in r.rungs:
            io_inputs = [op.tag_path for ins in rung.instructions for op in ins.operands if op.kind == "tag" and _is_io(op.tag_path) and ins.name.upper() in ("XIC", "XIO")]
            for ins in rung.instructions:
                name = ins.name.upper()
                ops = ins.operands
                if name in ("MOV", "COP", "CPS") and len(ops) >= 2 and ops[0].kind == "tag" and ops[1].kind == "tag":
                    if _is_io(ops[0].tag_path) and not _is_io(ops[1].tag_path):
                        source.setdefault(ops[1].tag_path.lower(), ops[0].tag_path)
                    elif _is_io(ops[1].tag_path) and not _is_io(ops[0].tag_path):
                        sink.setdefault(ops[0].tag_path.lower(), ops[1].tag_path)
                elif name in ("OTE", "OTL") and ops and ops[0].kind == "tag":
                    if _is_io(ops[0].tag_path):
                        for ins2 in rung.instructions:
                            if ins2.name.upper() == "XIC" and ins2.operands and ins2.operands[0].kind == "tag" and not _is_io(ins2.operands[0].tag_path):
                                sink.setdefault(ins2.operands[0].tag_path.lower(), ops[0].tag_path)
                    elif len(io_inputs) == 1:
                        source.setdefault(ops[0].tag_path.lower(), io_inputs[0])
        for node in r.fbd_nodes:
            if node.kind == "Wire":
                src, dst = node.inputs.get("from", ""), node.outputs.get("to", "")
                if _is_io(src) and not _is_io(dst):
                    source.setdefault(dst.lower(), src)
                elif _is_io(dst) and not _is_io(src):
                    sink.setdefault(src.lower(), dst)
    return source, sink


def io_channels(ctrl: Controller, xref: Optional[CrossRef] = None) -> List[Dict[str, Any]]:
    xref = xref or build_xref(ctrl)
    rows: List[Dict[str, Any]] = []
    for mod in ctrl.modules:
        if not mod.channels:
            continue
        rack = (mod.parent or "").lower()
        usages = xref.io_usages.get(rack, [])
        for ch in mod.channels:
            prefix = ch.path.lower()
            hits = [u for u in usages if u.path.lower() == prefix or u.path.lower().startswith(prefix + ".")]
            where = sorted({f"{u.scope}/{u.routine}" for u in hits})
            rows.append({
                "Modulo": mod.name, "Catalogo": mod.catalog or "", "Slot": mod.slot or "", "Direcao": ch.direction, "Canal": ch.path,
                "Descricao": ch.description or "", "Unidade": ch.engineering_unit or "", "Min": "" if ch.min is None else ch.min, "Max": "" if ch.max is None else ch.max,
                "Spare": "sim" if ch.spare else "", "Usos": len(hits), "Rotinas": ", ".join(where[:6]) + (" …" if len(where) > 6 else ""),
            })
    return rows


def channel_summary(rows: List[Dict[str, Any]]) -> Dict[str, int]:
    return {
        "canais": len(rows),
        "spare": sum(1 for r in rows if r["Spare"]),
        "sem_uso": sum(1 for r in rows if not r["Spare"] and r["Usos"] == 0),
        "sem_descricao": sum(1 for r in rows if not r["Descricao"]),
        "modulos": len({r["Modulo"] for r in rows}),
    }


def instruments(ctrl: Controller, xref: Optional[CrossRef] = None, include_all: bool = False) -> List[Dict[str, Any]]:
    xref = xref or build_xref(ctrl)
    source, sink = io_source_map(ctrl)
    chan_desc = {ch.path.lower(): ch.description or "" for m in ctrl.modules for ch in m.channels}
    rows: List[Dict[str, Any]] = []
    for aoi in ctrl.aois:
        call_params = [p.name for p in aoi.call_parameters]
        usages = {p.name.lower(): p.usage for p in aoi.parameters}
        for tag in xref.instances_of(aoi.name):
            if tag.scope.startswith(AOI_PREFIX):
                continue
            values = flatten(tag.values) if tag.values is not None else {}
            rmin = rmax = ""
            for member, val in values.items():
                if "." in member and member.count(".") > 1:
                    continue
                kind = range_kind(member)
                if kind == "min" and rmin == "":
                    rmin = val
                elif kind == "max" and rmax == "":
                    rmax = val
            io_in: List[str] = []
            io_out: List[str] = []
            sites = xref.calls_of_instance(tag)
            for s in sites:
                c = s.call
                if c.source == "fbd":
                    items = list(c.pins.items())
                else:
                    items = [(call_params[i] if i < len(call_params) else f"arg{i}", a) for i, a in enumerate(c.arguments)]
                for pin, val in items:
                    usage = usages.get(pin.lower(), "Input")
                    low = val.lower()
                    if _is_io(val):
                        (io_out if usage == "Output" else io_in).append(val)
                    elif usage == "Output" and low in sink:
                        io_out.append(sink[low])
                    elif usage != "Output" and low in source:
                        io_in.append(source[low])
            io_in = list(dict.fromkeys(io_in))
            io_out = list(dict.fromkeys(io_out))
            looks_like_device = bool(tag.engineering_unit or rmin != "" or rmax != "" or io_in or io_out)
            if not (include_all or looks_like_device):
                continue
            rows.append({
                "Tag": tag.name, "Tipo": aoi.name, "Descricao": " ".join((tag.description or "").split()), "Unidade": tag.engineering_unit or "",
                "Area": tag.area or "", "Label": tag.label or "", "Escopo": "controlador" if tag.scope == "controller" else tag.scope,
                "RangeMin": rmin, "RangeMax": rmax, "EntradaIO": ", ".join(io_in), "EntradaIODescricao": ", ".join(chan_desc.get(p.lower(), "") for p in io_in),
                "SaidaIO": ", ".join(io_out), "Chamada": ", ".join(f"{s.scope}/{s.routine} {s.call.location}" for s in sites[:2]),
            })
    rows.sort(key=lambda r: (r["Tipo"].lower(), r["Tag"].lower()))
    return rows


def io_tables(ctrl: Controller, xref: Optional[CrossRef] = None, include_all: bool = False) -> List[Tuple[str, List[str], List[Dict[str, Any]]]]:
    xref = xref or build_xref(ctrl)
    return [("Instrumentos", INSTRUMENT_HEADER, instruments(ctrl, xref, include_all)), ("Canais_IO", CHANNEL_HEADER, io_channels(ctrl, xref))]

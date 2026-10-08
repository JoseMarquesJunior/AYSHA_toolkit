"""L5X (Studio 5000 XML export) -> l5x_core.model.

Tolerant by design: unknown elements are ignored, a malformed rung becomes an
info Finding instead of aborting, protected (EncodedData) content is flagged
and discarded, never decoded.
"""

from __future__ import annotations

import io
from typing import Any, Dict, List, Optional, Set, Tuple

from lxml import etree

from .model import (
    Aoi, AoiParameter, Controller, DataType, DataTypeMember, Finding, Module,
    Program, Routine, Rung, STLine, Tag, TagRef, Task, READ, WRITE, READWRITE,
)
from .rung_text import parse_rung, RungSyntaxError, split_tag_path
from .st_text import extract_st_refs


class L5XError(Exception):
    """Raised when the file is not a usable L5X export."""


_PARSER = etree.XMLParser(huge_tree=True, remove_blank_text=True, resolve_entities=False, no_network=True)

# Structures whose values we keep (for the "timer without preset" rule).
_PRESET_BEARING_BASE = {"TIMER", "COUNTER"}


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def _text(el: Optional[etree._Element], child: str) -> Optional[str]:
    if el is None:
        return None
    c = el.find(child)
    if c is None or c.text is None:
        return None
    t = c.text.strip()
    return t or None


def _bool(value: Optional[str]) -> bool:
    return (value or "").strip().lower() == "true"


def _value(raw: Optional[str]) -> Any:
    if raw is None:
        return None
    s = raw.strip()
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        return s


def _decorated_to_dict(el: etree._Element) -> Any:
    """Convert a Decorated <Data> subtree into nested dict/list values."""
    tag = el.tag
    if tag in ("DataValueMember", "DataValue"):
        return _value(el.get("Value"))
    if tag in ("Structure", "StructureMember"):
        return {m.get("Name"): _decorated_to_dict(m) for m in el if m.get("Name")}
    if tag in ("Array", "ArrayMember"):
        return {e.get("Index"): _decorated_to_dict(e) for e in el if e.tag == "Element"}
    if tag == "Element":
        if el.get("Value") is not None:
            return _value(el.get("Value"))
        children = list(el)
        return _decorated_to_dict(children[0]) if children else None
    if tag == "Data":
        children = list(el)
        return _decorated_to_dict(children[0]) if children else None
    return None


def _preset_bearing_types(udts: List[DataType], aoi_names: Set[str]) -> Set[str]:
    """Data types that contain TIMER/COUNTER members (recursively)."""
    bearing = set(_PRESET_BEARING_BASE)
    by_name = {d.name.lower(): d for d in udts}
    changed = True
    while changed:
        changed = False
        for d in udts:
            if d.name in bearing:
                continue
            for m in d.members:
                if m.data_type in bearing or m.data_type.lower() in {b.lower() for b in bearing}:
                    bearing.add(d.name)
                    changed = True
                    break
    return bearing


# --------------------------------------------------------------------------
# Parser
# --------------------------------------------------------------------------
class _L5XParser:
    def __init__(self) -> None:
        self.findings: List[Finding] = []
        self.aoi_usages: Dict[str, List[str]] = {}
        self.bearing_types: Set[str] = set()

    # ---- entry -----------------------------------------------------------
    def parse_root(self, root: etree._Element) -> Controller:
        if root.tag != "RSLogix5000Content":
            raise L5XError(f"root element is <{root.tag}>, expected <RSLogix5000Content>")
        ctrl_el = root.find("Controller")
        if ctrl_el is None:
            raise L5XError("no <Controller> element found")

        ctrl = Controller(
            name=ctrl_el.get("Name") or root.get("TargetName") or "?",
            processor_type=ctrl_el.get("ProcessorType"),
            major_rev=ctrl_el.get("MajorRev"),
            minor_rev=ctrl_el.get("MinorRev"),
            software_revision=root.get("SoftwareRevision"),
            target_name=root.get("TargetName"),
            export_date=root.get("ExportDate"),
            description=_text(ctrl_el, "Description"),
        )
        safety_info = ctrl_el.find("SafetyInfo")
        ctrl.safety_controller = safety_info is not None and (len(safety_info) > 0 or bool(safety_info.attrib))

        ctrl.data_types = [self._datatype(d) for d in ctrl_el.iterfind("DataTypes/DataType")]
        ctrl.modules = [self._module(m) for m in ctrl_el.iterfind("Modules/Module")]

        # AOIs first: their parameter usages drive operand classification.
        aoi_els = list(ctrl_el.iterfind("AddOnInstructionDefinitions/AddOnInstructionDefinition"))
        aoi_els += list(ctrl_el.iterfind("AddOnInstructionDefinitions/EncodedData"))
        ctrl.aois = []
        for a in aoi_els:
            aoi = self._aoi_header(a)
            ctrl.aois.append(aoi)
            self.aoi_usages[aoi.name.lower()] = [p.usage for p in aoi.call_parameters]
        self.bearing_types = _preset_bearing_types(ctrl.data_types, {a.name for a in ctrl.aois})
        for a, aoi in zip(aoi_els, ctrl.aois):
            self._aoi_body(a, aoi)

        ctrl.tags = [self._tag(t, "controller") for t in ctrl_el.iterfind("Tags/Tag")]
        ctrl.programs = [self._program(p) for p in ctrl_el.iterfind("Programs/Program")]
        ctrl.tasks = [self._task(t) for t in ctrl_el.iterfind("Tasks/Task")]

        by_name = {p.name.lower(): p for p in ctrl.programs}
        for task in ctrl.tasks:
            for pname in task.programs:
                prog = by_name.get(pname.lower())
                if prog is not None and prog.task is None:
                    prog.task = task.name
                    if task.is_safety and not prog.class_:
                        prog.class_ = "Safety"
        ctrl.parse_findings = self.findings
        return ctrl

    # ---- leaf elements ---------------------------------------------------
    def _datatype(self, el: etree._Element) -> DataType:
        members = [
            DataTypeMember(
                name=m.get("Name") or "?",
                data_type=m.get("DataType") or "",
                dimension=int(m.get("Dimension") or 0),
                description=_text(m, "Description"),
                hidden=_bool(m.get("Hidden")),
            )
            for m in el.iterfind("Members/Member")
        ]
        return DataType(name=el.get("Name") or "?", family=el.get("Family"), class_=el.get("Class"), description=_text(el, "Description"), members=members)

    def _module(self, el: etree._Element) -> Module:
        return Module(
            name=el.get("Name") or "?",
            catalog=el.get("CatalogNumber"),
            parent=el.get("ParentModule"),
            description=_text(el, "Description"),
            inhibited=_bool(el.get("Inhibited")),
        )

    def _task(self, el: etree._Element) -> Task:
        return Task(
            name=el.get("Name") or "?",
            type=el.get("Type") or "?",
            rate=el.get("Rate"),
            priority=el.get("Priority"),
            class_=el.get("Class"),
            programs=[sp.get("Name") for sp in el.iterfind("ScheduledPrograms/ScheduledProgram") if sp.get("Name")],
        )

    def _tag(self, el: etree._Element, scope: str) -> Tag:
        data_type = el.get("DataType") or ""
        tag = Tag(
            name=el.get("Name") or "?",
            scope=scope,
            tag_type=el.get("TagType") or "Base",
            data_type=data_type,
            dimensions=el.get("Dimensions"),
            alias_for=el.get("AliasFor"),
            description=_text(el, "Description"),
            class_=el.get("Class"),
            usage=el.get("Usage"),
            constant=_bool(el.get("Constant")),
        )
        if data_type in self.bearing_types or data_type.upper() in self.bearing_types:
            for d in el.iterfind("Data"):
                if d.get("Format") == "Decorated":
                    try:
                        tag.values = _decorated_to_dict(d)
                    except Exception:  # pragma: no cover - defensive
                        tag.values = None
                    break
        return tag

    # ---- AOIs ------------------------------------------------------------
    def _aoi_header(self, el: etree._Element) -> Aoi:
        if el.tag == "EncodedData":
            return Aoi(name=el.get("Name") or "?", revision=el.get("Revision"), vendor=el.get("Vendor"), protected=True)
        params = [
            AoiParameter(
                name=p.get("Name") or "?",
                usage=p.get("Usage") or "Input",
                data_type=p.get("DataType") or "",
                required=_bool(p.get("Required")),
                visible=_bool(p.get("Visible")),
                description=_text(p, "Description"),
            )
            for p in el.iterfind("Parameters/Parameter")
        ]
        aoi = Aoi(
            name=el.get("Name") or "?",
            revision=el.get("Revision"),
            vendor=el.get("Vendor"),
            description=_text(el, "Description"),
            parameters=params,
            class_=el.get("Class"),
        )
        if el.find("EncodedData") is not None:
            aoi.protected = True
        return aoi

    def _aoi_body(self, el: etree._Element, aoi: Aoi) -> None:
        if aoi.protected:
            return
        scope = f"aoi:{aoi.name}"
        aoi.local_tags = [self._tag(t, scope) for t in el.iterfind("LocalTags/LocalTag")]
        for p in el.iterfind("Parameters/Parameter"):
            # parameters are also tags inside the AOI scope
            aoi.local_tags.append(
                Tag(name=p.get("Name") or "?", scope=scope, data_type=p.get("DataType") or "", usage=p.get("Usage"), description=_text(p, "Description"), alias_for=p.get("AliasFor"), tag_type="Alias" if p.get("AliasFor") else "Base")
            )
        aoi.routines = [self._routine(r, scope) for r in el.iterfind("Routines/Routine")]

    # ---- programs & routines --------------------------------------------
    def _program(self, el: etree._Element) -> Program:
        name = el.get("Name") or "?"
        prog = Program(
            name=name,
            main_routine=el.get("MainRoutineName"),
            fault_routine=el.get("FaultRoutineName"),
            disabled=_bool(el.get("Disabled")),
            class_=el.get("Class"),
            description=_text(el, "Description"),
        )
        prog.tags = [self._tag(t, name) for t in el.iterfind("Tags/Tag")]
        prog.routines = [self._routine(r, name) for r in el.iterfind("Routines/Routine")]
        return prog

    def _routine(self, el: etree._Element, scope: str) -> Routine:
        rtype = el.get("Type") or "?"
        routine = Routine(name=el.get("Name") or "?", type=rtype, description=_text(el, "Description"))
        if el.find("EncodedData") is not None:
            routine.protected = True
            return routine
        if rtype == "RLL":
            content = el.find("RLLContent")
            if content is not None:
                self._rll(content, routine, scope)
        elif rtype == "ST":
            content = el.find("STContent")
            if content is not None:
                self._st(content, routine)
        elif rtype in ("FBD", "SFC"):
            content = el.find(f"{rtype}Content")
            if content is not None:
                self._graphic(content, routine, rtype.lower())
        return routine

    def _rll(self, content: etree._Element, routine: Routine, scope: str) -> None:
        program = None if scope.startswith("aoi:") else scope
        for r in content.iterfind("Rung"):
            number = int(r.get("Number") or len(routine.rungs))
            text = (_text(r, "Text") or "")
            rung = Rung(number=number, type=r.get("Type") or "N", comment=_text(r, "Comment"), text=text)
            try:
                rung.instructions, rung.max_branch_depth = parse_rung(text, self.aoi_usages)
            except (RungSyntaxError, ValueError, IndexError) as exc:
                rung.parse_error = str(exc)
                self.findings.append(
                    Finding(
                        rule="rung_parse_error", severity="info",
                        message=f"Rung não interpretado ({exc}); tratado como opaco.",
                        program=program, routine=routine.name, rung=number, evidence=text[:300],
                    )
                )
            for ins in rung.instructions:
                if ins.name.upper() in ("JSR", "FOR") and ins.operands:
                    routine.calls.append(ins.operands[0].raw)
            routine.rungs.append(rung)
        routine.size = len(routine.rungs)

    def _st(self, content: etree._Element, routine: Routine) -> None:
        for ln in content.iterfind("Line"):
            routine.st_lines.append(STLine(number=int(ln.get("Number") or len(routine.st_lines)), text=ln.text or ""))
        routine.size = len(routine.st_lines)
        try:
            routine.tag_refs, routine.calls = extract_st_refs(routine.st_lines, self.aoi_usages)
        except Exception as exc:  # pragma: no cover - defensive
            self.findings.append(Finding(rule="st_parse_error", severity="info", message=f"Rotina ST não interpretada ({exc}).", routine=routine.name))

    def _graphic(self, content: etree._Element, routine: Routine, source: str) -> None:
        """FBD/SFC: record size and approximate tag references, nothing more."""
        nodes = 0
        sheet_no = 0
        for node in content.iter():
            nodes += 1
            tag = node.tag
            if tag == "Sheet":
                sheet_no += 1
                continue
            loc = f"sheet {sheet_no}" if sheet_no else source
            if tag == "IRef":
                self._graphic_ref(routine, node.get("Operand"), READ, loc, source)
            elif tag == "ORef":
                self._graphic_ref(routine, node.get("Operand"), WRITE, loc, source)
            elif tag in ("Block", "AddOnInstruction", "Action", "Step", "Transition"):
                self._graphic_ref(routine, node.get("Operand"), READWRITE, loc, source)
            elif tag == "InOutParameter":
                self._graphic_ref(routine, node.get("Argument"), READWRITE, loc, source)
            elif tag == "JSR" and node.get("Routine"):
                routine.calls.append(node.get("Routine"))
            elif tag == "STContent":
                lines = [STLine(number=int(l.get("Number") or i), text=l.text or "") for i, l in enumerate(node.iterfind("Line"))]
                try:
                    refs, calls = extract_st_refs(lines, self.aoi_usages)
                    for ref in refs:
                        ref.source = source
                        ref.location = f"{loc} {ref.location}"
                    routine.tag_refs.extend(refs)
                    routine.calls.extend(calls)
                except Exception:  # pragma: no cover - defensive
                    pass
        routine.size = nodes
        routine.sheets = sheet_no

    def _graphic_ref(self, routine: Routine, operand: Optional[str], access: str, loc: str, source: str) -> None:
        if not operand:
            return
        try:
            prog, base, rest = split_tag_path(operand.strip())
        except ValueError:
            return  # literal or expression
        routine.tag_refs.append(TagRef(tag_base=base, tag_path=base + rest, access=access, location=loc, source=source, program=prog))


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------
def parse_file(path: str) -> Controller:
    try:
        tree = etree.parse(path, _PARSER)
    except (etree.XMLSyntaxError, OSError) as exc:
        raise L5XError(f"não foi possível ler o L5X: {exc}") from exc
    try:
        return _L5XParser().parse_root(tree.getroot())
    finally:
        del tree


def parse_bytes(data: bytes) -> Controller:
    try:
        tree = etree.parse(io.BytesIO(data), _PARSER)
    except etree.XMLSyntaxError as exc:
        raise L5XError(f"não foi possível ler o L5X: {exc}") from exc
    return _L5XParser().parse_root(tree.getroot())

"""Cross references: tag usages (who reads / writes what, where) and the
routine call graph.

Resolution order for an operand inside a routine of program P:
``\\Other.Tag`` -> program Other only; otherwise program P tags, then
controller tags. Routines of an AOI only see the AOI parameters and local
tags. Aliases are recorded on the alias *and* on the target tag.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Tuple

from .model import Controller, Program, Routine, Tag, NONE, READ, WRITE, READWRITE
from .rung_text import split_tag_path

CONTROLLER = "controller"


def tag_key(scope: str, name: str) -> str:
    return f"{scope.lower()}|{name.lower()}"


@dataclass
class Usage:
    scope: str  # program name or "aoi:Name"
    routine: str
    location: str  # "rung 3", "line 12", "sheet 1"
    instruction: str
    access: str  # read | write | readwrite
    path: str  # path as written in the code
    rung: Optional[int] = None
    approximate: bool = False
    branch_depth: int = 0

    @property
    def reads(self) -> bool:
        return self.access in (READ, READWRITE)

    @property
    def writes(self) -> bool:
        return self.access in (WRITE, READWRITE)

    def where(self) -> str:
        return f"{self.scope} / {self.routine} / {self.location}"


@dataclass
class CrossRef:
    tags: Dict[str, Tag] = field(default_factory=dict)  # key -> Tag
    usages: Dict[str, List[Usage]] = field(default_factory=lambda: defaultdict(list))
    io_usages: Dict[str, List[Usage]] = field(default_factory=lambda: defaultdict(list))  # module name -> usages
    unresolved: Dict[Tuple[str, str], List[str]] = field(default_factory=lambda: defaultdict(list))  # (scope, routine) -> names
    calls: Dict[Tuple[str, str], List[str]] = field(default_factory=lambda: defaultdict(list))  # (program, routine) -> callee names (as written)
    callers: Dict[Tuple[str, str], List[Tuple[str, str]]] = field(default_factory=lambda: defaultdict(list))  # (program, routine.lower()) -> [(program, routine)]
    missing_calls: List[Tuple[str, str, str]] = field(default_factory=list)  # (program, routine, target)
    modules: Dict[str, str] = field(default_factory=dict)  # lower-cased module name -> name

    # ---- resolution ------------------------------------------------------
    def resolve(self, scope: str, base: str, program: Optional[str] = None) -> Optional[Tag]:
        if program:
            return self.tags.get(tag_key(program, base))
        tag = self.tags.get(tag_key(scope, base))
        if tag is None and not scope.startswith("aoi:"):
            tag = self.tags.get(tag_key(CONTROLLER, base))
        return tag

    def usages_of(self, tag: Tag) -> List[Usage]:
        return self.usages.get(tag_key(tag.scope, tag.name), [])

    def reads_of(self, tag: Tag) -> List[Usage]:
        return [u for u in self.usages_of(tag) if u.reads]

    def writes_of(self, tag: Tag) -> List[Usage]:
        return [u for u in self.usages_of(tag) if u.writes]

    def callers_of(self, program: str, routine: str) -> List[Tuple[str, str]]:
        return self.callers.get((program, routine.lower()), [])

    # ---- recording -------------------------------------------------------
    def _record(self, scope: str, base: str, program: Optional[str], usage: Usage, routine: str) -> None:
        if base.lower() in ("this",):
            return
        if ":" in base:
            module = base.split(":", 1)[0]
            self.io_usages[module.lower()].append(usage)
            return
        tag = self.resolve(scope, base, program)
        if tag is None:
            if not program and base.lower() in self.modules:
                # module reference (e.g. InOut parameter of a module-status AOI)
                self.io_usages[base.lower()].append(usage)
                return
            self.unresolved[(scope, routine)].append(base if not program else f"\\{program}.{base}")
            return
        self.usages[tag_key(tag.scope, tag.name)].append(usage)
        if tag.alias_for:
            try:
                aprog, abase, arest = split_tag_path(tag.alias_for)
            except ValueError:
                return
            if ":" in abase:
                self.io_usages[abase.split(":", 1)[0].lower()].append(usage)
                return
            target = self.resolve(tag.scope if tag.scope != CONTROLLER else CONTROLLER, abase, aprog)
            if target is not None and target is not tag:
                self.usages[tag_key(target.scope, target.name)].append(usage)


def _index_tags(ctrl: Controller, xref: CrossRef) -> None:
    for t in ctrl.tags:
        xref.tags[tag_key(CONTROLLER, t.name)] = t
    for p in ctrl.programs:
        for t in p.tags:
            xref.tags[tag_key(p.name, t.name)] = t
    for a in ctrl.aois:
        for t in a.local_tags:
            xref.tags[tag_key(f"aoi:{a.name}", t.name)] = t


def _index_routine(scope: str, routine: Routine, xref: CrossRef) -> None:
    if routine.protected:
        return
    for rung in routine.rungs:
        loc = f"rung {rung.number}"
        for ins in rung.instructions:
            for op in ins.operands:
                if op.kind == "tag":
                    access = op.access if op.access != NONE else READ
                    usage = Usage(scope=scope, routine=routine.name, location=loc, instruction=ins.name, access=access, path=op.tag_path or op.raw, rung=rung.number, branch_depth=ins.branch_depth)
                    xref._record(scope, op.tag_base or "", op.program, usage, routine.name)
                for extra in op.extra_tags:
                    try:
                        prog, base, rest = split_tag_path(extra)
                    except ValueError:
                        continue
                    usage = Usage(scope=scope, routine=routine.name, location=loc, instruction=ins.name, access=READ, path=base + rest, rung=rung.number, approximate=True, branch_depth=ins.branch_depth)
                    xref._record(scope, base, prog, usage, routine.name)
    for ref in routine.tag_refs:
        usage = Usage(scope=scope, routine=routine.name, location=ref.location, instruction=ref.source.upper(), access=ref.access, path=ref.tag_path, approximate=True)
        xref._record(scope, ref.tag_base, ref.program, usage, routine.name)


def _index_calls(prog: Program, xref: CrossRef) -> None:
    names = {r.name.lower(): r.name for r in prog.routines}
    for routine in prog.routines:
        for target in routine.calls:
            target = target.strip()
            xref.calls[(prog.name, routine.name)].append(target)
            real = names.get(target.lower())
            if real is None:
                xref.missing_calls.append((prog.name, routine.name, target))
            else:
                xref.callers[(prog.name, real.lower())].append((prog.name, routine.name))


def build_xref(ctrl: Controller) -> CrossRef:
    xref = CrossRef()
    xref.modules = {m.name.lower(): m.name for m in ctrl.modules}
    _index_tags(ctrl, xref)
    for prog in ctrl.programs:
        for routine in prog.routines:
            _index_routine(prog.name, routine, xref)
        _index_calls(prog, xref)
    for aoi in ctrl.aois:
        for routine in aoi.routines:
            _index_routine(f"aoi:{aoi.name}", routine, xref)
    return xref

"""Dataclasses describing a parsed L5X project.

Every structure here is plain Python (no lxml references) so it can be
serialized, cached and reused by other front-ends.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional


# Access kinds used throughout
READ = "read"
WRITE = "write"
READWRITE = "readwrite"
NONE = "none"

SEVERITIES = ("error", "warning", "info")


@dataclass
class Operand:
    raw: str
    kind: str  # "tag" | "literal" | "unused" | "expression" | "keyword"
    access: str = READ
    tag_base: Optional[str] = None  # base tag name, up to the first dot or bracket
    tag_path: Optional[str] = None  # full path as written, without the program prefix
    program: Optional[str] = None  # set for cross-program references
    extra_tags: List[str] = field(default_factory=list)  # tags found in expressions / indices

    @property
    def is_tag(self) -> bool:
        return self.kind == "tag"


@dataclass
class Instruction:
    name: str
    operands: List[Operand]
    index: int  # position inside the rung (0-based)
    branch_depth: int  # 0 = main line
    known: bool = True  # False -> unknown mnemonic, all operands treated as reads
    is_aoi: bool = False

    @property
    def in_branch(self) -> bool:
        return self.branch_depth > 0


@dataclass
class Rung:
    number: int
    type: str
    comment: Optional[str]
    text: str
    instructions: List[Instruction] = field(default_factory=list)
    max_branch_depth: int = 0
    parse_error: Optional[str] = None


@dataclass
class STLine:
    number: int
    text: str


@dataclass
class TagRef:
    """Approximate tag reference extracted from ST or FBD/SFC content."""

    tag_base: str
    tag_path: str
    access: str
    location: str  # e.g. "line 12" or "sheet 1"
    source: str  # "st" | "fbd" | "sfc"
    program: Optional[str] = None


@dataclass
class Routine:
    name: str
    type: str  # RLL | ST | FBD | SFC | other
    description: Optional[str] = None
    protected: bool = False
    rungs: List[Rung] = field(default_factory=list)
    st_lines: List[STLine] = field(default_factory=list)
    tag_refs: List[TagRef] = field(default_factory=list)  # approximate refs (ST/FBD/SFC)
    calls: List[str] = field(default_factory=list)  # JSR / FOR targets (routine names)
    size: int = 0  # rungs, lines or graphic nodes depending on type
    sheets: int = 0  # FBD only

    @property
    def analyzed(self) -> bool:
        return self.type in ("RLL", "ST") and not self.protected


@dataclass
class Tag:
    name: str
    scope: str  # "controller" | "<program name>" | "aoi:<aoi name>"
    tag_type: str = "Base"  # Base | Alias | Produced | Consumed
    data_type: str = ""
    dimensions: Optional[str] = None
    alias_for: Optional[str] = None
    description: Optional[str] = None
    class_: Optional[str] = None  # "Standard" | "Safety" | None
    usage: Optional[str] = None  # program parameters / AOI params: Input | Output | InOut | Public | Local
    constant: bool = False
    values: Optional[Dict[str, Any]] = None  # decorated values (kept only for timer/counter bearing tags)

    @property
    def is_io(self) -> bool:
        return ":" in self.name

    @property
    def is_safety(self) -> bool:
        return (self.class_ or "").lower() == "safety"


@dataclass
class DataTypeMember:
    name: str
    data_type: str
    dimension: int = 0
    description: Optional[str] = None
    hidden: bool = False


@dataclass
class DataType:
    name: str
    family: Optional[str] = None
    class_: Optional[str] = None
    description: Optional[str] = None
    members: List[DataTypeMember] = field(default_factory=list)


@dataclass
class AoiParameter:
    name: str
    usage: str  # Input | Output | InOut
    data_type: str = ""
    required: bool = False
    visible: bool = False
    description: Optional[str] = None


@dataclass
class Aoi:
    name: str
    revision: Optional[str] = None
    vendor: Optional[str] = None
    description: Optional[str] = None
    protected: bool = False
    parameters: List[AoiParameter] = field(default_factory=list)
    local_tags: List[Tag] = field(default_factory=list)
    routines: List[Routine] = field(default_factory=list)
    class_: Optional[str] = None

    @property
    def call_parameters(self) -> List[AoiParameter]:
        """Parameters that appear as operands in a rung call (Required or InOut)."""
        return [p for p in self.parameters if p.required or p.usage == "InOut"]


@dataclass
class Module:
    name: str
    catalog: Optional[str] = None
    parent: Optional[str] = None
    description: Optional[str] = None
    inhibited: bool = False


@dataclass
class Program:
    name: str
    main_routine: Optional[str] = None
    fault_routine: Optional[str] = None
    disabled: bool = False
    class_: Optional[str] = None
    description: Optional[str] = None
    tags: List[Tag] = field(default_factory=list)
    routines: List[Routine] = field(default_factory=list)
    task: Optional[str] = None  # task that schedules it, filled by the parser

    @property
    def is_safety(self) -> bool:
        return (self.class_ or "").lower() == "safety"

    def routine(self, name: str) -> Optional[Routine]:
        low = name.lower()
        for r in self.routines:
            if r.name.lower() == low:
                return r
        return None


@dataclass
class Task:
    name: str
    type: str
    rate: Optional[str] = None
    priority: Optional[str] = None
    class_: Optional[str] = None
    programs: List[str] = field(default_factory=list)

    @property
    def is_safety(self) -> bool:
        return (self.class_ or "").lower() == "safety"


@dataclass
class Finding:
    rule: str
    severity: str  # error | warning | info
    message: str
    program: Optional[str] = None
    routine: Optional[str] = None
    rung: Optional[int] = None
    tag: Optional[str] = None
    evidence: Optional[str] = None
    safety: bool = False  # touches Class="Safety" content: explain only, never suggest changes

    def location(self) -> str:
        parts = [p for p in (self.program, self.routine) if p]
        loc = " / ".join(parts)
        if self.rung is not None:
            loc = f"{loc} / Rung {self.rung}" if loc else f"Rung {self.rung}"
        return loc or "-"


@dataclass
class Metrics:
    programs: int = 0
    routines: int = 0
    routines_by_type: Dict[str, int] = field(default_factory=dict)
    rungs: int = 0
    st_lines: int = 0
    instructions: int = 0
    tags_controller: int = 0
    tags_program: int = 0
    aois: int = 0
    aois_protected: int = 0
    udts: int = 0
    modules: int = 0
    tasks: int = 0
    largest_routine: Optional[str] = None  # "Program/Routine"
    largest_routine_rungs: int = 0
    max_branch_depth: int = 0
    commented_rung_pct: float = 0.0
    protected_routines: int = 0
    safety_programs: int = 0
    unknown_instructions: Dict[str, int] = field(default_factory=dict)
    findings_by_severity: Dict[str, int] = field(default_factory=dict)


@dataclass
class Controller:
    name: str
    processor_type: Optional[str] = None
    major_rev: Optional[str] = None
    minor_rev: Optional[str] = None
    software_revision: Optional[str] = None
    target_name: Optional[str] = None
    export_date: Optional[str] = None
    description: Optional[str] = None
    safety_controller: bool = False
    data_types: List[DataType] = field(default_factory=list)
    modules: List[Module] = field(default_factory=list)
    aois: List[Aoi] = field(default_factory=list)
    tags: List[Tag] = field(default_factory=list)
    programs: List[Program] = field(default_factory=list)
    tasks: List[Task] = field(default_factory=list)
    parse_findings: List[Finding] = field(default_factory=list)  # rung parse errors etc.

    def program(self, name: str) -> Optional[Program]:
        low = name.lower()
        for p in self.programs:
            if p.name.lower() == low:
                return p
        return None

    def aoi(self, name: str) -> Optional[Aoi]:
        low = name.lower()
        for a in self.aois:
            if a.name.lower() == low:
                return a
        return None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

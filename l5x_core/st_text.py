"""Approximate tag/call extraction from Structured Text lines.

This is intentionally regex based: it finds identifiers that look like tag
paths, marks assignment targets as writes and everything else as reads, and
collects JSR / AOI calls. Results are flagged as approximate downstream.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Sequence, Tuple

from .model import STLine, TagRef, READ, WRITE, READWRITE
from .rung_text import split_tag_path

_ST_KEYWORDS = {
    "IF", "THEN", "ELSE", "ELSIF", "END_IF", "CASE", "OF", "END_CASE", "FOR", "TO",
    "BY", "DO", "END_FOR", "WHILE", "END_WHILE", "REPEAT", "UNTIL", "END_REPEAT",
    "EXIT", "RETURN", "TRUE", "FALSE", "AND", "OR", "XOR", "NOT", "MOD", "AND_THEN",
    "OR_ELSE", "THIS", "PROGRAM", "CONTROLLER", "TASK", "ROUTINE", "MODULE",
}

_LINE_COMMENT_RE = re.compile(r"//[^\n]*")
_BLOCK_COMMENT_RE = re.compile(r"\(\*.*?\*\)|/\*.*?\*/", re.S)
_STRING_RE = re.compile(r"'(?:[^']|'')*'")
_RADIX_RE = re.compile(r"\b\d+#[0-9A-Za-z_]+")
_IDENT_RE = re.compile(
    r"(?<![\w.#\\])\\?[A-Za-z_]\w*(?::[A-Za-z0-9_]+)*(?:\.\w+|\[[^\]\[]*\])*"
)
_CALL_RE = re.compile(r"\b([A-Za-z_]\w*)\s*\(")
_ASSIGN_RE = re.compile(r"^\s*(\\?[A-Za-z_][\w:.\[\]]*)\s*:=", re.M)


def _call_args(text: str, start: int) -> List[Tuple[str, int]]:
    """Split the argument list starting right after '(' at ``start``.

    Returns [(stripped_arg, offset_of_arg_in_text), ...] up to the matching ')'.
    """
    args: List[Tuple[str, int]] = []
    depth = 0
    cur_start = start
    i = start
    n = len(text)
    while i < n:
        ch = text[i]
        if ch in "([":
            depth += 1
        elif ch in ")]":
            if depth == 0:
                break
            depth -= 1
        elif ch == "," and depth == 0:
            raw = text[cur_start:i]
            stripped = raw.strip()
            if stripped:
                args.append((stripped, cur_start + (len(raw) - len(raw.lstrip()))))
            cur_start = i + 1
        i += 1
    raw = text[cur_start:i]
    stripped = raw.strip()
    if stripped:
        args.append((stripped, cur_start + (len(raw) - len(raw.lstrip()))))
    return args


def _blank_out(text: str, pattern: re.Pattern) -> str:
    """Replace matches with spaces (keeping newlines) so positions stay stable."""

    def repl(m: re.Match) -> str:
        return re.sub(r"[^\n]", " ", m.group(0))

    return pattern.sub(repl, text)


def extract_st_refs(
    lines: Sequence[STLine],
    aoi_usages: Optional[Dict[str, List[str]]] = None,
) -> Tuple[List[TagRef], List[str]]:
    """Return (tag references, called routine names) for an ST routine."""
    if not lines:
        return [], []
    numbers = [ln.number for ln in lines]
    text = "\n".join(ln.text for ln in lines)
    text = _blank_out(text, _BLOCK_COMMENT_RE)
    text = _blank_out(text, _LINE_COMMENT_RE)
    text = _blank_out(text, _STRING_RE)
    text = _blank_out(text, _RADIX_RE)

    # map character offset -> line number
    line_starts = [0]
    for i, ch in enumerate(text):
        if ch == "\n":
            line_starts.append(i + 1)

    def line_of(pos: int) -> int:
        lo, hi = 0, len(line_starts) - 1
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if line_starts[mid] <= pos:
                lo = mid
            else:
                hi = mid - 1
        return numbers[lo] if lo < len(numbers) else numbers[-1]

    calls: List[str] = []
    writes: Dict[Tuple[str, int], bool] = {}
    access_at: Dict[int, str] = {}  # offset of an AOI call argument -> access
    skip: set = set()  # offsets of function names and routine-name arguments

    for m in _CALL_RE.finditer(text):
        name = m.group(1)
        skip.add(m.start(1))
        args = _call_args(text, m.end())
        upper = name.upper()
        if upper in ("JSR", "FOR"):
            if args:
                calls.append(args[0][0])
                skip.add(args[0][1])
        elif aoi_usages is not None and name.lower() in aoi_usages:
            usages = aoi_usages[name.lower()]
            for i, (arg, offset) in enumerate(args):
                if i == 0:
                    access_at[offset] = READWRITE
                elif i - 1 < len(usages):
                    access_at[offset] = {"Input": READ, "Output": WRITE, "InOut": READWRITE}.get(usages[i - 1], READWRITE)

    for m in _ASSIGN_RE.finditer(text):
        writes[(m.group(1), m.start(1))] = True

    refs: List[TagRef] = []
    for m in _IDENT_RE.finditer(text):
        ident = m.group(0)
        if m.start() in skip:
            continue
        base_name = re.match(r"\\?([A-Za-z_]\w*)", ident).group(1)
        if base_name.upper() in _ST_KEYWORDS:
            continue
        try:
            prog, base, rest = split_tag_path(ident)
        except ValueError:
            continue
        if (ident, m.start()) in writes:
            access = WRITE
        elif m.start() in access_at:
            access = access_at[m.start()]
        else:
            access = READ
        refs.append(TagRef(tag_base=base, tag_path=base + rest, access=access, location=f"line {line_of(m.start())}", source="st", program=prog))
    return refs, calls

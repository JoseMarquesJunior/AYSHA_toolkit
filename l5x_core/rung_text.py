"""Tokenizer and classifier for the neutral rung text found in RLL routines.

Grammar (as observed in Studio 5000 exports):

    rung      := element* ';'
    element   := instruction | branch
    branch    := '[' leg (',' leg)* ']'
    leg       := element*
    instruction := NAME '(' operand (',' operand)* ')'

Operands are tag paths (``Motor``, ``Tank.Level``, ``Alarms[3].Active``,
``Local:1:I.Data.5``, ``\\Prog.Tag``), literals (``10``, ``1.5``, ``16#FF``,
``'text'``), expressions (``A*2+B``) or ``?`` (value lives in the tag data).
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Sequence, Tuple

from .model import Instruction, Operand, READ, WRITE, READWRITE, NONE


class RungSyntaxError(ValueError):
    pass


# --------------------------------------------------------------------------
# Operand roles per instruction (positional). Missing positions default to
# READ. "k" marks keyword operands (routine names, labels, class names) that
# are not tags.
# --------------------------------------------------------------------------
R, W, RW, K = READ, WRITE, READWRITE, "k"

ROLES: Dict[str, Sequence[str]] = {
    # bit
    "XIC": (R,), "XIO": (R,), "ONS": (RW,), "OSR": (RW, W), "OSF": (RW, W),
    "OTE": (W,), "OTL": (W,), "OTU": (W,),
    # timers / counters
    "TON": (RW, R, R), "TOF": (RW, R, R), "RTO": (RW, R, R),
    "CTU": (RW, R, R), "CTD": (RW, R, R), "RES": (W,),
    # move / logical
    "MOV": (R, W), "MVM": (R, R, W), "COP": (R, W, R), "CPS": (R, W, R),
    "FLL": (R, W, R), "CLR": (W,), "BTD": (R, R, W, R, R), "SWPB": (R, K, W),
    "AND": (R, R, W), "OR": (R, R, W), "XOR": (R, R, W), "NOT": (R, W),
    # math
    "ADD": (R, R, W), "SUB": (R, R, W), "MUL": (R, R, W), "DIV": (R, R, W),
    "MOD": (R, R, W), "XPY": (R, R, W), "NEG": (R, W), "ABS": (R, W),
    "SQR": (R, W), "SQRT": (R, W), "SIN": (R, W), "COS": (R, W), "TAN": (R, W),
    "ASN": (R, W), "ACS": (R, W), "ATN": (R, W), "LN": (R, W), "LOG": (R, W),
    "TRN": (R, W), "TRUNC": (R, W), "DEG": (R, W), "RAD": (R, W),
    "CPT": (W, R),
    # compare
    "EQU": (R, R), "NEQ": (R, R), "GRT": (R, R), "GEQ": (R, R), "LES": (R, R),
    "LEQ": (R, R), "LIM": (R, R, R), "MEQ": (R, R, R), "CMP": (R,),
    # program control
    "JSR": (K, R),  # special-cased: inputs read, returns written
    "SBR": (),  # all operands written (parameters received)
    "RET": (),  # all operands read
    "JMP": (K,), "LBL": (K,), "FOR": (K, W, R, R, R), "BRK": (), "TND": (),
    "NOP": (), "AFI": (), "MCR": (), "UID": (), "UIE": (), "EOT": (),
    "SFR": (K, K), "SFP": (K, K), "EVENT": (K,), "IOT": (W,),
    # system / comms
    "MSG": (RW,), "GSV": (K, K, K, W), "SSV": (K, K, K, R),
    # arrays / files
    "BSL": (W, RW, R, R), "BSR": (W, RW, R, R),
    "FFL": (R, W, RW, R, R), "FFU": (R, W, RW, R, R),
    "LFL": (R, W, RW, R, R), "LFU": (R, W, RW, R, R),
    "FAL": (RW, R, R, K, W, R), "FSC": (RW, R, R, K, R),
    "SIZE": (R, R, W), "DTR": (R, R, R), "SRT": (RW, R, K, RW), "STD": (RW, R, K, W), "AVE": (RW, R, K, W),
    "PID": (RW, R, R, W, R, R, R),
    # strings
    "CONCAT": (R, R, W), "LOWER": (R, W), "UPPER": (R, W), "DTOS": (R, W),
    "STOD": (R, W), "RTOS": (R, W), "STOR": (R, W), "MID": (R, R, R, W),
    "DELETE": (R, R, R, W), "INSERT": (R, R, R, W), "FIND": (R, R, R, W),
}

# Instructions whose operand 0 is a timer/counter structure and operand 1 a preset.
TIMER_COUNTER = {"TON", "TOF", "RTO", "CTU", "CTD"}

_EXPR_KEYWORDS = {
    "MOD", "AND", "OR", "XOR", "NOT", "ABS", "SQR", "SQRT", "SIN", "COS", "TAN",
    "ASN", "ACS", "ATN", "LN", "LOG", "TRN", "TRUNC", "DEG", "RAD", "FRD", "TOD",
    "NEG", "XPY",
}

_NUMBER_RE = re.compile(r"^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$")
_RADIX_RE = re.compile(r"^\d+#[0-9A-Za-z_]+$")
_TAG_RE = re.compile(
    r"^(?:\\(?P<prog>[A-Za-z_]\w*)\.)?"
    r"(?P<base>[A-Za-z_]\w*(?::[A-Za-z0-9_]+)*)"
    r"(?P<rest>(?:\.\w+|\[[^\]]*\])*)$"
)
_IDENT_RE = re.compile(
    r"(?<![\w.#\\])\\?[A-Za-z_]\w*(?::[A-Za-z0-9_]+)*(?:\.\w+|\[[^\]]*\])*"
)
_STRIP_LITERALS_RE = re.compile(r"'[^']*'|\"[^\"]*\"|\d+#[0-9A-Za-z_]+")
_NAME_RE = re.compile(r"[A-Za-z_]\w*")


# --------------------------------------------------------------------------
# Operand helpers
# --------------------------------------------------------------------------
def split_tag_path(path: str) -> Tuple[Optional[str], str, str]:
    """Return (program, base, rest) for a tag path. Raises ValueError if not a path."""
    m = _TAG_RE.match(path)
    if not m:
        raise ValueError(path)
    return m.group("prog"), m.group("base"), m.group("rest")


def tags_in_expression(expr: str) -> List[str]:
    """Identifiers that look like tag paths inside an expression (approximate)."""
    cleaned = _STRIP_LITERALS_RE.sub(" ", expr)
    out: List[str] = []
    for m in _IDENT_RE.finditer(cleaned):
        ident = m.group(0)
        name = _NAME_RE.match(ident.lstrip("\\")).group(0)
        if name.upper() in _EXPR_KEYWORDS:
            continue
        # function call names (e.g. ABS(x)) are not tags
        end = m.end()
        if end < len(cleaned) and cleaned[end:end + 1] == "(":
            continue
        out.append(ident)
        for inner in re.findall(r"\[([^\]]*)\]", ident):  # tags used as array indices
            out.extend(tags_in_expression(inner))
    return out


def classify_operand(raw: str, access: str = READ) -> Operand:
    text = raw.strip()
    if text == "?":
        return Operand(raw=text, kind="unused", access=NONE)
    if text == "":
        return Operand(raw=text, kind="literal", access=NONE)
    if (text[0] == "'" and text[-1] == "'") or (text[0] == '"' and text[-1] == '"'):
        return Operand(raw=text, kind="literal", access=NONE)
    if _NUMBER_RE.match(text) or _RADIX_RE.match(text):
        return Operand(raw=text, kind="literal", access=NONE)
    m = _TAG_RE.match(text)
    if m:
        prog, base, rest = m.group("prog"), m.group("base"), m.group("rest")
        extra: List[str] = []
        for idx in re.findall(r"\[([^\]]*)\]", rest):
            extra.extend(tags_in_expression(idx))
        return Operand(
            raw=text, kind="tag", access=access, tag_base=base,
            tag_path=base + rest, program=prog, extra_tags=extra,
        )
    extra = tags_in_expression(text)
    return Operand(raw=text, kind="expression", access=access, extra_tags=extra)


# --------------------------------------------------------------------------
# Tokenizer
# --------------------------------------------------------------------------
def _split_args(body: str) -> List[str]:
    """Split instruction arguments on top-level commas (respecting quotes, (), [])."""
    args: List[str] = []
    depth = 0
    quote: Optional[str] = None
    cur: List[str] = []
    for ch in body:
        if quote:
            cur.append(ch)
            if ch == quote:
                quote = None
            continue
        if ch in ("'", '"'):
            quote = ch
            cur.append(ch)
        elif ch in "([":
            depth += 1
            cur.append(ch)
        elif ch in ")]":
            depth -= 1
            cur.append(ch)
        elif ch == "," and depth == 0:
            args.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    args.append("".join(cur))
    if quote is not None:
        raise RungSyntaxError(f"unterminated string literal in '{body}'")
    if depth != 0:
        raise RungSyntaxError(f"unbalanced brackets in '{body}'")
    return [a.strip() for a in args]


def tokenize(text: str) -> Tuple[List[Tuple[str, List[str], int]], int]:
    """Return ([(name, raw_args, branch_depth), ...], max_branch_depth).

    Raises RungSyntaxError on malformed text.
    """
    out: List[Tuple[str, List[str], int]] = []
    i, n = 0, len(text)
    depth = 0
    max_depth = 0
    while i < n:
        ch = text[i]
        if ch.isspace():
            i += 1
        elif ch == "[":
            depth += 1
            max_depth = max(max_depth, depth)
            i += 1
        elif ch == "]":
            depth -= 1
            if depth < 0:
                raise RungSyntaxError(f"unexpected ']' at {i}")
            i += 1
        elif ch == ",":
            if depth == 0:
                raise RungSyntaxError(f"unexpected ',' outside a branch at {i}")
            i += 1
        elif ch == ";":
            i += 1
            if text[i:].strip():
                raise RungSyntaxError("text after ';'")
            break
        elif ch.isalpha() or ch == "_":
            m = _NAME_RE.match(text, i)
            name = m.group(0)
            j = m.end()
            while j < n and text[j].isspace():
                j += 1
            if j >= n or text[j] != "(":
                raise RungSyntaxError(f"expected '(' after '{name}' at {i}")
            # find matching ')', respecting quotes and nesting
            k = j + 1
            pdepth = 1
            quote: Optional[str] = None
            while k < n and pdepth > 0:
                c = text[k]
                if quote:
                    if c == quote:
                        quote = None
                elif c in ("'", '"'):
                    quote = c
                elif c == "(":
                    pdepth += 1
                elif c == ")":
                    pdepth -= 1
                k += 1
            if pdepth != 0:
                raise RungSyntaxError(f"unbalanced parentheses in '{name}' at {i}")
            body = text[j + 1:k - 1]
            args = _split_args(body) if body.strip() else []
            out.append((name, args, depth))
            i = k
        else:
            raise RungSyntaxError(f"unexpected character {ch!r} at {i}")
    if depth != 0:
        raise RungSyntaxError("unbalanced branch brackets")
    return out, max_depth


# --------------------------------------------------------------------------
# Classification
# --------------------------------------------------------------------------
def _roles_for(name: str, args: List[str], aoi_usages: Optional[Dict[str, List[str]]]) -> Tuple[List[str], bool, bool]:
    """Return (access per operand, known, is_aoi)."""
    upper = name.upper()
    nargs = len(args)
    if aoi_usages is not None:
        key = name.lower()
        if key in aoi_usages:
            usages = aoi_usages[key]
            roles = [RW]  # backing tag
            for u in usages:
                roles.append({"Input": R, "Output": W, "InOut": RW}.get(u, RW))
            roles += [RW] * max(0, nargs - len(roles))
            return roles[:max(nargs, 1)], True, True
    if upper == "JSR":
        roles = [K, R]
        count = 0
        if nargs >= 2 and args[1].strip().isdigit():
            count = int(args[1])
        roles += [R] * count
        roles += [W] * max(0, nargs - len(roles))
        return roles, True, False
    if upper == "SBR":
        return [W] * nargs, True, False
    if upper == "RET":
        return [R] * nargs, True, False
    if upper in ROLES:
        base = list(ROLES[upper])
        base += [R] * max(0, nargs - len(base))
        return base, True, False
    return [R] * nargs, False, False


def parse_rung(text: str, aoi_usages: Optional[Dict[str, List[str]]] = None) -> Tuple[List[Instruction], int]:
    """Parse neutral rung text into Instructions.

    ``aoi_usages`` maps lower-cased AOI name -> list of Usage strings for the
    parameters that appear in a call (Required or InOut), in definition order.
    """
    tokens, max_depth = tokenize(text)
    instructions: List[Instruction] = []
    for idx, (name, args, depth) in enumerate(tokens):
        roles, known, is_aoi = _roles_for(name, args, aoi_usages)
        operands: List[Operand] = []
        for pos, raw in enumerate(args):
            role = roles[pos] if pos < len(roles) else R
            if role == K:
                operands.append(Operand(raw=raw.strip(), kind="keyword", access=NONE))
            else:
                operands.append(classify_operand(raw, role))
        instructions.append(Instruction(name=name, operands=operands, index=idx, branch_depth=depth, known=known, is_aoi=is_aoi))
    return instructions, max_depth

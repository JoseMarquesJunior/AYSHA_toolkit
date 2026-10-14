"""Decoder for ``<Data Format="L5K">`` values.

Studio 5000 writes only the *parameters* of an AOI instance in the decorated
data; the local tags (timers, nested alarm blocks, HMI words...) exist only in
the raw L5K form: nested bracket lists in data-type layout order. Layout rules
(observed on Studio 5000 v33-v35 exports and self-validated against the
decorated parameters at parse time):

* AOI instance: members = parameters then local tags, in definition order.
  Scalar BOOLs are packed into hidden DINTs: a DINT is inserted at the
  position of the first BOOL of each group of 32; every non-BOOL member
  takes one value (a nested list for structures and arrays).
* UDT: members in definition order including the hidden ``ZZZZZZZZZZ``
  SINTs; BOOL members read ``BitNumber`` of their ``Target`` hidden member.
* TIMER / COUNTER / CONTROL: ``[status, PRE, ACC]`` (``[status, LEN, POS]``).
* STRING and string-family UDTs: ``[LEN, 'text$00...']``.
* Atomic values: ``12``, ``-3.5e+000``, ``2#0000_0001``, ``16#ff``.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .model import Aoi, Controller, DataType

_TOKEN = re.compile(r"\[|\]|,|'(?:[^'$]|\$.)*'|[^\[\],'\s]+")
_ATOMIC = {"BOOL", "SINT", "INT", "DINT", "LINT", "USINT", "UINT", "UDINT", "ULINT", "REAL", "LREAL", "BIT"}
_TIMER_LIKE = {"TIMER": ("PRE", "ACC", (("EN", 31), ("TT", 30), ("DN", 29))), "COUNTER": ("PRE", "ACC", (("CU", 31), ("CD", 30), ("DN", 29), ("OV", 28), ("UN", 27))), "CONTROL": ("LEN", "POS", (("EN", 31), ("EU", 30), ("DN", 29), ("EM", 28), ("ER", 27), ("UL", 26), ("IN", 25), ("FD", 24)))}
_ESC = re.compile(r"\$([0-9A-Fa-f]{2}|.)")


class L5KError(ValueError):
    pass


# --------------------------------------------------------------------------
# Tokenizer / parser
# --------------------------------------------------------------------------
def _atom(tok: str) -> Any:
    if tok[0] == "'":
        return _unescape(tok[1:-1])
    if "#" in tok:
        upper = tok.upper()
        if "QNAN" in upper or "#IND" in upper or "NAN" in upper:
            return float("nan")
        if "#INF" in upper:
            return float("-inf") if tok.startswith("-") else float("inf")
        radix, _, digits = tok.partition("#")
        try:
            base = int(radix)
            return int(digits.replace("_", ""), base)
        except ValueError:
            return tok
    try:
        return int(tok)
    except ValueError:
        pass
    try:
        return float(tok)
    except ValueError:
        return tok


def _unescape(text: str) -> str:
    def repl(m: re.Match) -> str:
        code = m.group(1)
        if len(code) == 2 and all(c in "0123456789abcdefABCDEF" for c in code):
            ch = chr(int(code, 16))
            return "" if ch == "\x00" else ch
        return {"$": "$", "'": "'", "L": "\n", "N": "\n", "R": "\r", "T": "\t", "P": "\f"}.get(code.upper(), code)

    return _ESC.sub(repl, text)


def parse_l5k(text: str) -> Any:
    """'[1,2,[3,4],'ab']' -> [1, 2, [3, 4], 'ab']."""
    tokens = _TOKEN.findall(text)
    pos = 0

    def value() -> Any:
        nonlocal pos
        if pos >= len(tokens):
            raise L5KError("unexpected end")
        tok = tokens[pos]
        if tok == "[":
            pos += 1
            items: List[Any] = []
            while True:
                if pos >= len(tokens):
                    raise L5KError("unterminated list")
                if tokens[pos] == "]":
                    pos += 1
                    return items
                items.append(value())
                if pos < len(tokens) and tokens[pos] == ",":
                    pos += 1
        if tok in ("]", ","):
            raise L5KError(f"unexpected {tok}")
        pos += 1
        return _atom(tok)

    out = value()
    return out


# --------------------------------------------------------------------------
# Layouts
# --------------------------------------------------------------------------
class Layouts:
    def __init__(self, ctrl: Controller) -> None:
        self.udts: Dict[str, DataType] = {d.name.lower(): d for d in ctrl.data_types}
        self.aois: Dict[str, Aoi] = {a.name.lower(): a for a in ctrl.aois if not a.protected}

    # ---- dispatch -------------------------------------------------------------
    def decode(self, value: Any, data_type: str, dims: Optional[str] = None) -> Any:
        if dims:
            if not isinstance(value, list):
                return value
            if "," in dims:
                return value  # multi-dimensional: kept raw
            return {f"[{i}]": self.decode(v, data_type) for i, v in enumerate(value)}
        upper = data_type.upper()
        if upper in _ATOMIC or not isinstance(value, list):
            return value
        if upper in _TIMER_LIKE:
            pre, acc, bits = _TIMER_LIKE[upper]
            out: Dict[str, Any] = {}
            status = value[0] if len(value) > 0 and isinstance(value[0], int) else 0
            for name, bit in bits:
                out[name] = (status >> bit) & 1
            out[pre] = value[1] if len(value) > 1 else None
            out[acc] = value[2] if len(value) > 2 else None
            return out
        if upper == "STRING" or (upper.startswith("STRING") and self.udts.get(data_type.lower()) is None):
            return self._string(value)
        udt = self.udts.get(data_type.lower())
        if udt is not None:
            if (udt.family or "").lower() == "stringfamily":
                return self._string(value)
            return self._udt(value, udt)
        aoi = self.aois.get(data_type.lower())
        if aoi is not None:
            return self.decode_aoi(value, aoi)
        return value

    @staticmethod
    def _string(value: List[Any]) -> Dict[str, Any]:
        length = value[0] if value and isinstance(value[0], int) else 0
        data = value[1] if len(value) > 1 and isinstance(value[1], str) else ""
        return {"LEN": length, "DATA": data[:length] if length >= 0 else data}

    def _udt(self, value: List[Any], udt: DataType) -> Dict[str, Any]:
        out: Dict[str, Any] = {}
        hidden: Dict[str, Any] = {}
        pos = 0
        for m in udt.members:
            if m.data_type.upper() in ("BOOL", "BIT") and not m.dimension and m.target:
                base = hidden.get(m.target.lower())
                out[m.name] = ((base >> (m.bit_number or 0)) & 1) if isinstance(base, int) else None
                continue
            v = value[pos] if pos < len(value) else None
            pos += 1
            if m.hidden:
                hidden[m.name.lower()] = v
                continue
            out[m.name] = self.decode(v, m.data_type, str(m.dimension) if m.dimension else None)
        return out

    def decode_aoi(self, value: List[Any], aoi: Aoi) -> Dict[str, Any]:
        """Parameters + locals; scalar BOOLs packed 32 per hidden DINT."""
        out: Dict[str, Any] = {}
        pos = 0
        packed: Optional[int] = None
        used = 32
        # alias parameters (AliasFor=local member) and InOut parameters (references) hold no data
        members: List[Tuple[str, str, Optional[str]]] = [(p.name, p.data_type, None) for p in aoi.parameters if not p.alias_for and p.usage != "InOut"]
        members += [(t.name, t.data_type, t.dimensions) for t in aoi.local_tags if not t.usage]
        for name, dtype, dims in members:
            if dtype.upper() == "BOOL" and not dims:
                if used >= 32:
                    packed = value[pos] if pos < len(value) and isinstance(value[pos], int) else 0
                    pos += 1
                    used = 0
                out[name] = (packed >> used) & 1 if packed is not None else None
                used += 1
                continue
            v = value[pos] if pos < len(value) else None
            pos += 1
            out[name] = self.decode(v, dtype, dims)
        return out


# --------------------------------------------------------------------------
# Validation against decorated parameters
# --------------------------------------------------------------------------
def _is_nan(x: Any) -> bool:
    if isinstance(x, float):
        return x != x
    return isinstance(x, str) and ("NAN" in x.upper() or "#IND" in x.upper())


def _same(a: Any, b: Any) -> bool:
    if isinstance(a, dict) or isinstance(b, dict):
        if not (isinstance(a, dict) and isinstance(b, dict)):
            return True  # cannot compare shapes: tolerate
        return all(_same(a.get(k), b.get(k)) for k in set(a) & set(b))
    if _is_nan(a) or _is_nan(b):
        return _is_nan(a) and _is_nan(b)
    if isinstance(a, int) and isinstance(b, int) and (a < 0) != (b < 0):
        return (a - b) % (1 << 32) == 0 or (a - b) % (1 << 16) == 0 or (a - b) % (1 << 8) == 0  # same bit pattern, signed vs unsigned
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= 1e-5 * max(1.0, abs(float(a)), abs(float(b)))
    if a is None or b is None:
        return True
    return str(a) == str(b)


def validate(decoded: Dict[str, Any], decorated: Dict[str, Any]) -> Optional[str]:
    """Name of the first parameter whose decoded value differs from the decorated one, else None."""
    for key, val in decorated.items():
        if key in decoded and not _same(decoded[key], val):
            return key
    return None


def merge(decorated: Any, decoded: Any) -> Any:
    """Decorated values win (they are authoritative); decoded fills what is missing (locals)."""
    if isinstance(decorated, dict) and isinstance(decoded, dict):
        out = dict(decoded)
        for k, v in decorated.items():
            out[k] = merge(v, decoded.get(k)) if isinstance(v, dict) else v
        return out
    return decorated if decorated is not None else decoded

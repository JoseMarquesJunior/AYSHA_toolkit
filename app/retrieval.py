"""Seleção de rotinas relevantes para uma pergunta, sem banco vetorial.

Pontuação por: nome da rotina e do programa citados na pergunta, tags citadas
(resolvidas pelo xref), palavras da descrição da rotina e dos comentários de rung.
AOIs citadas pelo nome (ou por uma tag de instância) trazem as rotinas internas
da AOI, identificadas pelo escopo ``aoi:<Nome>``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Set, Tuple, TYPE_CHECKING

from l5x_core.serialize import Serialized
from l5x_core.xref import tag_key

if TYPE_CHECKING:  # pragma: no cover
    from app.llm import ProjectContext

_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*(?:[.:][A-Za-z0-9_]+)*")
_WORD_RE = re.compile(r"[a-zA-ZÀ-ÿ0-9_]{3,}")
_STOPWORDS = {
    "que", "qual", "quais", "como", "para", "por", "com", "sem", "uma", "um", "the", "and", "what", "which",
    "does", "when", "where", "why", "how", "this", "that", "esta", "este", "essa", "esse", "isso", "faz",
    "onde", "quando", "porque", "por que", "rotina", "programa", "rung", "tag", "tags", "lógica", "logica",
    "routine", "program", "explique", "explica", "mostre", "liga", "desliga", "the", "des", "dos", "das",
    "nas", "nos", "ele", "ela", "tem", "são", "sao", "está", "esta",
}


@dataclass
class RoutineHit:
    program: str
    routine: str
    score: float
    reasons: List[str] = field(default_factory=list)

    @property
    def key(self) -> Tuple[str, str]:
        return (self.program, self.routine)


def query_terms(question: str) -> Tuple[Set[str], Set[str]]:
    """Return (identifiers, words) extracted from the question, lower-cased."""
    idents = {m.group(0).lower() for m in _IDENT_RE.finditer(question) if len(m.group(0)) >= 3}
    words = {w.lower() for w in _WORD_RE.findall(question)}
    words = {w for w in words if w not in _STOPWORDS}
    return idents, words


def context_text(history: Optional[Sequence[Dict[str, str]]], turns: int = 2) -> str:
    """Text of the last ``turns`` user messages (and the assistant answers between them),
    so a follow-up like "esses dois AOIs" still resolves to the names cited before."""
    if not history:
        return ""
    users = [m.get("content", "") for m in history if m.get("role") == "user"][-turns:]
    assistants = [m.get("content", "") for m in history if m.get("role") == "assistant"][-1:]
    return "\n".join(users + assistants)


def _score_text(ctx: "ProjectContext", text: str, hits: Dict[Tuple[str, str], RoutineHit], weight: float, tag_reason: str) -> None:
    idents, words = query_terms(text)

    def hit(program: str, routine: str) -> RoutineHit:
        h = hits.get((program, routine))
        if h is None:
            h = hits[(program, routine)] = RoutineHit(program, routine, 0.0)
        return h

    ctrl = ctx.ctrl
    xref = ctx.xref
    aoi_by_name = {a.name.lower(): a for a in ctrl.aois}

    def hit_aoi(aoi, score: float, reason: str) -> None:
        if aoi.protected:
            return
        for r in aoi.routines:
            if not r.has_content:
                continue
            h = hit(f"aoi:{aoi.name}", r.name)
            if reason not in h.reasons:
                h.score += score * weight
                h.reasons.append(reason)

    # tags / AOIs cited in the text -> routines that use them / AOI internals
    for ident in idents:
        base = re.split(r"[.\[:]", ident, maxsplit=1)[0]
        aoi = aoi_by_name.get(base)
        if aoi is not None:
            hit_aoi(aoi, 10, f"AOI {aoi.name}{tag_reason}")
        candidates = [tag_key("controller", base)] + [tag_key(p.name, base) for p in ctrl.programs]
        for key in candidates:
            tag = xref.tags.get(key)
            if tag is None:
                continue
            inst_aoi = aoi_by_name.get(tag.data_type.lower())
            if inst_aoi is not None:
                hit_aoi(inst_aoi, 5, f"instância da AOI {inst_aoi.name}{tag_reason}")
            for u in xref.usages_of(tag):
                if u.scope.startswith("aoi:"):
                    continue
                h = hit(u.scope, u.routine)
                reason = f"tag {tag.name}{tag_reason}"
                if reason not in h.reasons:
                    h.score += 6 * weight
                    h.reasons.append(reason)

    index = ctx.routine_word_index()
    for prog in ctrl.programs:
        pname = prog.name.lower()
        prog_match = pname in idents or pname in words
        for r in prog.routines:
            if r.protected:
                continue
            rname = r.name.lower()
            score = 0.0
            reasons: List[str] = []
            if rname in idents or rname in words:
                score += 10
                reasons.append("nome da rotina" + tag_reason)
            else:
                partial = [w for w in words if len(w) >= 4 and w in rname]
                if partial:
                    score += 4
                    reasons.append("nome parcial da rotina" + tag_reason)
            if prog_match:
                score += 3
                reasons.append("nome do programa" + tag_reason)
            rwords = index.get((prog.name, r.name), set())
            common = words & rwords
            if common:
                desc_words = ctx.routine_description_words(prog.name, r.name)
                n_desc = len(common & desc_words)
                n_comment = min(len(common - desc_words), 5)
                score += 2 * n_desc + 1 * n_comment
                reasons.append("descrição/comentários: " + ", ".join(sorted(common)[:5]) + tag_reason)
            if score > 0:
                h = hit(prog.name, r.name)
                h.score += score * weight
                h.reasons.extend(reasons)


def score_routines(ctx: "ProjectContext", question: str, history: Optional[Sequence[Dict[str, str]]] = None) -> List[RoutineHit]:
    """Score every non-protected routine (and AOI routine) of the project against the
    question; identifiers from the recent conversation count with half weight."""
    hits: Dict[Tuple[str, str], RoutineHit] = {}
    _score_text(ctx, question, hits, 1.0, "")
    previous = context_text(history)
    if previous:
        _score_text(ctx, previous, hits, 0.5, " (histórico)")
    ctrl = ctx.ctrl
    ranked = sorted(hits.values(), key=lambda h: (-h.score, h.program.lower(), h.routine.lower()))
    if not ranked:
        # fallback: main routines of scheduled programs
        for prog in ctrl.programs:
            if prog.task is None or not prog.main_routine:
                continue
            r = prog.routine(prog.main_routine)
            if r is not None and not r.protected:
                ranked.append(RoutineHit(prog.name, r.name, 0.5, ["rotina principal (sem correspondência direta)"]))
    return ranked


def select_routines(
    ctx: "ProjectContext",
    question: str,
    budget_tokens: int,
    max_routines: int = 8,
    history: Optional[Sequence[Dict[str, str]]] = None,
) -> List[Tuple[RoutineHit, Serialized]]:
    """Pick routines in score order while their text fits in ``budget_tokens``."""
    out: List[Tuple[RoutineHit, Serialized]] = []
    used = 0
    for h in score_routines(ctx, question, history):
        if len(out) >= max_routines:
            break
        sr = ctx.routine_text(h.program, h.routine)
        if used + sr.tokens > budget_tokens:
            if not out and sr.tokens > budget_tokens:
                # even the best routine is too big: return it truncated to the budget
                cut = int(budget_tokens * 3.5)
                out.append((h, Serialized(text=sr.text[:cut] + "\n… (rotina truncada ao orçamento de contexto)", tokens=budget_tokens, truncated=True)))
                break
            continue
        out.append((h, sr))
        used += sr.tokens
    return out

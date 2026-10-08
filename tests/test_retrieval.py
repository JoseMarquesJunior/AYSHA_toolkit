import pytest

from app.llm import ProjectContext
from app.retrieval import query_terms, score_routines, select_routines


@pytest.fixture(scope="module")
def pctx(ctrl):
    return ProjectContext(ctrl)


def test_query_terms():
    idents, words = query_terms("O que liga a tag Tank1.Level na rotina MainRoutine?")
    assert "tank1.level" in idents and "mainroutine" in idents
    assert "mainroutine" in words and "liga" not in words and "que" not in words


def test_routine_name_match_ranks_first(pctx):
    hits = score_routines(pctx, "explique a rotina Called")
    assert hits[0].routine == "Called" and "nome da rotina" in hits[0].reasons


def test_tag_match_finds_routines_using_it(pctx):
    hits = score_routines(pctx, "onde a tag Remote é usada?")
    keys = {(h.program, h.routine) for h in hits}
    assert ("MainProgram", "MainRoutine") in keys and ("Unscheduled", "Main") in keys
    assert all("tag Remote" in h.reasons for h in hits if h.key in keys)


def test_tag_with_member_path_resolves_base(pctx):
    hits = score_routines(pctx, "quem escreve Tank1.Level?")
    routines = {h.routine for h in hits}
    assert {"MainRoutine", "Called"} <= routines


def test_comment_and_description_words(pctx):
    hits = score_routines(pctx, "partida do motor")
    assert hits[0].routine == "MainRoutine"
    assert any("comentários" in r or "descrição" in r for r in hits[0].reasons)
    hits = score_routines(pctx, "rotina de falha")
    assert hits[0].routine == "FaultRoutine"


def test_fallback_to_main_routines(pctx):
    hits = score_routines(pctx, "zzz qqq xyz")
    assert hits and all(h.score == 0.5 for h in hits)
    assert {(h.program, h.routine) for h in hits} == {("MainProgram", "MainRoutine"), ("SafetyProg", "Main")}


def test_protected_routines_never_selected(pctx):
    hits = score_routines(pctx, "rotina Protected")
    assert all(h.routine != "Protected" for h in hits)


def test_select_respects_budget(pctx):
    sel = select_routines(pctx, "Motor Start Stop Lamp", budget_tokens=100_000)
    assert sel and sel[0][0].routine == "MainRoutine"
    total = sum(sr.tokens for _, sr in sel)
    assert total <= 100_000
    tiny = select_routines(pctx, "Motor Start Stop Lamp", budget_tokens=50)
    assert len(tiny) == 1 and tiny[0][1].truncated and "truncada" in tiny[0][1].text
    few = select_routines(pctx, "Motor Start Stop Lamp", budget_tokens=100_000, max_routines=1)
    assert len(few) == 1

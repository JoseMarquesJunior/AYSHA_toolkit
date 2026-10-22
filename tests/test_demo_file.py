"""O L5X de demonstração que acompanha o app (app/demo) precisa continuar parseando
e acionando os problemas plantados de propósito, que o Diagnóstico mostra."""

import os

from conftest import ROOT, by_rule
from l5x_core import analyze, build_xref, parse_file
from l5x_core.serialize import routine_text

DEMO = os.path.join(ROOT, "app", "demo", "EB01_Estacao_Bombeamento_DEMO.L5X")


def test_demo_file_parses_and_keeps_planted_findings():
    ctrl = parse_file(DEMO)
    assert ctrl.name == "EB01_DEMO"
    assert [p.name for p in ctrl.programs] == ["Bombeamento", "Utilidades", "Simulacao"]
    hero = next(r for r in ctrl.programs[0].routines if r.name == "Partida_Bombas")
    assert len(hero.rungs) == 10 and hero.description  # curta o bastante para explicação rung a rung
    assert sum(1 for r in hero.rungs if r.comment) >= 6

    findings, metrics = analyze(ctrl, build_xref(ctrl))
    assert metrics.programs == 3 and metrics.routines == 10
    assert not [f for f in findings if f.severity == "error"]
    assert {(f.program, f.routine) for f in by_rule(findings, "routine_never_called")} == {("Bombeamento", "Teste_Comissionamento")}
    assert {f.program for f in by_rule(findings, "program_not_scheduled")} == {"Simulacao"}
    assert {f.tag for f in by_rule(findings, "output_mixed_latch")} == {"B_102_Cmd"}
    assert {f.tag for f in by_rule(findings, "timer_no_preset")} == {"T_Sem_Preset"}
    assert {f.tag for f in by_rule(findings, "tag_unused")} == {"Reserva_01"}
    assert {f.tag for f in by_rule(findings, "output_multiple_writes")} == {"Sirene", "B_101_Fdbk"}
    assert len(by_rule(findings, "fbd_unconnected_tag")) == 1
    assert not by_rule(findings, "unresolved_tags")


def test_demo_routine_text_exposes_timer_presets():
    """O texto que vai ao modelo mostra o preset guardado na tag (o rung traz TON(tag,?,?))."""
    ctrl = parse_file(DEMO)
    text = routine_text(ctrl, "Bombeamento", "Partida_Bombas", build_xref(ctrl)).text
    assert "T_Partida_B102 (TIMER, ctrl, PRE=10000;" in text
    assert "C_Partidas_B101 (COUNTER, ctrl, PRE=500;" in text

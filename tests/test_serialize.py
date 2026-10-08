import json
import subprocess
import sys

import pytest

from l5x_core.serialize import project_overview, routine_text, tag_sheet, to_json, estimate_tokens
from conftest import MINIMAL


def test_project_overview(ctrl, findings, metrics):
    ov = project_overview(ctrl, findings, metrics)
    text = ov.text
    assert ov.tokens == estimate_tokens(text)
    assert ov.tokens <= 6000 and not ov.truncated
    assert "# Projeto MINIMAL" in text
    assert "1756-L85E" in text
    assert "MainTask (CONTINUOUS)" in text and "SafetyTask (PERIODIC 20 ms) [SAFETY]" in text
    assert "Programa Unscheduled [NÃO AGENDADO, DESABILITADO]" in text
    assert "Programa SafetyProg [SAFETY]" in text
    assert "- MainRoutine [RLL, 12 rungs] — Rotina principal" in text
    assert "- Protected [RLL, protegida]" in text
    assert "- StCalled [FBD, 1 folha(s), não analisada em detalhe]" in text
    assert "MyAoi rev. 1.0" in text and "In1:I, Out1:O, IO1:I" in text
    assert "SecretAoi rev. 2.0 [PROTEGIDA]" in text
    assert "- Tank: Level:REAL, FillTimer:TIMER" in text
    assert "IO_Rack1 (1756-IB16, pai: Local)" in text
    assert "jsr_missing_routine: 1" in text
    assert "[error] MainProgram / MainRoutine / Rung 6" in text
    assert "conteúdo de segurança" in text


def test_overview_respects_token_budget(ctrl, findings, metrics):
    ov = project_overview(ctrl, findings, metrics, max_tokens=300)
    assert ov.truncated
    assert ov.tokens <= 300 + 5
    assert "# Projeto MINIMAL" in ov.text


def test_routine_text_rll(ctrl, xref):
    sr = routine_text(ctrl, "MainProgram", "MainRoutine", xref)
    text = sr.text
    assert text.startswith("# MainProgram / MainRoutine [RLL]")
    assert "Chama: Called, Missing" in text
    assert "  // Partida do motor" in text
    assert "Rung 0: XIC(Start)XIO(Stop)OTE(Motor);" in text
    assert "Rung 9:" in text and "(não interpretado:" in text
    assert "## Tags usadas nesta rotina" in text
    assert "- Motor (BOOL, ctrl; L1/E2) — Motor principal" in text
    assert "- Start (BOOL, MainProgram; L" in text  # the program-scoped Start, not the controller one
    assert "- StartAlias (BOOL, ctrl alias de Start;" not in text  # alias has no DataType in L5X
    assert "- StartAlias (, ctrl alias de Start;" in text
    assert "- Idx (DINT, ctrl; L1/E0)" in text
    assert sr.tokens > 100


def test_routine_text_st_fbd_protected_safety(ctrl, xref):
    st = routine_text(ctrl, "MainProgram", "Called", xref).text
    assert "[ST]" in st and "L2:     Tank1.Level := Setpoint * 2.0;" in st
    assert "Chamada por: MainProgram/MainRoutine" in st
    assert "- Tank1 (Tank, ctrl;" in st
    fbd = routine_text(ctrl, "MainProgram", "StCalled", xref).text
    assert "Rotina FBD: 1 folha(s)" in fbd and "- Lamp (BOOL, ctrl;" in fbd
    prot = routine_text(ctrl, "MainProgram", "Protected", xref).text
    assert "protegida" in prot and "Rung" not in prot
    safe = routine_text(ctrl, "SafetyProg", "Main", xref).text
    assert "[SAFETY — somente leitura]" in safe
    with pytest.raises(KeyError):
        routine_text(ctrl, "MainProgram", "Nope", xref)
    with pytest.raises(KeyError):
        routine_text(ctrl, "Nope", "Main", xref)


def test_tag_sheet(ctrl, xref):
    sheet = tag_sheet(ctrl, xref)
    lines = sheet.text.splitlines()
    assert lines[0].startswith("tag | escopo | tipo | leituras | escritas")
    rows = {ln.split(" | ")[0]: ln for ln in lines[2:]}
    assert rows["Motor"] == "Motor | controller | BOOL | 1 | 2 | Motor principal"
    assert rows["Arr"].split(" | ")[2] == "DINT[10]"
    assert rows["Pump"].split(" | ")[1] == "MainProgram"
    prog_only = tag_sheet(ctrl, xref, scope="MainProgram").text
    assert "Motor |" not in prog_only and "Pump |" in prog_only
    small = tag_sheet(ctrl, xref, max_rows=3)
    assert small.truncated and "truncada" in small.text


def test_to_json(ctrl, findings, metrics):
    payload = json.loads(to_json(ctrl, findings, metrics))
    assert payload["controller"]["name"] == "MINIMAL"
    assert payload["metrics"]["programs"] == 3
    assert any(f["rule"] == "jsr_missing_routine" for f in payload["findings"])
    compact = json.loads(to_json(ctrl, findings, metrics, include_rungs=False))
    main = next(r for r in compact["controller"]["programs"][0]["routines"] if r["name"] == "MainRoutine")
    assert main["rungs"] == 12


def test_cli_runs(tmp_path):
    out = tmp_path / "out.json"
    proc = subprocess.run(
        [sys.executable, "-m", "l5x_core.cli", MINIMAL, "--summary", "--findings", "--min-severity", "warning",
         "--dump-routine", "MainProgram/MainRoutine", "--tags", "--json", str(out)],
        capture_output=True, text=True, encoding="utf-8",
    )
    assert proc.returncode == 0, proc.stderr
    assert "# Projeto MINIMAL" in proc.stdout
    assert "jsr_missing_routine" in proc.stdout
    assert "routine_never_called" in proc.stdout
    assert "Rung 0: XIC(Start)XIO(Stop)OTE(Motor);" in proc.stdout
    assert "info     tag_unused" not in proc.stdout  # info rows filtered out of the table
    assert out.exists() and json.loads(out.read_text(encoding="utf-8"))["metrics"]["programs"] == 3
    bad = subprocess.run([sys.executable, "-m", "l5x_core.cli", str(tmp_path / "missing.l5x")], capture_output=True, text=True)
    assert bad.returncode == 2

"""Elementos FBD sem fio (tags soltas na folha): parser, regra fbd_unconnected_tag,
texto da rotina, CLI e helpers do app."""

import os
import subprocess
import sys

from l5x_core import parse_file, build_xref, analyze
from l5x_core.serialize import routine_text
from app.llm import ProjectContext
from app.streamlit_app import UI, unconnected_csv
from conftest import ROOT, by_rule

AOI_L5X = os.path.join(ROOT, "tests", "data", "aoi.l5x")


def test_parser_records_dangling_elements():
    ctrl = parse_file(AOI_L5X)
    fbd = ctrl.program("Main").routine("FbdRoutine")
    assert fbd.sheets == 2
    assert [(d.kind, d.operand, d.sheet) for d in fbd.dangling] == [
        ("IRef", "Loose_In", 2), ("IRef", "XV1.Sts_Fault", 2), ("ORef", "Loose_Out", 2), ("Block", "Lonely_BNOT", 2),
    ]
    assert fbd.dangling[0].is_tag and not fbd.dangling[3].is_tag and fbd.dangling[0].x == 100
    # folha 1 está toda conectada; TextBox e nós ligados não entram
    assert all(d.sheet == 2 for d in fbd.dangling)
    # a rotina FBD do fixture mínimo está toda ligada
    minimal = parse_file(os.path.join(ROOT, "tests", "data", "minimal.l5x"))
    assert minimal.program("MainProgram").routine("StCalled").dangling == []


def test_rule_fbd_unconnected_tag(findings):
    ctrl = parse_file(AOI_L5X)
    rows, _ = analyze(ctrl, build_xref(ctrl))
    f = by_rule(rows, "fbd_unconnected_tag")
    assert len(f) == 1
    f = f[0]
    assert (f.severity, f.program, f.routine, f.tag) == ("warning", "Main", "FbdRoutine", "Loose_In")
    assert "3 referência(s) de tag sem fio" in f.message and "2 entrada(s), 1 saída(s)" in f.message and "1 bloco(s) sem nenhum fio" in f.message
    assert f.evidence == "folha 2: IRef Loose_In, folha 2: IRef XV1.Sts_Fault, folha 2: ORef Loose_Out, folha 2: bloco isolado Block Lonely_BNOT"
    assert by_rule(findings, "fbd_unconnected_tag") == []  # minimal.l5x


def test_routine_text_mentions_unwired(findings):
    ctrl = parse_file(AOI_L5X)
    text = routine_text(ctrl, "Main", "FbdRoutine").text
    assert "SEM FIO (elementos na folha sem nenhuma conexão): IRef Loose_In (folha 2), IRef XV1.Sts_Fault (folha 2), ORef Loose_Out (folha 2), Block Lonely_BNOT (folha 2)" in text


def test_cli_unconnected():
    proc = subprocess.run([sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--unconnected"], capture_output=True, text=True, encoding="utf-8")
    assert proc.returncode == 0, proc.stderr
    assert "4 elemento(s) FBD sem fio em 1 rotina(s)" in proc.stdout
    assert "Main" in proc.stdout and "FbdRoutine" in proc.stdout and "Loose_Out" in proc.stdout and "Lonely_BNOT" in proc.stdout


def test_app_helpers():
    ctx = ProjectContext(parse_file(AOI_L5X))
    rows = ctx.unconnected_rows()
    assert [r["elemento"] for r in rows] == ["Loose_In", "XV1.Sts_Fault", "Loose_Out", "Lonely_BNOT"]
    assert sum(1 for r in rows if r["tag"]) == 3
    csv = unconnected_csv(rows)
    assert csv.splitlines()[0] == "programa;rotina;folha;tipo;elemento" and "Main;FbdRoutine;2;IRef;Loose_In" in csv
    for lang in UI:
        assert "unconnected_section" in UI[lang]

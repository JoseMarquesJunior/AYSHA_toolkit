"""Exportação da configuração das instâncias de AOI (valores guardados na tag de instância)."""

import os
import subprocess
import sys

import pytest
from openpyxl import load_workbook

from l5x_core import parse_file, build_xref
from l5x_core.aoi_config import FIXED_COLUMNS, build_sheets, flatten, members_for, to_csv, write_xlsx, xlsx_bytes
from app.llm import ProjectContext
from app.streamlit_app import UI
from conftest import ROOT

AOI_L5X = os.path.join(ROOT, "tests", "data", "aoi.l5x")


@pytest.fixture(scope="module")
def actrl():
    return parse_file(AOI_L5X)


@pytest.fixture(scope="module")
def axref(actrl):
    return build_xref(actrl)


def test_flatten():
    assert flatten({"A": 1, "T": {"PRE": 5, "ACC": 0}, "S": {"LEN": 3, "DATA": "abc"}, "Arr": {"[0]": {"X": 1}, "[1]": {"X": 2}}}) == {
        "A": 1, "T.PRE": 5, "T.ACC": 0, "S": "abc", "Arr[0].X": 1, "Arr[1].X": 2,
    }
    assert flatten(7, "v") == {"v": 7} and flatten(None) == {}


def test_parser_keeps_instance_values_and_extended_properties(actrl):
    xv1 = next(t for t in actrl.tags if t.name == "XV1")
    assert xv1.values["Cfg_TravelT"] == 7000 and xv1.values["Wrk_T"]["PRE"] == 7000 and xv1.values["Wrk_Scl"]["In"] == 0.0
    assert xv1.engineering_unit == "%" and xv1.area == "HCS" and xv1.label is None
    arr = next(t for t in actrl.tags if t.name == "XV_Arr")
    assert arr.values["[1]"]["Cfg_TravelT"] == 200
    assert next(t for t in actrl.tags if t.name == "XV3").values is None  # sem Data decorado
    assert next(t for t in actrl.tags if t.name == "Fbk1").values is None  # BOOL comum não guarda valores


def test_members_config_vs_all(actrl):
    valve = actrl.aoi("Valve")
    cfg = [(m.name, m.kind) for m in members_for(valve, "config")]
    # entradas não obrigatórias + locais editáveis pelo HMI
    assert cfg == [("Cmd_Open", "parâmetro"), ("Cfg_TravelT", "parâmetro"), ("Cfg_HMI", "local")]
    everything = [m.name for m in members_for(valve, "all")]
    assert everything[:3] == ["Cmd_Open", "Fbk_Open", "Cfg_TravelT"] and "EnableIn" not in everything
    assert {"Wrk_T", "Wrk_Scl", "Sts_HMI"} <= set(everything)
    with pytest.raises(ValueError):
        build_sheets(actrl, mode="nope")


def test_build_sheets_config(actrl, axref):
    sheets = {s.aoi: s for s in build_sheets(actrl, axref, mode="config")}
    assert set(sheets) == {"Valve", "Scaler", "Locked"}  # Orphan não tem instâncias
    v = sheets["Valve"]
    assert v.columns == ["Cmd_Open", "Cfg_TravelT", "Cfg_HMI"] and v.header == FIXED_COLUMNS + v.columns
    rows = {r["Instancia"]: r for r in v.rows}
    assert list(rows) == ["XV1", "XV2", "XV3", "XV_Arr[0]", "XV_Arr[1]"]  # arrays expandidos; aninhada (Wrk_Scl) não entra
    assert rows["XV1"]["Cfg_TravelT"] == 7000 and rows["XV1"]["Cmd_Open"] == 1 and rows["XV1"]["Cfg_HMI"] == 2.5
    assert rows["XV1"]["PLC"] == "AOITEST" and rows["XV1"]["Escopo"] == "controlador" and rows["XV1"]["Descricao"] == "Válvula 1"
    assert rows["XV1"]["Unidade"] == "%" and rows["XV1"]["Area"] == "HCS"
    assert rows["XV3"]["Cfg_TravelT"] == "" and rows["XV_Arr[1]"]["Cfg_TravelT"] == 200
    assert [m.name for m in v.dictionary] == v.columns
    # Scaler: In é obrigatório -> nada de configuração, mas a aba existe (sem colunas de membro)
    assert sheets["Scaler"].columns == [] and [r["Instancia"] for r in sheets["Scaler"].rows] == ["Scl1", "Scl2"]
    assert sheets["Locked"].protected and sheets["Locked"].columns == [] and sheets["Locked"].dictionary == []
    only = build_sheets(actrl, axref, mode="config", aoi="valve")
    assert len(only) == 1 and only[0].aoi == "Valve"


def test_build_sheets_all_flattens_structures(actrl, axref):
    v = next(s for s in build_sheets(actrl, axref, mode="all") if s.aoi == "Valve")
    assert v.columns[:4] == ["Cmd_Open", "Fbk_Open", "Cfg_TravelT", "Out_Open"]
    assert {"Wrk_T.PRE", "Wrk_T.ACC", "Wrk_T.DN", "Wrk_Scl.In", "Wrk_Scl.Out", "Sts_HMI"} <= set(v.columns)
    xv1 = next(r for r in v.rows if r["Instancia"] == "XV1")
    assert xv1["Wrk_T.ACC"] == 123 and xv1["Wrk_Scl.In"] == 0.0
    # em modo config os timers ficam só com PRE
    cfg = next(s for s in build_sheets(actrl, axref, mode="config") if s.aoi == "Valve")
    assert not any(c.startswith("Wrk_T") for c in cfg.columns)


def test_csv_and_xlsx(actrl, axref, tmp_path):
    sheets = build_sheets(actrl, axref, mode="config")
    valve = next(s for s in sheets if s.aoi == "Valve")
    csv = to_csv(valve)
    lines = csv.splitlines()
    assert lines[0] == ";".join(FIXED_COLUMNS + ["Cmd_Open", "Cfg_TravelT", "Cfg_HMI"])
    assert lines[1] == "AOITEST;Valve;XV1;controlador;Válvula 1;%;HCS;1;7000;2.5"
    path = tmp_path / "cfg.xlsx"
    write_xlsx(sheets, str(path), "config")
    wb = load_workbook(path)
    assert wb.sheetnames == ["Resumo", "Valve", "Scaler", "Locked", "Dicionario"]
    ws = wb["Valve"]
    assert [c.value for c in ws[1]] == valve.header and ws["C2"].value == "XV1" and ws["I2"].value == 7000
    assert ws.freeze_panes == "D2"
    summary = {r[0]: r for r in wb["Resumo"].iter_rows(min_row=2, values_only=True)}
    assert summary["Valve"][1] == 5 and summary["Locked"][3] == "sim"
    dic = [r for r in wb["Dicionario"].iter_rows(min_row=2, values_only=True)]
    assert ("Valve", "Cfg_TravelT", "parâmetro", "Input", "DINT", None, None, "Read/Write", 5000, "Tempo de curso (ms)") in dic  # openpyxl lê "" como None
    assert xlsx_bytes(sheets)[:2] == b"PK"


def test_cli_export(tmp_path):
    out = tmp_path / "cfg.xlsx"
    proc = subprocess.run([sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--export-config", str(out)], capture_output=True, text=True, encoding="utf-8")
    assert proc.returncode == 0, proc.stderr
    assert out.exists() and "Valve (5 instâncias, 3 colunas)" in proc.stderr
    csv_out = tmp_path / "valve.csv"
    proc = subprocess.run([sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--aoi", "Valve", "--export-config", str(csv_out), "--config-mode", "all"], capture_output=True, text=True, encoding="utf-8")
    assert proc.returncode == 0, proc.stderr
    assert csv_out.read_text(encoding="utf-8-sig").splitlines()[0].startswith("PLC;AOI;Instancia;Escopo;Descricao;Unidade;Area;Cmd_Open;Fbk_Open")
    bad = subprocess.run([sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--export-config", str(tmp_path / "x.csv")], capture_output=True, text=True, encoding="utf-8")
    assert bad.returncode == 2 and "única AOI" in bad.stderr


def test_project_context_helpers_and_ui(actrl):
    ctx = ProjectContext(actrl)
    sheet = ctx.config_sheet("Valve")
    assert sheet is not None and len(sheet.rows) == 5
    assert ctx.config_sheet("Orphan") is None
    data = ctx.config_xlsx("config")
    assert data[:2] == b"PK" and ctx.config_xlsx("config") is data  # cache
    for lang in UI:
        assert "export_xlsx" in UI[lang] and "export_mode_config" in UI[lang]

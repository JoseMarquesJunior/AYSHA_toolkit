import os
import time

import pytest

from l5x_core import parse_file, parse_bytes, L5XError, build_xref, analyze
from conftest import SAMPLE


def test_header_and_controller(ctrl):
    assert ctrl.name == "MINIMAL"
    assert ctrl.target_name == "MINIMAL"
    assert ctrl.software_revision == "35.00"
    assert ctrl.export_date.startswith("Wed Oct 08")
    assert ctrl.processor_type == "1756-L85E"
    assert ctrl.major_rev == "35"
    assert "sintético" in ctrl.description


def test_tasks_and_scheduling(ctrl):
    assert [t.name for t in ctrl.tasks] == ["MainTask", "SafetyTask"]
    main = ctrl.tasks[0]
    assert main.type == "CONTINUOUS" and main.programs == ["MainProgram"]
    assert ctrl.tasks[1].is_safety and ctrl.tasks[1].rate == "20"
    assert ctrl.program("MainProgram").task == "MainTask"
    assert ctrl.program("Unscheduled").task is None
    assert ctrl.program("Unscheduled").disabled is True
    assert ctrl.program("SafetyProg").is_safety


def test_programs_and_routines(ctrl):
    prog = ctrl.program("MainProgram")
    assert prog.main_routine == "MainRoutine"
    assert prog.fault_routine == "FaultRoutine"
    assert prog.description == "Programa principal"
    assert [r.name for r in prog.routines] == ["MainRoutine", "Called", "StCalled", "Orphan", "Protected", "FaultRoutine"]
    types = {r.name: r.type for r in prog.routines}
    assert types["Called"] == "ST" and types["StCalled"] == "FBD"
    main = prog.routine("MainRoutine")
    assert len(main.rungs) == 12
    assert main.rungs[0].comment == "Partida do motor"
    assert main.rungs[1].comment is None
    assert main.rungs[0].text.startswith("XIC(Start)")
    assert main.description == "Rotina principal"
    assert main.calls == ["Called", "Missing"]


def test_protected_routine_and_aoi(ctrl):
    prot = ctrl.program("MainProgram").routine("Protected")
    assert prot.protected is True
    assert prot.rungs == [] and not prot.analyzed
    secret = ctrl.aoi("SecretAoi")
    assert secret.protected is True
    assert secret.revision == "2.0"
    assert secret.routines == [] and secret.parameters == []


def test_rung_parse_error_becomes_finding(ctrl):
    main = ctrl.program("MainProgram").routine("MainRoutine")
    broken = main.rungs[9]
    assert broken.parse_error
    assert broken.instructions == []
    errs = [f for f in ctrl.parse_findings if f.rule == "rung_parse_error"]
    assert len(errs) == 1
    assert errs[0].severity == "info"
    assert errs[0].program == "MainProgram" and errs[0].routine == "MainRoutine" and errs[0].rung == 9
    assert "XIC(Start OTE(Lamp)" in errs[0].evidence
    # the rest of the routine is still parsed
    assert main.rungs[10].instructions[0].name == "XIC"


def test_tags_by_scope(ctrl):
    names = {t.name for t in ctrl.tags}
    assert {"Start", "Stop", "Motor", "T1", "Tank1", "StartAlias", "Produced1", "Consumed1"} <= names
    start = next(t for t in ctrl.tags if t.name == "Start")
    assert start.description == "Botão de partida" and start.data_type == "BOOL"
    alias = next(t for t in ctrl.tags if t.name == "StartAlias")
    assert alias.tag_type == "Alias" and alias.alias_for == "Start"
    safe = next(t for t in ctrl.tags if t.name == "SafeTag")
    assert safe.is_safety
    prod = next(t for t in ctrl.tags if t.name == "Produced1")
    assert prod.tag_type == "Produced"
    prog_tags = {t.name: t for t in ctrl.program("MainProgram").tags}
    assert set(prog_tags) == {"LocalFlag", "Pump", "LocalUnused", "Start"}
    assert prog_tags["Pump"].scope == "MainProgram"
    arr = next(t for t in ctrl.tags if t.name == "Arr")
    assert arr.dimensions == "10"
    const = next(t for t in ctrl.tags if t.name == "ConstTag")
    assert const.constant


def test_timer_values_only_for_bearing_types(ctrl):
    tags = {t.name: t for t in ctrl.tags}
    assert tags["T1"].values == {"PRE": 0, "ACC": 0, "EN": 0, "TT": 0, "DN": 0}
    assert tags["T2"].values["PRE"] == 5000
    assert tags["Tank1"].values["FillTimer"]["PRE"] == 0
    assert tags["Counter1"].values["PRE"] == 10
    assert tags["Start"].values is None  # plain BOOL: values not kept


def test_udts(ctrl):
    assert len(ctrl.data_types) == 1
    tank = ctrl.data_types[0]
    assert tank.name == "Tank" and tank.description == "Tanque genérico"
    assert [(m.name, m.data_type) for m in tank.members] == [("Level", "REAL"), ("FillTimer", "TIMER")]


def test_aois(ctrl):
    aoi = ctrl.aoi("MyAoi")
    assert aoi.revision == "1.0" and aoi.vendor == "Test"
    assert [p.name for p in aoi.call_parameters] == ["In1", "Out1", "IO1"]
    assert [p.usage for p in aoi.call_parameters] == ["Input", "Output", "InOut"]
    assert len(aoi.routines) == 1 and aoi.routines[0].name == "Logic"
    assert aoi.routines[0].rungs[0].instructions[0].name == "GRT"
    local = {t.name for t in aoi.local_tags}
    assert {"Wrk_T", "In1", "Out1", "IO1", "Opt"} <= local


def test_modules(ctrl):
    assert [(m.name, m.catalog, m.parent) for m in ctrl.modules] == [
        ("Local", "1756-L85E", "Local"),
        ("IO_Rack1", "1756-IB16", "Local"),
    ]
    assert ctrl.modules[1].description == "Cartão de entradas digitais"


def test_st_routine(ctrl):
    st = ctrl.program("MainProgram").routine("Called")
    assert len(st.st_lines) == 7
    assert st.calls == ["StCalled"]
    refs = {(r.tag_path, r.access) for r in st.tag_refs}
    assert ("Tank1.Level", "write") in refs
    assert ("Setpoint", "read") in refs
    assert ("Start", "read") in refs
    assert ("StrTag", "write") in refs
    assert ("MyAoi1", "readwrite") in refs
    # comment and string contents are not references
    paths = {r.tag_path for r in st.tag_refs}
    assert "comentário" not in paths and "Rotina" not in paths and "JSR" not in paths


def test_fbd_routine_registered_not_analyzed(ctrl):
    fbd = ctrl.program("MainProgram").routine("StCalled")
    assert fbd.type == "FBD" and fbd.sheets == 1 and fbd.size > 0
    assert not fbd.analyzed
    refs = {(r.tag_path, r.access) for r in fbd.tag_refs}
    assert refs == {("Setpoint", "read"), ("Lamp", "write"), ("Scl1", "readwrite")}


def test_aoi_call_operands_classified(ctrl):
    rung = ctrl.program("MainProgram").routine("MainRoutine").rungs[5]
    aoi = rung.instructions[0]
    assert aoi.name == "MyAoi" and aoi.is_aoi
    assert [o.access for o in aoi.operands] == ["readwrite", "read", "write", "readwrite"]


def test_parse_bytes_and_invalid_input():
    with open(os.path.join(os.path.dirname(__file__), "data", "minimal.l5x"), "rb") as fh:
        ctrl = parse_bytes(fh.read())
    assert ctrl.name == "MINIMAL"
    with pytest.raises(L5XError):
        parse_bytes(b"<html><body>not an l5x</body></html>")
    with pytest.raises(L5XError):
        parse_bytes(b"<RSLogix5000Content></RSLogix5000Content>")
    with pytest.raises(L5XError):
        parse_bytes(b"\x00\x01 not xml")


@pytest.mark.skipif(not os.path.exists(SAMPLE), reason="samples/P80_HULL_HCS01.L5X não disponível")
def test_real_sample_parses_fast():
    t0 = time.perf_counter()
    ctrl = parse_file(SAMPLE)
    xref = build_xref(ctrl)
    findings, metrics = analyze(ctrl, xref)
    elapsed = time.perf_counter() - t0
    print(f"\n{SAMPLE}: {elapsed:.2f}s; metrics={metrics}")
    assert elapsed < 30
    assert metrics.programs > 0 and metrics.routines > 0 and metrics.rungs > 0
    assert metrics.tags_controller > 0
    rules = {f.rule for f in findings}
    # the acceptance criterion: orphan routines, unused tags and duplicated outputs with location
    assert "routine_never_called" in rules
    assert "tag_unused" in rules
    for f in findings:
        if f.rule == "routine_never_called":
            assert f.program and f.routine
        if f.rule in ("output_multiple_writes", "output_mixed_latch"):
            assert f.program and f.routine and f.rung is not None
    assert not metrics.unknown_instructions, metrics.unknown_instructions

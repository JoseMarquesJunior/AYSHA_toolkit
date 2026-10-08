from conftest import by_rule


def one(findings, rule, **match):
    rows = [f for f in by_rule(findings, rule) if all(getattr(f, k) == v for k, v in match.items())]
    assert len(rows) == 1, f"{rule} {match}: {rows}"
    return rows[0]


def test_rule1_routine_never_called(findings):
    rows = by_rule(findings, "routine_never_called")
    assert {(f.program, f.routine) for f in rows} == {("MainProgram", "Orphan"), ("MainProgram", "Protected")}
    assert all(f.severity == "warning" for f in rows)


def test_rule2_program_not_scheduled_or_disabled(findings):
    f = one(findings, "program_not_scheduled")
    assert f.program == "Unscheduled" and f.severity == "warning"
    f = one(findings, "program_disabled")
    assert f.program == "Unscheduled"


def test_rule3_tag_unused(findings):
    rows = by_rule(findings, "tag_unused")
    assert {(f.program, f.tag) for f in rows} == {(None, "UnusedTag"), ("MainProgram", "LocalUnused")}
    # produced/consumed are ignored
    assert not any(f.tag in ("Produced1", "Consumed1") for f in rows)


def test_rule4_write_only_and_read_only(findings):
    write_only = {f.tag for f in by_rule(findings, "tag_write_only")}
    assert {"WriteOnly", "Lamp", "AoiOut", "SafeOut"} <= write_only
    assert "Motor" not in write_only and "Remote" not in write_only
    read_only = {f.tag for f in by_rule(findings, "tag_read_only")}
    assert {"Setpoint", "Stop", "Idx", "SafeTag"} <= read_only
    assert "ConstTag" not in read_only  # constants are expected to be read-only
    assert "StartAlias" not in read_only  # aliases excluded
    safe = one(findings, "tag_write_only", tag="SafeOut")
    assert safe.safety is True
    f = one(findings, "tag_write_only", tag="WriteOnly")
    assert "MainProgram / MainRoutine / rung 8" in f.evidence


def test_rule5_output_written_in_multiple_places(findings):
    f = one(findings, "output_multiple_writes")
    assert f.tag == "Motor" and f.program == "MainProgram" and f.routine == "MainRoutine" and f.rung == 0
    assert "rung 0" in f.evidence and "rung 1" in f.evidence
    f = one(findings, "output_mixed_latch")
    assert f.tag == "Pump" and f.rung == 7
    assert "OTL" in f.evidence and "OTU" in f.evidence and "OTE" in f.evidence


def test_rule6_timer_without_preset(findings):
    rows = by_rule(findings, "timer_no_preset")
    tags = {f.tag: f for f in rows}
    assert set(tags) == {"T1", "Tank1.FillTimer", "T2"}
    assert all(f.program == "MainProgram" and f.routine == "MainRoutine" and f.rung == 2 for f in rows)
    assert "literal 0" in tags["T2"].message
    assert "PRE=0" in tags["T1"].message
    # counter with PRE=10 must not be flagged
    assert "Counter1" not in tags


def test_rule7_jsr_missing(findings):
    f = one(findings, "jsr_missing_routine")
    assert f.severity == "error"
    assert (f.program, f.routine, f.rung) == ("MainProgram", "MainRoutine", 6)
    assert "Missing" in f.message


def test_rule8_documentation(findings):
    rows = by_rule(findings, "rungs_without_comment")
    per_routine = {(f.program, f.routine): f for f in rows}
    assert set(per_routine) == {("MainProgram", "MainRoutine"), ("MainProgram", "Orphan"), ("Unscheduled", "Main")}
    main = per_routine[("MainProgram", "MainRoutine")]
    assert main.rung == 1 and "1 de 12" in main.message and "92% comentados" in main.message
    no_desc = {(f.program, f.routine) for f in by_rule(findings, "routine_no_description")}
    assert no_desc == {("MainProgram", "Called"), ("MainProgram", "StCalled"), ("MainProgram", "Orphan"), ("Unscheduled", "Main"), ("SafetyProg", "Main")}
    assert all(f.severity == "info" for f in rows)


def test_rule9_protected(findings):
    f = one(findings, "routine_protected")
    assert (f.program, f.routine, f.severity) == ("MainProgram", "Protected", "info")
    f = one(findings, "aoi_protected")
    assert f.tag == "SecretAoi"


def test_rule10_safety(findings):
    rows = by_rule(findings, "safety_content")
    assert all(f.safety for f in rows)
    msgs = " ".join(f.message for f in rows)
    assert "SafetyTask" in msgs and "SafetyProg" in msgs and "SafeTag" in msgs and "SafeOut" in msgs
    # every finding inside the safety program carries the flag
    for f in findings:
        if f.program == "SafetyProg":
            assert f.safety, f


def test_extra_rules(findings):
    f = one(findings, "unknown_instruction")
    assert f.tag == "FOO" and "1 ocorrência" in f.message
    f = one(findings, "unresolved_tags")
    assert f.routine == "MainRoutine" and "Lamp2" in f.evidence
    f = one(findings, "rung_parse_error")
    assert f.rung == 9


def test_rule11_metrics(metrics):
    assert metrics.programs == 3
    assert metrics.routines == 8
    assert metrics.routines_by_type == {"RLL": 6, "ST": 1, "FBD": 1}
    assert metrics.rungs == 16
    assert metrics.st_lines == 7
    assert metrics.instructions > 20
    assert metrics.tags_controller == 24 and metrics.tags_program == 6
    assert metrics.aois == 2 and metrics.aois_protected == 1
    assert metrics.udts == 1 and metrics.modules == 2 and metrics.tasks == 2
    assert metrics.largest_routine == "MainProgram/MainRoutine" and metrics.largest_routine_rungs == 12
    assert metrics.max_branch_depth == 2
    assert metrics.protected_routines == 1 and metrics.safety_programs == 1
    assert metrics.unknown_instructions == {"FOO": 1}
    assert metrics.findings_by_severity["error"] == 1
    assert 0 < metrics.commented_rung_pct < 100


def test_findings_sorted_by_severity(findings):
    order = {"error": 0, "warning": 1, "info": 2}
    sev = [order[f.severity] for f in findings]
    assert sev == sorted(sev)

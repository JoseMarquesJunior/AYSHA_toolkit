from l5x_core.xref import tag_key


def tag(xref, scope, name):
    return xref.tags[tag_key(scope, name)]


def test_program_scope_shadows_controller(ctrl, xref):
    # MainProgram declares its own 'Start'; rung 0 must resolve to it
    prog_start = tag(xref, "MainProgram", "Start")
    ctrl_start = tag(xref, "controller", "Start")
    prog_uses = xref.usages_of(prog_start)
    assert any(u.routine == "MainRoutine" and u.rung == 0 and u.instruction == "XIC" for u in prog_uses)
    assert all(u.scope == "MainProgram" for u in prog_uses)
    # controller Start is reached through the alias (rung 1) and from the other program
    ctrl_uses = xref.usages_of(ctrl_start)
    assert any(u.scope == "Unscheduled" for u in ctrl_uses)
    assert any(u.scope == "MainProgram" and u.rung == 1 for u in ctrl_uses)


def test_alias_resolves_to_target_but_keeps_own_usage(ctrl, xref):
    alias = tag(xref, "controller", "StartAlias")
    uses = xref.usages_of(alias)
    assert len(uses) == 1 and uses[0].rung == 1 and uses[0].path == "StartAlias"
    target = tag(xref, "controller", "Start")
    assert any(u.rung == 1 and u.path == "StartAlias" for u in xref.usages_of(target))


def test_cross_program_reference(ctrl, xref):
    remote = tag(xref, "Unscheduled", "Remote")
    uses = xref.usages_of(remote)
    assert any(u.scope == "MainProgram" and u.rung == 10 and u.reads for u in uses)
    assert any(u.scope == "Unscheduled" and u.writes for u in uses)


def test_read_write_classification(ctrl, xref):
    motor = tag(xref, "controller", "Motor")
    assert sorted((u.rung, u.access) for u in xref.usages_of(motor)) == [(0, "write"), (1, "write"), (2, "read")]
    t1 = tag(xref, "controller", "T1")
    assert [u.access for u in xref.usages_of(t1)] == ["readwrite"]
    write_only = tag(xref, "controller", "WriteOnly")
    assert xref.writes_of(write_only) and not xref.reads_of(write_only)


def test_member_and_index_paths(ctrl, xref):
    tank = tag(xref, "controller", "Tank1")
    paths = {u.path for u in xref.usages_of(tank)}
    assert {"Tank1.FillTimer", "Tank1.Level"} <= paths
    idx = tag(xref, "controller", "Idx")
    uses = xref.usages_of(idx)
    assert uses and uses[0].approximate and uses[0].access == "read"
    arr = tag(xref, "controller", "Arr")
    assert {u.path for u in xref.usages_of(arr)} == {"Arr[Idx]", "Arr[0]"}


def test_st_and_fbd_usages_are_approximate(ctrl, xref):
    sp = tag(xref, "controller", "Setpoint")
    st = [u for u in xref.usages_of(sp) if u.routine == "Called"]
    assert st and all(u.approximate for u in st)
    fbd = [u for u in xref.usages_of(sp) if u.routine == "StCalled"]
    assert fbd and fbd[0].location == "sheet 1" and fbd[0].instruction == "FBD"
    lamp = tag(xref, "controller", "Lamp")
    assert any(u.routine == "StCalled" and u.writes for u in xref.usages_of(lamp))


def test_unresolved_and_keywords(ctrl, xref):
    unresolved = xref.unresolved[("MainProgram", "MainRoutine")]
    assert unresolved == ["Lamp2"]
    # keyword operands (PROGRAM, THIS, routine names) never show up as unresolved
    for names in xref.unresolved.values():
        assert "THIS" not in names and "PROGRAM" not in names and "Called" not in names


def test_call_graph(ctrl, xref):
    assert xref.calls[("MainProgram", "MainRoutine")] == ["Called", "Missing"]
    assert xref.calls[("MainProgram", "Called")] == ["StCalled"]
    assert xref.callers_of("MainProgram", "Called") == [("MainProgram", "MainRoutine")]
    assert xref.callers_of("MainProgram", "StCalled") == [("MainProgram", "Called")]
    assert xref.callers_of("MainProgram", "Orphan") == []
    assert xref.missing_calls == [("MainProgram", "MainRoutine", "Missing")]


def test_aoi_scope_isolated(ctrl, xref):
    in1 = tag(xref, "aoi:MyAoi", "In1")
    uses = xref.usages_of(in1)
    assert uses and uses[0].scope == "aoi:MyAoi" and uses[0].routine == "Logic"
    wrk = tag(xref, "aoi:MyAoi", "Wrk_T")
    assert xref.usages_of(wrk)[0].access == "readwrite"

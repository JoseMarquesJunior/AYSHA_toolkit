"""Função de análise, documentação e debug de AOIs: parser (chamadas RLL/ST/FBD),
xref (instâncias e pontos de chamada), regras aoi_*, relatório, Markdown,
serialização para o LLM, CLI, modo LLM "aoi" (cliente falso), retrieval e app."""

import os
import subprocess
import sys

import pytest

from l5x_core import parse_file, build_xref, analyze
from l5x_core.aoi import aoi_markdown, aoi_summary_rows, build_aoi_report, instance_counts
from l5x_core.serialize import aoi_interface_text, aoi_text, project_overview, routine_text
from app.llm import LLMClient, LLMConfig, LLMError, ProjectContext, MODES, DEEP_MODES, load_prompt
from app.retrieval import score_routines, select_routines
from app.streamlit_app import action_cost, UI
from conftest import ROOT, by_rule
from fake_anthropic import FakeAnthropic

AOI_L5X = os.path.join(ROOT, "tests", "data", "aoi.l5x")


@pytest.fixture(scope="module")
def actrl():
    return parse_file(AOI_L5X)


@pytest.fixture(scope="module")
def axref(actrl):
    return build_xref(actrl)


@pytest.fixture(scope="module")
def afindings(actrl, axref):
    return analyze(actrl, axref)[0]


@pytest.fixture(scope="module")
def ametrics(actrl, axref):
    return analyze(actrl, axref)[1]


@pytest.fixture(scope="module")
def apctx(actrl):
    return ProjectContext(actrl)


# --------------------------------------------------------------------------
# parser
# --------------------------------------------------------------------------
def test_aoi_definition_metadata_and_defaults(actrl):
    v = actrl.aoi("Valve")
    assert v.execute_prescan and v.execute_enableinfalse and not v.execute_postscan
    assert v.revision_note == "2.1 - adicionada falha de curso"
    assert v.help_text.startswith("Comanda a válvula")
    assert v.created_date.startswith("2024-01-02") and v.software_revision == "v35.00"
    assert [p.name for p in v.call_parameters] == ["Fbk_Open", "Out_Open", "Ref_Cfg", "Spare_In", "Inp_Written"]
    assert [p.name for p in v.user_parameters][:2] == ["Cmd_Open", "Fbk_Open"] and all(not p.is_system for p in v.user_parameters)
    assert v.parameter("Cfg_TravelT").default == 5000 and v.parameter("Cmd_Open").default == 0
    assert v.parameter("Fbk_Open").default is None
    assert v.parameter("Out_Open").external_access == "Read Only"
    wrk_t = next(t for t in v.local_tags if t.name == "Wrk_T")
    assert wrk_t.values["PRE"] == 5000 and wrk_t.external_access == "None" and not wrk_t.hmi_readable
    sts = next(t for t in v.local_tags if t.name == "Sts_HMI")
    assert sts.hmi_readable and not sts.hmi_writable
    cfg = next(t for t in v.local_tags if t.name == "Cfg_HMI")
    assert cfg.hmi_writable
    assert v.routine("prescan").name == "Prescan" and v.analyzable
    assert actrl.aoi("Locked").protected and not actrl.aoi("Locked").analyzable


def test_aoi_calls_rll_st_fbd(actrl):
    main = actrl.program("Main")
    rll = main.routine("MainRoutine").aoi_calls
    assert [(c.aoi, c.instance, c.location, c.source) for c in rll] == [
        ("Valve", "XV1", "rung 0", "rll"), ("Valve", "XV_Arr[0]", "rung 1", "rll"), ("Locked", "Lk1", "rung 2", "rll"),
    ]
    assert rll[0].arguments == ["Fbk1", "Out1", "Cfg1", "Spare1", "RawA"] and rll[0].rung == 0
    assert rll[1].instance_base == "XV_Arr" and rll[1].arguments == ["Fbk1", "Out1"]
    st = main.routine("StRoutine").aoi_calls
    assert [(c.aoi, c.instance, c.location, c.arguments) for c in st] == [
        ("Scaler", "Scl1", "line 0", ["RawA", "SclOut"]),
        ("Valve", "XV_Arr[1]", "line 1", ["Fbk1", "Out1", "Cfg1", "Spare1", "RawA"]),
    ]
    fbd = main.routine("FbdRoutine").aoi_calls
    assert [(c.aoi, c.instance, c.location, c.source) for c in fbd] == [("Scaler", "Scl2", "sheet 1", "fbd"), ("Valve", "XV1", "sheet 1", "fbd")]
    assert fbd[0].pins == {"In": "RawA", "Out": "SclOut"}
    assert fbd[1].pins == {"Ref_Cfg": "Cfg1", "Cmd_Open": "Cmd2", "Out_Open": "Out1", "Sts_Fault": "Fault1"}
    # nested calls inside AOI routines
    nested = actrl.aoi("Valve").routine("Logic").aoi_calls
    assert [(c.aoi, c.instance, c.rung) for c in nested] == [("Scaler", "Wrk_Scl", 6)]
    assert actrl.aoi("Orphan").routine("Logic").aoi_calls[0].aoi == "Locked"


# --------------------------------------------------------------------------
# xref
# --------------------------------------------------------------------------
def test_xref_instances_and_call_sites(actrl, axref):
    assert sorted(t.name for t in axref.instances_of("Valve")) == ["XV1", "XV2", "XV3", "XV_Arr"]
    assert sorted((t.scope, t.name) for t in axref.instances_of("Scaler")) == [("Main", "Scl2"), ("aoi:Valve", "Wrk_Scl"), ("controller", "Scl1")]
    assert sorted(t.name for t in axref.instances_of("Locked")) == ["Lk1", "Wrk_Lk"]
    assert axref.instances_of("Orphan") == []
    sites = axref.calls_of("Valve")
    assert [s.where() for s in sites] == [
        "Main / MainRoutine / rung 0", "Main / MainRoutine / rung 1", "Main / StRoutine / line 1", "Main / FbdRoutine / sheet 1",
    ]
    assert all(s.instance_tag is not None for s in sites)
    assert {s.instance_tag.name for s in sites} == {"XV1", "XV_Arr"}
    xv1 = axref.tags["controller|xv1"]
    assert [s.call.source for s in axref.calls_of_instance(xv1)] == ["rll", "fbd"]
    assert axref.calls_of_instance(axref.tags["controller|xv2"]) == []
    scl = axref.calls_of("Scaler")
    assert {(s.scope, s.call.instance) for s in scl} == {("Main", "Scl1"), ("Main", "Scl2"), ("aoi:Valve", "Wrk_Scl")}
    assert instance_counts(actrl) == {"valve": 4, "scaler": 3, "locked": 2}


# --------------------------------------------------------------------------
# rules
# --------------------------------------------------------------------------
def test_rules_valve(afindings):
    valve = [f for f in afindings if f.program == "aoi:Valve"]
    rules = {(f.rule, f.tag) for f in valve}
    assert ("aoi_param_unused", "Spare_In") in rules
    assert next(f for f in valve if f.rule == "aoi_param_unused").severity == "warning"  # operando da chamada
    f = next(f for f in valve if f.rule == "aoi_output_never_written")
    assert f.tag == "Out_Never" and f.severity == "warning" and "Logic rung 5" in f.evidence
    f = next(f for f in valve if f.rule == "aoi_input_written")
    assert f.tag == "Inp_Written" and f.evidence == "Logic rung 3"  # a escrita do Prescan não conta
    assert ("aoi_local_unused", "Wrk_Unused") in rules
    assert ("aoi_local_write_only", "Wrk_Bit") in rules
    assert ("aoi_local_write_only", "Sts_HMI") not in rules and ("aoi_local_write_only", "Wrk_Scaled") not in rules  # visíveis ao HMI
    f = next(f for f in valve if f.rule == "aoi_local_read_only")
    assert f.tag == "Wrk_Const" and "1.5" in f.message
    assert ("aoi_local_read_only", "Cfg_HMI") not in rules  # editável pelo HMI
    f = next(f for f in valve if f.rule == "aoi_output_multiple_writes")
    assert f.tag == "Sts_Fault" and f.routine == "Logic" and f.rung == 2 and f.evidence == "rungs 2, 4"
    f = next(f for f in valve if f.rule == "aoi_unresolved_tags")
    assert f.routine == "Logic" and "Ghost" in f.evidence
    f = next(f for f in valve if f.rule == "aoi_params_no_description")
    assert "4 de 9" in f.message and "Spare_In" in f.evidence
    assert not any(f.rule in ("aoi_no_enableinfalse", "aoi_no_description", "aoi_unused") for f in valve)
    # instâncias: XV3 só lida por membro -> info; XV2 sem uso -> coberta por tag_unused, não duplicada
    f = next(f for f in valve if f.rule == "aoi_instance_never_called")
    assert f.tag == "XV3" and f.severity == "info" and "MainRoutine" in f.evidence
    assert not any(f.rule == "aoi_instance_never_called" and f.tag == "XV2" for f in afindings)
    assert any(f.rule == "tag_unused" and f.tag == "XV2" for f in afindings)
    assert all(not f.safety for f in valve)


def test_rules_arity_scaler_orphan_locked(afindings):
    f = next(f for f in afindings if f.rule == "aoi_call_arity")
    assert f.severity == "error" and (f.program, f.routine, f.rung, f.tag) == ("Main", "MainRoutine", 1, "XV_Arr[0]")
    assert "2 operando(s)" in f.message and "espera 5" in f.message and f.evidence == "Valve(XV_Arr[0], Fbk1, Out1)"
    assert sum(1 for f in afindings if f.rule == "aoi_call_arity") == 1  # a chamada ST com 5 argumentos está correta
    scaler = {f.rule for f in afindings if f.program == "aoi:Scaler"}
    assert {"aoi_no_description", "aoi_no_enableinfalse"} <= scaler
    assert "aoi_param_unused" not in scaler and "aoi_output_never_written" not in scaler  # ST: In lida, Out escrita
    orphan = {f.rule: f for f in afindings if f.program == "aoi:Orphan"}
    assert "aoi_unused" in orphan and orphan["aoi_unused"].severity == "info"
    assert orphan["aoi_nested_protected"].tag == "Wrk_Lk" and "Locked" in orphan["aoi_nested_protected"].message
    locked = {f.rule for f in afindings if f.program == "aoi:Locked"}
    assert locked == {"aoi_protected"}  # protegida: nada além do aviso de proteção


def test_metrics_and_minimal_regression(ametrics, findings):
    assert ametrics.aois == 4 and ametrics.aois_protected == 1 and ametrics.aois_unused == 1
    assert ametrics.aoi_instances == 9 and ametrics.aoi_calls == 9
    assert ametrics.findings_by_severity["error"] == 1
    # fixture minimal: Opt não é obrigatório -> info; SecretAoi sem instâncias
    opt = next(f for f in by_rule(findings, "aoi_param_unused") if f.tag == "Opt")
    assert opt.severity == "info" and opt.program == "aoi:MyAoi"
    assert by_rule(findings, "aoi_unused")[0].program == "aoi:SecretAoi"
    assert by_rule(findings, "aoi_protected")[0].program == "aoi:SecretAoi"


# --------------------------------------------------------------------------
# report + markdown
# --------------------------------------------------------------------------
def test_report(actrl, axref, afindings):
    rep = build_aoi_report(actrl, "valve", axref, afindings)
    assert rep.aoi.name == "Valve" and rep.scope == "aoi:Valve"
    assert [i.tag.name for i in rep.instances] == ["XV1", "XV2", "XV3", "XV_Arr"]
    xv1 = rep.instances[0]
    assert [s.call.source for s in xv1.calls] == ["rll", "fbd"]
    assert xv1.member_reads == {"Sts_Fault": ["Main / MainRoutine / rung 5", "Main / FbdRoutine / sheet 2"]}  # o IRef solto da folha 2 ainda conta como leitura
    assert rep.instances[2].member_reads and not rep.instances[2].calls
    assert len(rep.calls) == 4 and rep.calls_by_source() == {"rll": 2, "st": 1, "fbd": 1}
    assert rep.params["out_open"].writes == 3 and rep.params["out_open"].reads == 1
    assert rep.params["spare_in"].used is False and rep.locals["wrk_t"].reads >= 2
    assert rep.nested == ["Scaler"] and rep.used_by == []
    assert rep.rungs == 9 and rep.commented_rungs == 1 and rep.unresolved == ["Ghost"]
    rules = [f.rule for f in rep.findings]
    assert "aoi_call_arity" in rules and "tag_unused" in rules  # XV2 e a chamada errada entram no relatório da AOI
    assert rules == sorted(rules, key=lambda r: {"error": 0, "warning": 1, "info": 2}[next(f.severity for f in rep.findings if f.rule == r)])
    scl = build_aoi_report(actrl, "Scaler", axref, afindings)
    assert scl.used_by == ["Valve"] and scl.instances[1].member_writes == {"In": ["Main / StRoutine / line 2"]}
    # sem findings do projeto: as regras rodam só para a AOI
    alone = build_aoi_report(actrl, "Valve", axref)
    assert {f.rule for f in alone.findings} >= {"aoi_param_unused", "aoi_output_multiple_writes"}
    with pytest.raises(KeyError):
        build_aoi_report(actrl, "Nope", axref)


def test_markdown(actrl, axref, afindings):
    md = aoi_markdown(build_aoi_report(actrl, "Valve", axref, afindings))
    assert md.startswith("# AOI Valve — rev. 2.1 (Test)")
    for section in ("## Resumo", "## Interface (parâmetros)", "## Tags locais", "## Lógica interna", "## Instâncias", "## Pontos de chamada", "## Findings do parser", "## Roteiro de debug"):
        assert section in md, section
    assert "| Cfg_TravelT | Input | DINT | não | não | 5000 | L1/E0 | Tempo de curso (ms) |" in md
    assert "| Wrk_T | TIMER | {PRE=5000, ACC=0, EN=0, TT=0, DN=0} | " in md
    assert "Rung 0: XIC(Cmd_Open)XIO(Sts_Fault)OTE(Out_Open);" in md and "// Comando da válvula" in md
    assert "| XV2 | controlador | Valve | **nenhuma** |" in md
    assert "| XV_Arr | controlador | Valve[2] | Main/MainRoutine rung 1, Main/StRoutine line 1 |" in md
    assert "| Main / FbdRoutine / sheet 1 | FBD | XV1 | Ref_Cfg=Cfg1; Cmd_Open=Cmd2; Out_Open=Out1; Sts_Fault=Fault1 |" in md
    assert "| Main / MainRoutine / rung 0 | RLL | XV1 | Fbk_Open=Fbk1; Out_Open=Out1; Ref_Cfg=Cfg1; Spare_In=Spare1; Inp_Written=RawA |" in md
    assert "`aoi_call_arity`" in md and "`Out_Never` **nunca é escrita**" in md
    assert "`Sts_Fault` escrita em Logic rung 2, Logic rung 4, EnableInFalse rung 0 (OTE, OTU)" in md
    assert "`Wrk_T` (TON) preset 5000 (default da tag local), em Logic rung 1" in md
    assert "a rotina EnableInFalse executa" in md and "Prescan executa" in md
    assert "AOIs aninhadas (Scaler)" in md and "Ghost" in md
    locked = aoi_markdown(build_aoi_report(actrl, "Locked", axref, afindings))
    assert "AOI protegida" in locked and "## Lógica interna" not in locked and "| Lk1 | controlador |" in locked
    assert "Main / MainRoutine / rung 2 | RLL | Lk1 | (só a instância)" in locked
    small = aoi_markdown(build_aoi_report(actrl, "Valve", axref, afindings), max_rows=2, max_rungs=2)
    assert "(+2 instâncias)" in small and "(+5 rungs omitidos)" in small


def test_summary_rows(actrl, axref, afindings):
    rows = {r["aoi"]: r for r in aoi_summary_rows(actrl, axref, afindings)}
    assert rows["Valve"]["parameters"] == 9 and rows["Valve"]["locals"] == 8 and rows["Valve"]["instances"] == 4 and rows["Valve"]["calls"] == 4
    assert rows["Valve"]["nested"] == ["Scaler"] and rows["Valve"]["findings"]["warning"] >= 4
    assert rows["Locked"]["protected"] and rows["Locked"]["routines"] == "" and rows["Locked"]["findings"]["info"] == 1
    assert rows["Orphan"]["instances"] == 0 and rows["Orphan"]["calls"] == 0


# --------------------------------------------------------------------------
# serialize
# --------------------------------------------------------------------------
def test_routine_text_aoi_scope(actrl, axref):
    text = routine_text(actrl, "aoi:Valve", "Logic", axref).text
    assert text.startswith("# aoi:Valve / Logic [RLL]")
    assert "Chama AOIs: Scaler(Wrk_Scl)" in text
    assert "- Cfg_TravelT (DINT, parâmetro Input, default 5000; L1/E0) — Tempo de curso (ms)" in text
    assert "- Wrk_T (TIMER, tag local, default {PRE=5000, ACC=0, EN=0, TT=0, DN=0}; L" in text
    assert "- Out_Open (BOOL, parâmetro Output; L1/E3) — Saída para solenoide" in text
    with pytest.raises(KeyError):
        routine_text(actrl, "aoi:Nope", "Logic", axref)
    with pytest.raises(KeyError):
        routine_text(actrl, "aoi:Valve", "Nope", axref)
    # rotinas de programa continuam iguais e mostram as AOIs chamadas
    main = routine_text(actrl, "Main", "MainRoutine", axref).text
    assert main.startswith("# Main / MainRoutine [RLL]") and "Chama AOIs: Valve(XV1), Valve(XV_Arr[0]), Locked(Lk1)" in main


def test_aoi_text_and_overview(actrl, axref, afindings, ametrics):
    sr = aoi_interface_text(actrl, "Valve", axref)
    t = sr.text
    assert t.startswith("# AOI Valve rev. 2.1 (Test)") and "Nota de revisão: 2.1" in t
    assert "EnableInFalse: sim" in t
    assert "- Cfg_TravelT (Input, DINT, default 5000; L1/E0) — Tempo de curso (ms)" in t
    assert "- Fbk_Open (Input, BOOL, chamada; L1/E0) — Fim de curso aberto" in t
    assert "- Sts_HMI (DINT, HMI Read Only; L0/E1) — Status para o HMI" in t
    assert "## Instâncias (4): controlador: 4" in t and "- XV2 (controlador; chamada em NUNCA CHAMADA)" in t
    assert "- XV1 (controlador; chamada em Main/MainRoutine rung 0, Main/FbdRoutine sheet 1; lidos fora: Sts_Fault) — Válvula 1" in t
    assert "## Pontos de chamada (4): FBD 1, RLL 2, ST 1" in t
    assert "- Main/FbdRoutine sheet 1 [FBD] XV1: Ref_Cfg=Cfg1; Cmd_Open=Cmd2; Out_Open=Out1; Sts_Fault=Fault1" in t
    assert "Chama outras AOIs: Scaler" in t
    full = aoi_text(actrl, "Valve", axref).text
    assert "# aoi:Valve / Logic [RLL]" in full and "# aoi:Valve / Prescan [RLL]" in full and "# aoi:Valve / EnableInFalse [RLL]" in full
    locked = aoi_text(actrl, "Locked", axref).text
    assert locked.startswith("# AOI Locked rev. 1.0 (Vendor) [PROTEGIDA]") and "# aoi:Locked" not in locked  # protegida: só a interface
    short = aoi_interface_text(actrl, "Valve", axref, max_instances=1, max_calls=1).text
    assert "(+3 instâncias)" in short and "(+3 chamadas)" in short
    ov = project_overview(actrl, afindings, ametrics).text
    assert "- Valve rev. 2.1 (3 rotinas; 4 instâncias; operandos: Fbk_Open:I, Out_Open:O, Ref_Cfg:I, Spare_In:I, Inp_Written:I) — Válvula on/off" in ov
    assert "- Locked rev. 1.0 [PROTEGIDA] (0 rotinas; 2 instâncias; operandos: só tag de instância)" in ov


def test_cli_aoi_flags(tmp_path):
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--aois", "--aoi", "Valve", "--dump-routine", "aoi:Scaler/Logic"],
        capture_output=True, text=True, encoding="utf-8",
    )
    assert proc.returncode == 0, proc.stderr
    assert "4 AOIs" in proc.stdout and "# AOI Valve — rev. 2.1 (Test)" in proc.stdout and "## Roteiro de debug" in proc.stdout
    assert "# aoi:Scaler / Logic [ST]" in proc.stdout and "L1: Out := In * 2.0;" in proc.stdout
    txt = subprocess.run([sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--aoi", "Valve", "--aoi-text"], capture_output=True, text=True, encoding="utf-8")
    assert txt.returncode == 0 and "# AOI Valve rev. 2.1" in txt.stdout and "# aoi:Valve / Logic [RLL]" in txt.stdout and "tokens" in txt.stderr
    bad = subprocess.run([sys.executable, "-X", "utf8", "-m", "l5x_core.cli", AOI_L5X, "--aoi", "Nope"], capture_output=True, text=True, encoding="utf-8")
    assert bad.returncode == 2 and "não encontrada" in bad.stderr


# --------------------------------------------------------------------------
# LLM layer (fake client)
# --------------------------------------------------------------------------
def make(tmp_path, responder=None, **overrides):
    cfg = LLMConfig(model_fast="fast-model", model_deep="deep-model", log_path=str(tmp_path / "log.csv"), **overrides)
    fake = FakeAnthropic(responder)
    return LLMClient(cfg, client=fake), fake


def test_project_context_aoi_helpers(apctx):
    assert apctx.aoi_names() == ["Valve", "Scaler", "Orphan", "Locked"]
    rep = apctx.aoi_report("VALVE")
    assert rep.aoi.name == "Valve" and apctx.aoi_report("Valve") is rep  # cache, case-insensitive
    assert apctx.aoi_markdown("Valve").startswith("# AOI Valve")
    assert {f.rule for f in apctx.aoi_findings("Valve")} >= {"aoi_call_arity", "aoi_output_multiple_writes", "tag_unused"}
    assert apctx.aoi_routines("Valve") == [("aoi:Valve", "Logic"), ("aoi:Valve", "Prescan"), ("aoi:Valve", "EnableInFalse")]
    assert apctx.aoi_routines("Locked") == []
    assert apctx.routine_text("aoi:Valve", "Logic").text.startswith("# aoi:Valve / Logic")
    assert [r["aoi"] for r in apctx.aoi_rows()] == apctx.aoi_names()
    with pytest.raises(KeyError):
        apctx.aoi_report("Nope")


def test_llm_aoi_mode(apctx, tmp_path):
    assert "aoi" in MODES and "aoi" in DEEP_MODES
    prompt = load_prompt("aoi")
    assert "{interface}" in prompt and "{findings}" in prompt and "{content}" in prompt and "{target}" in prompt
    llm, fake = make(tmp_path)
    chunks = []
    r = llm.aoi(apctx, "Valve", on_text=chunks.append)
    assert r.mode == "aoi" and r.model == "deep-model" and r.calls == 1 and "".join(chunks) == r.text
    call = fake.calls[0]
    user = call["messages"][0]["content"]
    assert call["model"] == "deep-model" and call["system"][1]["cache_control"] == {"type": "ephemeral"}
    assert "AOI: Valve" in user and "# AOI Valve rev. 2.1" in user
    assert "aoi_call_arity" in user and "aoi_output_multiple_writes" in user
    assert "# aoi:Valve / Logic [RLL]" in user and "# aoi:Valve / EnableInFalse [RLL]" in user
    assert "<Rung" not in user and "CDATA" not in user
    with pytest.raises(LLMError) as exc:
        llm.aoi(apctx, "Locked")
    assert "source protection" in exc.value.message
    with pytest.raises(LLMError):
        llm.aoi(apctx, "Nope")


def test_llm_aoi_map_reduce(apctx, tmp_path, monkeypatch):
    llm, fake = make(tmp_path)
    original = LLMClient._pack_units

    def two_chunks(self, ctx, units, budget):
        return [ctx.routine_text(p, r).text for p, r in units][:2]

    monkeypatch.setattr(LLMClient, "_pack_units", two_chunks)
    r = llm.aoi(apctx, "Valve")
    assert r.calls == 3 and len(r.partial_results) == 2
    parts = [c["messages"][0]["content"] for c in fake.calls]
    assert "Valve — parte 1/2" in parts[0] and "Valve — parte 2/2" in parts[1]
    assert "resultados parciais" in parts[2] and 'modo "aoi"' in parts[2] and "AOI Valve" in parts[2]
    monkeypatch.setattr(LLMClient, "_pack_units", original)


# --------------------------------------------------------------------------
# retrieval + app helpers
# --------------------------------------------------------------------------
def test_retrieval_brings_aoi_routines(apctx):
    hits = score_routines(apctx, "explique a AOI Valve")
    keys = {(h.program, h.routine): h for h in hits}
    assert ("aoi:Valve", "Logic") in keys and "AOI Valve" in keys[("aoi:Valve", "Logic")].reasons
    assert ("aoi:Valve", "EnableInFalse") in keys
    hits = score_routines(apctx, "por que XV1 está em falha?")
    keys = {(h.program, h.routine): h for h in hits}
    assert ("Main", "MainRoutine") in keys and "tag XV1" in keys[("Main", "MainRoutine")].reasons
    assert "instância da AOI Valve" in keys[("aoi:Valve", "Logic")].reasons
    assert all(h.program != "aoi:Locked" for h in hits)
    sel = select_routines(apctx, "explique a AOI Valve", budget_tokens=100_000)
    assert sel[0][0].program == "aoi:Valve" and sel[0][1].text.startswith("# aoi:Valve / ")


def test_app_helpers():
    assert action_cost("aoi", "aoi") == 1
    for lang in UI:
        assert "aoi_analyze" in UI[lang] and "mode_aoi" in UI[lang]

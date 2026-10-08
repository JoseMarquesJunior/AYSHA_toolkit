import pytest

from l5x_core.rung_text import (
    parse_rung, tokenize, classify_operand, tags_in_expression, RungSyntaxError,
)
from l5x_core.model import READ, WRITE, READWRITE, NONE


def names(instructions):
    return [i.name for i in instructions]


def test_simple_rung():
    ins, depth = parse_rung("XIC(Start)XIO(Stop)OTE(Motor);")
    assert names(ins) == ["XIC", "XIO", "OTE"]
    assert depth == 0
    assert ins[0].operands[0].access == READ
    assert ins[2].operands[0].access == WRITE
    assert ins[2].operands[0].tag_base == "Motor"
    assert all(not i.in_branch for i in ins)


def test_nested_branches_and_depth():
    text = "XIC(A)[XIO(B)[OTE(C) ,XIC(D)OTL(E) ] ,MOV(F,G) ];"
    ins, depth = parse_rung(text)
    assert depth == 2
    assert names(ins) == ["XIC", "XIO", "OTE", "XIC", "OTL", "MOV"]
    assert [i.branch_depth for i in ins] == [0, 1, 2, 2, 2, 1]
    assert ins[2].in_branch and not ins[0].in_branch


def test_literal_with_comma_inside_quotes():
    ins, _ = parse_rung("CONCAT(StrTag,'a,b',StrTag)EQU(X,'x)y');")
    assert len(ins[0].operands) == 3
    assert ins[0].operands[1].kind == "literal"
    assert ins[0].operands[1].raw == "'a,b'"
    assert ins[0].operands[2].access == WRITE
    assert ins[1].operands[1].raw == "'x)y'"


def test_unused_question_mark_and_timer_roles():
    ins, _ = parse_rung("TON(T1,?,?)CTU(C1,10,0);")
    ton = ins[0]
    assert ton.operands[0].access == READWRITE
    assert ton.operands[1].kind == "unused"
    assert ton.operands[1].access == NONE
    ctu = ins[1]
    assert ctu.operands[1].kind == "literal"
    assert ctu.operands[2].kind == "literal"


def test_tag_paths_members_arrays_and_indices():
    op = classify_operand("Alarmes[3].Ativo")
    assert op.kind == "tag" and op.tag_base == "Alarmes" and op.tag_path == "Alarmes[3].Ativo"
    op = classify_operand("Arr[Idx]")
    assert op.tag_base == "Arr" and op.extra_tags == ["Idx"]
    op = classify_operand("Tanque.Nivel")
    assert op.tag_base == "Tanque"
    op = classify_operand("Local:1:I.Data.5")
    assert op.kind == "tag" and op.tag_base == "Local:1:I" and op.tag_path == "Local:1:I.Data.5"
    op = classify_operand("Word.3")
    assert op.kind == "tag" and op.tag_path == "Word.3"


def test_cross_program_reference():
    op = classify_operand("\\Prog1.Tag.Bit")
    assert op.kind == "tag"
    assert op.program == "Prog1"
    assert op.tag_base == "Tag"
    assert op.tag_path == "Tag.Bit"


def test_literals():
    for raw in ("10", "1.5", "-3", "1.0e+3", "16#00FF", "2#0101", "'texto'"):
        assert classify_operand(raw).kind == "literal", raw


def test_expression_operand_extracts_tags():
    ins, _ = parse_rung("CPT(Dest,Arr[Idx]*2+Tank.Level MOD 16#FF - ABS(X));")
    cpt = ins[0]
    assert cpt.operands[0].access == WRITE and cpt.operands[0].tag_base == "Dest"
    expr = cpt.operands[1]
    assert expr.kind == "expression"
    assert set(expr.extra_tags) == {"Arr[Idx]", "Idx", "Tank.Level", "X"}
    assert tags_in_expression("A AND B OR NOT C") == ["A", "B", "C"]


def test_jsr_roles():
    ins, _ = parse_rung("JSR(Sub,2,In1,In2,Ret1);")
    jsr = ins[0]
    assert jsr.operands[0].kind == "keyword"
    assert jsr.operands[1].kind == "literal"
    assert jsr.operands[2].access == READ
    assert jsr.operands[3].access == READ
    assert jsr.operands[4].access == WRITE


def test_gsv_ssv_keywords():
    ins, _ = parse_rung("GSV(PROGRAM,THIS,LASTSCANTIME,Arr[0])SSV(Module,Mod1,Mode,Val);")
    gsv, ssv = ins
    assert [o.kind for o in gsv.operands[:3]] == ["keyword"] * 3
    assert gsv.operands[3].access == WRITE
    assert ssv.operands[3].access == READ


def test_aoi_roles_from_usage():
    usages = {"myaoi": ["Input", "Output", "InOut"]}
    ins, _ = parse_rung("MyAoi(Inst,In1,Out1,IO1);", usages)
    aoi = ins[0]
    assert aoi.is_aoi and aoi.known
    assert [o.access for o in aoi.operands] == [READWRITE, READ, WRITE, READWRITE]


def test_unknown_instruction_all_reads():
    ins, _ = parse_rung("FOO(A,B);")
    assert ins[0].known is False
    assert [o.access for o in ins[0].operands] == [READ, READ]


def test_empty_and_nop():
    ins, _ = parse_rung("NOP();")
    assert ins[0].name == "NOP" and ins[0].operands == []


def test_mov_cop_roles():
    ins, _ = parse_rung("MOV(Src,Dst)COP(Src,Dst,5)CLR(X)ADD(A,B,C);")
    assert ins[0].operands[1].access == WRITE
    assert ins[1].operands[1].access == WRITE and ins[1].operands[2].kind == "literal"
    assert ins[2].operands[0].access == WRITE
    assert ins[3].operands[2].access == WRITE


def test_multiline_text_and_spaces():
    ins, depth = parse_rung("XIC(I_bOn)[XIO(O_bTimerDone) RTO(Wrk_TotRun,?,?) ,XIC(Wrk_TotRun.TT) OTE(O_bTROn) ];\n")
    assert names(ins) == ["XIC", "XIO", "RTO", "XIC", "OTE"]
    assert depth == 1


@pytest.mark.parametrize("bad", [
    "XIC(Start OTE(Lamp);",
    "XIC(A)]OTE(B);",
    "XIC(A)[OTE(B);",
    "XIC(A) , OTE(B);",
    "XIC(A)OTE(B); XIC(C);",
    "42(A);",
])
def test_syntax_errors(bad):
    with pytest.raises(RungSyntaxError):
        tokenize(bad)

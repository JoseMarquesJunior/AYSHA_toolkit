"""Decodificação dos dados L5K das instâncias de AOI (tags locais que o Decorated não traz)."""

import pytest

from l5x_core import parse_bytes
from l5x_core.l5k import L5KError, Layouts, parse_l5k, validate, merge
from l5x_core.alarms import alarm_rows


def test_parse_l5k_tokens():
    assert parse_l5k("[3,0,3,[1496478,0,0],[1496478,3000,0]]") == [3, 0, 3, [1496478, 0, 0], [1496478, 3000, 0]]
    assert parse_l5k("[-3.62500119e+000,2#0000_0101,16#ff,[7,'P_Alarm$00$00'],'a$'b$$c']") == [-3.62500119, 5, 255, [7, "P_Alarm"], "a'b$c"]
    assert parse_l5k("  [ 1 ,\n 2 ] ") == [1, 2]
    with pytest.raises(L5KError):
        parse_l5k("[1,2")


XML = """<?xml version="1.0" encoding="UTF-8"?>
<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="35.00" TargetName="L5K" TargetType="Controller">
<Controller Use="Target" Name="L5K" ProcessorType="1756-L85E" MajorRev="35" MinorRev="11">
<DataTypes>
<DataType Name="Flags" Family="NoFamily" Class="User">
<Members>
<Member Name="ZZZZZZZZZZFlags0" DataType="SINT" Dimension="0" Radix="Decimal" Hidden="true"/>
<Member Name="A" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZFlags0" BitNumber="0"/>
<Member Name="B" DataType="BIT" Dimension="0" Radix="Decimal" Hidden="false" Target="ZZZZZZZZZZFlags0" BitNumber="1"/>
<Member Name="Count" DataType="DINT" Dimension="0" Radix="Decimal" Hidden="false"/>
</Members>
</DataType>
<DataType Name="STRING_8" Family="StringFamily" Class="User">
<Members>
<Member Name="LEN" DataType="DINT" Dimension="0" Radix="Decimal" Hidden="false"/>
<Member Name="DATA" DataType="SINT" Dimension="8" Radix="ASCII" Hidden="false"/>
</Members>
</DataType>
</DataTypes>
<AddOnInstructionDefinitions>
<AddOnInstructionDefinition Name="P_Alm" Revision="1.0" ExecutePrescan="false" ExecutePostscan="false" ExecuteEnableInFalse="false">
<Parameters>
<Parameter Name="EnableIn" TagType="Base" DataType="BOOL" Usage="Input" Required="false" Visible="false"/>
<Parameter Name="EnableOut" TagType="Base" DataType="BOOL" Usage="Output" Required="false" Visible="false"/>
<Parameter Name="Inp" TagType="Base" DataType="BOOL" Usage="Input" Required="true" Visible="true"/>
<Parameter Name="Cfg_Severity" TagType="Base" DataType="INT" Usage="Input" Required="false" Visible="false"/>
<Parameter Name="Cfg_Exists" TagType="Base" DataType="BOOL" Usage="Input" Required="false" Visible="false"/>
</Parameters>
<LocalTags>
<LocalTag Name="Wrk_T" DataType="TIMER" ExternalAccess="None"/>
<LocalTag Name="Cfg_Tag" DataType="STRING_8" ExternalAccess="Read/Write"/>
</LocalTags>
<Routines><Routine Name="Logic" Type="RLL"><RLLContent><Rung Number="0" Type="N"><Text><![CDATA[XIC(Inp)OTE(EnableOut);]]></Text></Rung></RLLContent></Routine></Routines>
</AddOnInstructionDefinition>
<AddOnInstructionDefinition Name="Big" Revision="1.0" ExecutePrescan="false" ExecutePostscan="false" ExecuteEnableInFalse="false">
<Parameters>
<Parameter Name="EnableIn" TagType="Base" DataType="BOOL" Usage="Input" Required="false" Visible="false"/>
<Parameter Name="EnableOut" TagType="Base" DataType="BOOL" Usage="Output" Required="false" Visible="false"/>
<Parameter Name="Inp_PV" TagType="Base" DataType="REAL" Usage="Input" Required="true" Visible="true"/>
%s
<Parameter Name="Cfg_RangeMax" TagType="Base" DataType="REAL" Usage="Input" Required="false" Visible="false"/>
</Parameters>
<LocalTags>
<LocalTag Name="B33" DataType="BOOL" ExternalAccess="None"/>
<LocalTag Name="AH_Alm" DataType="P_Alm" ExternalAccess="None"><Description><![CDATA[Alto]]></Description></LocalTag>
<LocalTag Name="Fl" DataType="Flags" ExternalAccess="None"/>
<LocalTag Name="Arr" DataType="DINT" Dimensions="3" ExternalAccess="None"/>
</LocalTags>
<Routines><Routine Name="Logic" Type="RLL"><RLLContent><Rung Number="0" Type="N"><Text><![CDATA[XIC(EnableIn)OTE(EnableOut);]]></Text></Rung></RLLContent></Routine></Routines>
</AddOnInstructionDefinition>
</AddOnInstructionDefinitions>
<Tags>
<Tag Name="Alm1" TagType="Base" DataType="P_Alm" Constant="false">
<Data Format="L5K"><![CDATA[[5,250,[-1610612736,5000,12],[3,'abc$00$00$00$00$00']]]]></Data>
<Data Format="Decorated"><Structure DataType="P_Alm">
<DataValueMember Name="EnableIn" DataType="BOOL" Value="1"/>
<DataValueMember Name="EnableOut" DataType="BOOL" Value="0"/>
<DataValueMember Name="Inp" DataType="BOOL" Value="1"/>
<DataValueMember Name="Cfg_Severity" DataType="INT" Radix="Decimal" Value="250"/>
<DataValueMember Name="Cfg_Exists" DataType="BOOL" Value="0"/>
</Structure></Data>
</Tag>
<Tag Name="Bad1" TagType="Base" DataType="P_Alm" Constant="false">
<Data Format="L5K"><![CDATA[[5,999,[0,0,0],[0,'']]]]></Data>
<Data Format="Decorated"><Structure DataType="P_Alm">
<DataValueMember Name="Cfg_Severity" DataType="INT" Radix="Decimal" Value="250"/>
</Structure></Data>
</Tag>
<Tag Name="Big1" TagType="Base" DataType="Big" Constant="false">
<Data Format="L5K"><![CDATA[[3,1.5,100.0,1,[7,500,[0,1,0],[0,'']],[2,42],[10,20,30]]]]></Data>
<Data Format="Decorated"><Structure DataType="Big">
<DataValueMember Name="EnableIn" DataType="BOOL" Value="1"/>
<DataValueMember Name="Inp_PV" DataType="REAL" Radix="Float" Value="1.5"/>
<DataValueMember Name="Cfg_RangeMax" DataType="REAL" Radix="Float" Value="100.0"/>
</Structure></Data>
</Tag>
<Tag Name="AlmArr" TagType="Base" DataType="P_Alm" Dimensions="2" Constant="false">
<Data Format="L5K"><![CDATA[[[1,10,[0,1,0],[0,'']],[0,20,[0,2,0],[0,'']]]]]></Data>
<Data Format="Decorated"><Array DataType="P_Alm" Dimensions="2">
<Element Index="[0]"><Structure DataType="P_Alm"><DataValueMember Name="Cfg_Severity" DataType="INT" Radix="Decimal" Value="10"/></Structure></Element>
<Element Index="[1]"><Structure DataType="P_Alm"><DataValueMember Name="Cfg_Severity" DataType="INT" Radix="Decimal" Value="20"/></Structure></Element>
</Array></Data>
</Tag>
</Tags>
<Programs/>
<Tasks/>
</Controller>
</RSLogix5000Content>
"""
# 30 BOOL parameters so that, with EnableIn/EnableOut, the 33rd scalar BOOL (B33) opens a second packing DINT
BOOLS = "\n".join(f'<Parameter Name="F{i:02d}" TagType="Base" DataType="BOOL" Usage="Input" Required="false" Visible="false"/>' for i in range(30))


@pytest.fixture(scope="module")
def ctrl():
    return parse_bytes((XML % BOOLS).encode("utf-8"))


def test_decode_simple_instance(ctrl):
    tag = next(t for t in ctrl.tags if t.name == "Alm1")
    v = tag.values
    # parâmetros: Decorated prevalece; locais vêm do L5K
    assert (v["EnableIn"], v["EnableOut"], v["Inp"], v["Cfg_Severity"], v["Cfg_Exists"]) == (1, 0, 1, 250, 0)
    assert v["Wrk_T"] == {"EN": 1, "TT": 0, "DN": 1, "PRE": 5000, "ACC": 12}
    assert v["Cfg_Tag"] == {"LEN": 3, "DATA": "abc"}


def test_mismatch_is_rejected_and_reported(ctrl):
    bad = next(t for t in ctrl.tags if t.name == "Bad1")
    assert bad.values == {"Cfg_Severity": 250} and "Wrk_T" not in bad.values
    f = [f for f in ctrl.parse_findings if f.rule == "l5k_decode_mismatch"]
    assert len(f) == 1 and f[0].tag == "Bad1" and "Cfg_Severity" in f[0].message and f[0].program == "aoi:P_Alm"


def test_packing_overflow_nested_udt_and_array(ctrl):
    v = next(t for t in ctrl.tags if t.name == "Big1").values
    assert v["EnableIn"] == 1 and v["EnableOut"] == 1 and v["F00"] == 0 and v["Inp_PV"] == 1.5 and v["Cfg_RangeMax"] == 100.0
    assert v["B33"] == 1  # 33º BOOL: novo DINT de empacotamento na posição dele
    assert v["AH_Alm"]["Cfg_Severity"] == 500 and v["AH_Alm"]["EnableIn"] == 1 and v["AH_Alm"]["Inp"] == 1 and v["AH_Alm"]["Wrk_T"]["PRE"] == 1
    assert v["Fl"] == {"A": 0, "B": 1, "Count": 42}  # UDT com SINT oculto e BitNumber
    assert v["Arr"] == {"[0]": 10, "[1]": 20, "[2]": 30}
    arr = next(t for t in ctrl.tags if t.name == "AlmArr").values
    assert arr["[0]"]["Cfg_Severity"] == 10 and arr["[0]"]["Wrk_T"]["PRE"] == 1 and arr["[1]"]["Wrk_T"]["PRE"] == 2 and arr["[1]"]["EnableIn"] == 0


def test_nested_alarm_reaches_alarm_list(ctrl):
    header, rows = alarm_rows(ctrl)
    by = {r["Alarme"]: r for r in rows}
    assert by["Big1.AH_Alm"]["Severidade"] == 500 and by["Big1.AH_Alm"]["DescricaoMembro"] == "Alto"
    assert by["Alm1"]["Severidade"] == 250 and by["AlmArr"]["Severidade"] == ""  # array de alarmes: linha única sem severidade única


def test_helpers():
    assert validate({"a": 1.0000001, "b": {"x": 1}}, {"a": 1, "b": {"x": 1}}) is None
    assert validate({"a": 2}, {"a": 1}) == "a"
    nan = parse_l5k("[1.#QNAN000e+000,-1.#INF,2#0000_0101]")
    assert nan[0] != nan[0] and nan[1] == float("-inf") and nan[2] == 5
    assert validate({"a": nan[0]}, {"a": "1.#QNAN"}) is None and validate({"a": 1.0}, {"a": "1.#QNAN"}) == "a"
    assert validate({"w": -1521025024}, {"w": 2773942272}) is None  # mesmo padrão de bits (DINT com bit 31)
    from l5x_core.parser import _value
    assert _value("2#1010_0101_0101_0101_0000_0000_0000_0000", "DINT") == -1521025024 and _value("16#FF", "SINT") == -1 and _value("16#FF", "USINT") == 255 and _value("16#FF") == 255
    assert merge({"a": 1, "s": {"x": 1}}, {"a": 9, "s": {"x": 2, "y": 3}, "loc": 7}) == {"a": 1, "s": {"x": 1, "y": 3}, "loc": 7}
    lay = Layouts(parse_bytes((XML % BOOLS).encode("utf-8")))
    assert lay.decode([5, 10, 7], "COUNTER") == {"CU": 0, "CD": 0, "DN": 0, "OV": 0, "UN": 0, "PRE": 10, "ACC": 7}
    assert lay.decode([[1, 2], [3, 4]], "DINT", "2,2") == [[1, 2], [3, 4]]  # multidimensional fica bruto
    assert lay.decode([1, 2], "Unknown_T") == [1, 2]

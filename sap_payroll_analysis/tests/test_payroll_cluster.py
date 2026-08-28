"""Fase 3 (sem SAP): RGDIR parsing, contexto, tentativa de RT, retro, run 1299."""

from dataclasses import replace
from decimal import Decimal
from types import SimpleNamespace

import pytest

from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.payroll_cluster import (
    PayrollDirectoryEntry,
    analyse_cluster,
    attempt_read_rt,
    build_timeline,
    collect_payroll_context,
    discover_payroll_context,
    discover_payroll_result_tables,
    read_rgdir,
)
from sap_payroll_analysis.security import SecurityError, safe_rfc_call
from sap_payroll_analysis.tests.fakes import FakeConnection, FakeRfcError

PARAMS = replace(DEFAULTS, empresa="1010", conta="23120000", primary_run="0000001298")

RGDIR_F = [("PERNR", "N", 8, ""), ("SEQNR", "N", 5, ""), ("ABKRS", "C", 2, ""),
           ("FPPER", "N", 6, ""), ("FPBEG", "D", 8, ""), ("FPEND", "D", 8, ""),
           ("INPER", "N", 6, ""), ("IPEND", "D", 8, ""), ("SRTZA", "C", 1, ""),
           ("PAYTY", "C", 1, ""), ("PAYID", "C", 1, ""), ("VOID", "C", 1, ""), ("PERMO", "C", 2, "")]
PA0001_F = [("PERNR", "N", 8, ""), ("BEGDA", "D", 8, ""), ("ENDDA", "D", 8, ""),
            ("ABKRS", "C", 2, ""), ("BUKRS", "C", 4, "")]
T549A_F = [("ABKRS", "C", 2, ""), ("PERMO", "C", 2, "")]
T500L_F = [("MOLGA", "C", 2, ""), ("RELID", "C", 2, ""), ("INTCA", "C", 2, ""), ("LAND1", "C", 3, "")]


def _rg(pernr, seqnr, fpper, inper, srtza="A", payty="", void=""):
    return {"PERNR": pernr, "SEQNR": seqnr, "ABKRS": "Z2", "FPPER": fpper,
            "FPBEG": fpper + "01", "FPEND": fpper + "30", "INPER": inper, "IPEND": inper + "30",
            "SRTZA": srtza, "PAYTY": payty, "PAYID": "", "VOID": void, "PERMO": "01"}


def _base_tables():
    return {
        "HRPY_RGDIR": {"fields": RGDIR_F, "rows": [
            _rg("00000005", "01100", "202606", "202606"),          # só componente do período
            _rg("00000009", "01129", "202605", "202606"),          # componente retro (maio)
            _rg("00000009", "01130", "202606", "202606", "P"),     # + componente do período
            _rg("00000017", "01130", "202605", "202606"),          # só componente retro
            _rg("00000005", "00050", "202512", "202512"),          # fora do IN-period do run
        ]},
        "PA0001": {"fields": PA0001_F, "rows": [
            {"PERNR": "00000005", "BEGDA": "20200101", "ENDDA": "99991231", "ABKRS": "Z2", "BUKRS": "1010"},
            {"PERNR": "00000009", "BEGDA": "20200101", "ENDDA": "99991231", "ABKRS": "Z2", "BUKRS": "1010"},
        ]},
        "T549A": {"fields": T549A_F, "rows": [{"ABKRS": "Z2", "PERMO": "01"}]},
        "T500L": {"fields": T500L_F, "rows": [
            {"MOLGA": "01", "RELID": "RD", "INTCA": "DE", "LAND1": "DE"},
            {"MOLGA": "19", "RELID": "RP", "INTCA": "PT", "LAND1": "PT"},
        ]},
    }


def _link(sample_pernrs_amounts):
    """link_report stand-in com o mínimo que analyse_cluster consome."""
    link_sample = []
    for pn, amt in sample_pernrs_amounts:
        link_sample.append({"PERNR": pn, "LGART": "/559", "BETRG": amt, "KOMOK": "S003",
                            "POSTNUM": "00013", "RTLINE": "00055", "TSLIN": "0000000004",
                            "DOCNUM": "0000005392", "DOCLIN": "0000000326"})
    ref_total = sum(Decimal(a.rstrip("-")) * (-1 if a.endswith("-") else 1)
                    for _, a in sample_pernrs_amounts)
    return SimpleNamespace(
        resolved=True, link_sample=link_sample,
        ppoix_total=ref_total, posting_line_amount=ref_total + Decimal("265.65"),
        reference_total=ref_total, by_wage_type={"/559": {"rows": len(link_sample), "amount": str(ref_total)}},
        warnings=[],
    )


# --------------------------------------------------------------------------- RGDIR

def test_directory_entry_retro_flag():
    e = PayrollDirectoryEntry(pernr="9", seqnr="1", fpper="202605", inper="202606")
    assert e.is_retro is True
    e2 = PayrollDirectoryEntry(pernr="9", seqnr="2", fpper="202606", inper="202606")
    assert e2.is_retro is False


def test_read_rgdir_parses_and_filters():
    conn = FakeConnection(_base_tables())
    entries = read_rgdir(conn, ["00000005", "00000009", "00000017"])
    assert len(entries) == 5
    a_june = [e for e in entries if e.inper == "202606" and e.srtza == "A"]
    assert {e.pernr for e in a_june} == {"00000005", "00000009", "00000017"}
    assert sum(1 for e in a_june if e.is_retro) == 2


def test_discover_context_picks_pt():
    conn = FakeConnection(_base_tables())
    ctx = discover_payroll_context(conn, PARAMS, ["00000005", "00000009"])
    assert ctx["abkrs"] == "Z2"
    assert ctx["permo"] == "01"
    assert ctx["molga"] == "19"
    assert ctx["relid"] == "RP"


# --------------------------------------------------------------------------- RT attempt

def test_attempt_read_rt_reports_da300_limitation():
    conn = FakeConnection(
        _base_tables(),
        functions={"PYXX_READ_PAYROLL_RESULT": FakeRfcError("ID:DA Type:E Number:300")},
    )
    att = attempt_read_rt(conn, PARAMS, "00000009", "01129", "RP")
    assert att.attempted is True
    assert att.ok is False
    assert "300" in att.reason
    assert "nametab" in att.detail.lower()


def test_attempt_read_rt_success_path():
    payload = {"PAYROLL_RESULT": {"INTER": {"RT": [
        {"LGART": "/559", "BETRG": "1000.00"}, {"LGART": "/558", "BETRG": "10.00"}]}}}
    conn = FakeConnection(_base_tables(), functions={"PYXX_READ_PAYROLL_RESULT": payload})
    att = attempt_read_rt(conn, PARAMS, "00000009", "01129", "RP")
    assert att.ok is True
    assert len(att.sample) == 2


def test_payroll_result_fms_are_not_write_blocked_but_gated():
    # PYXX_READ_PAYROLL_RESULT está na whitelist (leitura); HR_GET_PAYROLL_RESULTS não.
    conn = FakeConnection({}, functions={"PYXX_READ_PAYROLL_RESULT": {"PAYROLL_RESULT": {}}})
    safe_rfc_call(conn, "PYXX_READ_PAYROLL_RESULT", EMPLOYEENUMBER="1", SEQUENCENUMBER="1", CLUSTERID="RP")
    with pytest.raises(SecurityError):
        safe_rfc_call(conn, "HR_GET_PAYROLL_RESULTS", PERNR="1")
    with pytest.raises(SecurityError):
        safe_rfc_call(conn, "HR_UPDATE_PAYROLL_RESULT")  # tokens de escrita


# --------------------------------------------------------------------------- analyse_cluster

def test_analyse_cluster_classifies_retro_and_residuals():
    tables = _base_tables()
    conn = FakeConnection(
        tables,
        functions={"PYXX_READ_PAYROLL_RESULT": FakeRfcError("ID:DA Type:E Number:300")},
    )
    link = _link([("00000005", "100.00-"), ("00000009", "724000.00-"), ("00000017", "374.38-")])
    rep = analyse_cluster(conn, PARAMS, payroll_report=SimpleNamespace(), link_report=link)

    assert rep.resolved
    assert rep.molga == "19" and rep.abkrs == "Z2" and rep.relid == "RP"
    # 9 e 17 têm componente retro; 5 e 9 têm componente do próprio período
    assert set(rep.retro_pernr) == {"00000009", "00000017"}
    assert set(rep.current_pernr) == {"00000005", "00000009"}
    # buckets mutuamente exclusivos + "mixed"
    assert rep.ppoix_ref_retro_total == Decimal("-374.38")      # só 17
    assert rep.ppoix_ref_current_total == Decimal("-100.00")    # só 5
    assert rep.residual_notes.get("ppoix_ref_mixed_current_and_retro") == "-724000.00"  # 9
    assert rep.ppoix_ref_total == Decimal("-724474.38")
    # timeline + pares construídos
    assert any(tl.pernr == "00000009" for tl in rep.timelines)
    # RGDIR classificação registada
    assert rep.classification_distribution
    assert rep.rt_attempt.ok is False
    assert "ppoix_vs_ppdit" in rep.residual_notes
    assert rep.run_1299_comparison.get("ok") is False


def test_build_timeline_pairs_and_classification():
    e = [
        PayrollDirectoryEntry(pernr="9", seqnr="10", fpper="202605", inper="202605", srtza="O"),
        PayrollDirectoryEntry(pernr="9", seqnr="11", fpper="202605", inper="202606", srtza="A"),
        PayrollDirectoryEntry(pernr="9", seqnr="12", fpper="202606", inper="202606", srtza="A"),
    ]
    tl = build_timeline("9", e)
    assert [p.fpper for p in tl.pairs] == ["202605", "202606"]
    may = tl.pairs[0]
    assert may.status == "RESULT_RECALCULATED"
    assert may.in_periods == ["202605", "202606"]
    assert may.original.seqnr == "10"
    assert may.current.seqnr == "11"
    jun = tl.pairs[1]
    assert jun.status == "RESULT_UNCHANGED"
    # classificação por entrada (RETRO_LAG = desfasamento de rotina de 1 mês)
    assert e[0].classify() == "ORIGINAL/OLD"
    assert e[1].classify() == "RETRO_LAG/CURRENT"
    assert e[1].months_late == 1
    assert e[2].classify() == "ORIGINAL/CURRENT"
    e_corr = PayrollDirectoryEntry(pernr="9", seqnr="20", fpper="202601", inper="202606", srtza="A")
    assert e_corr.classify() == "RETRO_CORR/CURRENT"
    assert e_corr.months_late == 5


def test_collect_payroll_context_is_automatic():
    """A partir do run, deriva PERNR + RGDIR + PA0001 sem input manual."""
    from sap_payroll_analysis.tests.test_wagetype_link import _conn as _wt_conn, PARAMS as WT_PARAMS

    conn = _wt_conn()
    conn.tables.update({k: v for k, v in _base_tables().items() if k in ("PA0001", "T549A", "T500L")})
    conn.tables["HRPY_RGDIR"] = {"fields": RGDIR_F, "rows": [
        _rg(p, "01000", "202605", "202606")
        for p in sorted({r["PERNR"] for r in conn.tables["PPOIX"]["rows"]})]}

    ctx = collect_payroll_context(conn, WT_PARAMS, runid="0000001298")
    assert ctx.resolved
    assert ctx.pernrs  # derivados do posting
    assert ctx.doc_lines == [("0000005392", "0000000326")]
    assert all(pn in ctx.rgdir_by_pernr for pn in ctx.pernrs)
    assert ctx.abkrs == "Z2" and ctx.relid == "RP"
    # PPOIX por rubrica presente
    assert any("/559" in wt for wt in ctx.ppoix_by_pernr_wt.values())


def test_discover_payroll_result_tables_marks_empty_and_missing():
    tables = _base_tables()
    tables["DD02L"] = {"fields": [("TABNAME", "C", 30, ""), ("TABCLASS", "C", 6, "")], "rows": [
        {"TABNAME": "P2RX_RT", "TABCLASS": "TRANSP"},
        {"TABNAME": "HRPY_RGDIR", "TABCLASS": "TRANSP"},
    ]}
    tables["DD03L"] = {"fields": [("TABNAME", "C", 30, ""), ("FIELDNAME", "C", 30, "")], "rows": [
        {"TABNAME": "P2RX_RT", "FIELDNAME": "LGART"}, {"TABNAME": "P2RX_RT", "FIELDNAME": "BETRG"},
        {"TABNAME": "HRPY_RGDIR", "FIELDNAME": "PERNR"},
    ]}
    tables["P2RX_RT"] = {"fields": [("LGART", "C", 4, ""), ("BETRG", "P", 15, "")], "rows": []}
    conn = FakeConnection(tables)
    cat = {t.table: t for t in discover_payroll_result_tables(conn)}
    assert cat["P2RX_RT"].accessible is True
    assert cat["P2RX_RT"].populated is False           # existe no DDIC mas vazia
    assert cat["HRPY_RGDIR"].accessible is True
    assert cat["HRPY_RGDIR"].populated is True         # o fake tem linhas
    # uma candidata inexistente no fake
    assert cat["P2RX_CRT"].accessible is False


def test_analyse_cluster_run1299_repeat_posting():
    """Fake com 1298 e 1299 idênticos -> classificado como REPEAT_POSTING."""
    from sap_payroll_analysis.tests.test_wagetype_link import _conn as _wt_conn, PARAMS as WT_PARAMS
    from sap_payroll_analysis.payroll_posting import analyze as analyze_payroll
    from sap_payroll_analysis.payroll_wagetypes import link_wage_types_to_posting_line

    conn = _wt_conn()
    # acrescenta run 1299 = cópia do 1298 (doc 5394/326) + RGDIR/PA0001/T500L/T549A
    conn.tables["PPDHD"]["rows"].append(
        {"DOCNUM": "0000005394", "RUNID": "0000001299", "BUKRS": "1010", "BUDAT": "20260630",
         "DOCTYP": "1", "REVDOC": "0000000000", "XBLNR": "HRPAY00001"})
    for r in list(conn.tables["PPDIT"]["rows"]):
        if r["DOCNUM"] == "0000005392":
            r2 = dict(r); r2["DOCNUM"] = "0000005394"; conn.tables["PPDIT"]["rows"].append(r2)
    for r in list(conn.tables["PPDIX"]["rows"]):
        r2 = dict(r); r2["RUNID"] = "0000001299"; r2["DOCNUM"] = "0000005394"
        conn.tables["PPDIX"]["rows"].append(r2)
    for r in list(conn.tables["PPOIX"]["rows"]):
        r2 = dict(r); r2["RUNID"] = "0000001299"; conn.tables["PPOIX"]["rows"].append(r2)
    conn.tables.update({k: v for k, v in _base_tables().items()
                        if k in ("PA0001", "T549A", "T500L")})
    conn.tables["HRPY_RGDIR"] = {"fields": RGDIR_F, "rows": [
        _rg(p, "01000", "202605", "202606") for p in
        {r["PERNR"] for r in conn.tables["PPOIX"]["rows"]}]}
    conn.functions = {"PYXX_READ_PAYROLL_RESULT": FakeRfcError("ID:DA Type:E Number:300")}

    payroll = analyze_payroll(conn, WT_PARAMS)
    link = link_wage_types_to_posting_line(conn, WT_PARAMS, payroll)
    rep = analyse_cluster(conn, WT_PARAMS, payroll, link)
    assert rep.run_1299_comparison.get("ok") is True
    assert rep.run_1299_comparison["classification"].startswith("REPEAT_POSTING")
    assert rep.run_1299_comparison["same_pernr_set"] is True

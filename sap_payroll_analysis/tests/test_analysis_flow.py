"""Fluxo (sem SAP): PPDHD->PPDIT, FI via BSIS, PPOIX e reconciliação.

A FakeConnection ignora OPTIONS (WHERE); as linhas fornecidas já
representam o universo filtrado.
"""

from decimal import Decimal

from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.fi_analysis import analyze as analyze_fi
from sap_payroll_analysis.payroll_posting import analyze as analyze_payroll
from sap_payroll_analysis.payroll_wagetypes import probe_reference_wage_types
from sap_payroll_analysis.report import build_result, effective_payroll_total, reconcile
from sap_payroll_analysis.tests.fakes import FakeConnection

PPDHD_F = [("DOCNUM", "N", 10, ""), ("RUNID", "N", 10, ""), ("BUKRS", "C", 4, ""),
           ("BUDAT", "D", 8, ""), ("DOCTYP", "C", 1, ""), ("REVDOC", "N", 10, ""), ("XBLNR", "C", 16, "")]
PPDIT_F = [("DOCNUM", "N", 10, ""), ("DOCLIN", "N", 10, ""), ("BUKRS", "C", 4, ""),
           ("HKONT", "C", 10, ""), ("KTOSL", "C", 3, ""), ("WAERS", "C", 5, ""),
           ("NEG_POSTNG", "C", 1, ""), ("PERNR", "N", 8, ""), ("ITTYP", "C", 1, ""),
           ("WRBTR", "P", 15, ""), ("SGTXT", "C", 50, "")]
PPOIX_F = [("RUNID", "N", 10, ""), ("LGART", "C", 4, ""), ("KOMOK", "C", 4, ""),
           ("WAERS", "C", 5, ""), ("NEG_POSTNG", "C", 1, ""), ("BETRG", "P", 15, "")]
BSIS_F = [("BUKRS", "C", 4, ""), ("HKONT", "C", 10, ""), ("GJAHR", "N", 4, ""), ("BELNR", "C", 10, ""),
          ("BUZEI", "N", 3, ""), ("SHKZG", "C", 1, ""), ("WRBTR", "P", 15, ""), ("WAERS", "C", 5, ""),
          ("MONAT", "N", 2, ""), ("BUDAT", "D", 8, ""), ("BLART", "C", 2, "")]


def _h(doc, run, bukrs):
    return {"DOCNUM": doc, "RUNID": run, "BUKRS": bukrs, "BUDAT": "20260630",
            "DOCTYP": "1", "REVDOC": "0000000000", "XBLNR": "HRPAY00001"}


def _it(doc, hkont, wrbtr, bukrs="1010"):
    return {"DOCNUM": doc, "DOCLIN": "0000000001", "BUKRS": bukrs, "HKONT": hkont, "KTOSL": "HRF",
            "WAERS": "EUR", "NEG_POSTNG": "", "PERNR": "00000000", "ITTYP": "", "WRBTR": wrbtr, "SGTXT": ""}


def _conn() -> FakeConnection:
    return FakeConnection({
        "PEVST": {"fields": [("RUNID", "N", 10, "")], "rows": []},
        "PPDHD": {"fields": PPDHD_F, "rows": [_h("0000005392", "0000001298", "1010"),
                                             _h("0000005394", "0000001299", "1010")]},
        "PPDIX": {"fields": [("RUNID", "N", 10, ""), ("DOCNUM", "N", 10, "")], "rows": []},
        "PPDIT": {"fields": PPDIT_F, "rows": [
            _it("0000005392", "0023120000", "727258.35-"),
            _it("0000005392", "0063200100", "500000.00"),
            _it("0000005394", "0063200100", "123456.78"),
        ]},
        "PPOIX": {"fields": PPOIX_F, "rows": [
            {"RUNID": "0000001298", "LGART": "/559", "KOMOK": "S003", "WAERS": "EUR", "NEG_POSTNG": "", "BETRG": "720000.00-"},
            {"RUNID": "0000001298", "LGART": "/558", "KOMOK": "S003", "WAERS": "EUR", "NEG_POSTNG": "", "BETRG": "4046.64-"},
        ]},
        "BSIS": {"fields": BSIS_F, "rows": [
            {"BUKRS": "1010", "HKONT": "0023120000", "GJAHR": "2026", "BELNR": "0100000001", "BUZEI": "001",
             "SHKZG": "H", "WRBTR": "727258.35", "WAERS": "EUR", "MONAT": "06", "BUDAT": "20260630", "BLART": "HR"},
        ]},
        "BSAS": {"fields": BSIS_F, "rows": []},
        "BKPF": {"fields": [("BUKRS", "C", 4, ""), ("BELNR", "C", 10, "")], "rows": []},
    })


def test_payroll_maps_runs_to_docs_and_extracts_account():
    rep = analyze_payroll(_conn(), DEFAULTS)
    assert rep.resolved
    assert rep.resolved_fields["PPDHD.run"] == "RUNID"
    assert rep.resolved_fields["PPDIT.account"] == "HKONT"
    assert rep.doc_to_run == {"0000005392": "0000001298", "0000005394": "0000001299"}
    # só a linha 23120000 do doc 5392 conta
    assert len(rep.items) == 1
    assert rep.items[0].signed_amount == Decimal("-727258.35")
    assert rep.total == Decimal("-727258.35")
    assert rep.match_company == "1010"
    # todas as contas movimentadas ficam no resumo
    accts = {(r["company"], r["account"]) for r in rep.by_company_account}
    assert ("1010", "0023120000") in accts and ("1010", "0063200100") in accts


def test_payroll_warns_when_requested_company_absent():
    from dataclasses import replace

    params = replace(DEFAULTS, empresa="9999")  # empresa sem movimento
    rep = analyze_payroll(_conn(), params)
    assert any("9999" in w and "1010" in w for w in rep.warnings)


def test_fi_uses_bsis_with_period_filter_and_sign():
    rep = analyze_fi(_conn(), DEFAULTS)
    assert rep.resolved
    assert rep.source == "BSIS/BSAS"
    assert rep.total == Decimal("-727258.35")  # crédito (H) -> negativo
    assert rep.total_credit == Decimal("-727258.35")


def test_wage_type_probe_reference_total():
    wt = probe_reference_wage_types(_conn(), DEFAULTS)
    assert wt.resolved
    assert set(wt.by_wage_type) == {"/558", "/559"}
    assert wt.reference_total == Decimal("-724046.64")
    assert wt.symbolic_accounts == ["S003"]


def test_full_reconciliation():
    conn = _conn()
    payroll = analyze_payroll(conn, DEFAULTS)
    fi = analyze_fi(conn, DEFAULTS)
    wt = probe_reference_wage_types(conn, DEFAULTS)
    lines = reconcile(DEFAULTS, payroll, fi, wt)

    rh_fi = next(l for l in lines if l.label.startswith("Posting RH x FI"))
    assert rh_fi.left == Decimal("727258.35") and rh_fi.right == Decimal("727258.35")
    assert rh_fi.status == "OK"

    eff, comp = effective_payroll_total(payroll, DEFAULTS)
    assert eff == Decimal("727258.35")
    assert comp.startswith("1010")

    result = build_result(DEFAULTS, payroll, fi, lines, {"user": "X"}, wt)
    assert result.fi["source"] == "BSIS/BSAS"
    assert result.payroll_posting["reconciliation_company"].startswith("1010")
    assert result.payroll_posting["wage_types"]["reference_total_all_companies"] == "-724046.64"
    assert result.next_steps

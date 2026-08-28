"""Fase 2 (sem SAP): ligar rubricas PPOIX à linha de posting via PPDIX.

Cadeia testada:  PPOIX.TSLIN == PPDIX.LINUM ;
                 PPDIX.(DOCNUM,DOCLIN) == PPDIT.(DOCNUM,DOCLIN)
"""

from dataclasses import replace
from decimal import Decimal

from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.payroll_posting import analyze as analyze_payroll
from sap_payroll_analysis.payroll_wagetypes import (
    link_wage_types_to_posting_line,
    resolve_account_determination,
)
from sap_payroll_analysis.tests.fakes import FakeConnection

PARAMS = replace(DEFAULTS, empresa="1010", conta="23120000", primary_run="0000001298")

PPDHD_F = [("DOCNUM", "N", 10, ""), ("RUNID", "N", 10, ""), ("BUKRS", "C", 4, ""),
           ("BUDAT", "D", 8, ""), ("DOCTYP", "C", 1, ""), ("REVDOC", "N", 10, ""), ("XBLNR", "C", 16, "")]
PPDIT_F = [("DOCNUM", "N", 10, ""), ("DOCLIN", "N", 10, ""), ("BUKRS", "C", 4, ""),
           ("HKONT", "C", 10, ""), ("KTOSL", "C", 3, ""), ("WAERS", "C", 5, ""),
           ("NEG_POSTNG", "C", 1, ""), ("PERNR", "N", 8, ""), ("ITTYP", "C", 1, ""),
           ("WRBTR", "P", 15, ""), ("SGTXT", "C", 50, "")]
PPDIX_F = [("RUNID", "N", 10, ""), ("EVTYP", "C", 2, ""), ("LINUM", "N", 10, ""),
           ("DOCNUM", "N", 10, ""), ("DOCLIN", "N", 10, "")]
PPOIX_F = [("RUNID", "N", 10, ""), ("PERNR", "N", 8, ""), ("POSTNUM", "N", 5, ""),
           ("RTLINE", "N", 5, ""), ("TSLIN", "N", 10, ""), ("LGART", "C", 4, ""),
           ("KOMOK", "C", 4, ""), ("BETRG", "P", 15, ""), ("ACTSIGN", "C", 1, ""),
           ("NEG_POSTNG", "C", 1, "")]
T52EL_F = [("MOLGA", "C", 2, ""), ("LGART", "C", 4, ""), ("ENDDA", "D", 8, ""),
           ("SIGN", "C", 1, ""), ("SYMKO", "C", 4, ""), ("SPPRC", "C", 1, "")]
T52EK_F = [("SYMKO", "C", 4, ""), ("KOART", "C", 2, ""), ("U_MOMAG", "C", 1, ""), ("NEG_POSTNG", "C", 1, "")]
T030_F = [("KTOPL", "C", 4, ""), ("KTOSL", "C", 3, ""), ("BWMOD", "C", 4, ""),
          ("KOMOK", "C", 2, ""), ("BKLAS", "C", 4, ""), ("KONTS", "C", 10, ""), ("KONTH", "C", 10, "")]


def _poix(pernr, postnum, tslin, lgart, betrg):
    return {"RUNID": "0000001298", "PERNR": pernr, "POSTNUM": postnum, "RTLINE": "001",
            "TSLIN": tslin, "LGART": lgart, "KOMOK": "S003", "BETRG": betrg,
            "ACTSIGN": "A", "NEG_POSTNG": ""}


def _t030(bwmod, komok, acct):
    return {"KTOPL": "PCPT", "KTOSL": "HRF", "BWMOD": bwmod, "KOMOK": komok,
            "BKLAS": "", "KONTS": acct, "KONTH": acct}


def _conn() -> FakeConnection:
    return FakeConnection({
        "PEVST": {"fields": [("RUNID", "N", 10, "")], "rows": []},
        "PPDHD": {"fields": PPDHD_F, "rows": [
            {"DOCNUM": "0000005392", "RUNID": "0000001298", "BUKRS": "1010", "BUDAT": "20260630",
             "DOCTYP": "1", "REVDOC": "0000000000", "XBLNR": "HRPAY00001"},
        ]},
        "PPDIT": {"fields": PPDIT_F, "rows": [
            {"DOCNUM": "0000005392", "DOCLIN": "0000000326", "BUKRS": "1010", "HKONT": "0023120000",
             "KTOSL": "HRF", "WAERS": "EUR", "NEG_POSTNG": "", "PERNR": "00000000", "ITTYP": "",
             "WRBTR": "727258.35-", "SGTXT": ""},
            {"DOCNUM": "0000005392", "DOCLIN": "0000000001", "BUKRS": "1010", "HKONT": "0063200100",
             "KTOSL": "HRC", "WAERS": "EUR", "NEG_POSTNG": "", "PERNR": "00000000", "ITTYP": "",
             "WRBTR": "500000.00", "SGTXT": ""},
        ]},
        # LINUM 4 e 347 -> DOCLIN 326 (linha 23120000); LINUM 9 -> outra linha
        "PPDIX": {"fields": PPDIX_F, "rows": [
            {"RUNID": "0000001298", "EVTYP": "PP", "LINUM": "0000000004", "DOCNUM": "0000005392", "DOCLIN": "0000000326"},
            {"RUNID": "0000001298", "EVTYP": "PP", "LINUM": "0000000347", "DOCNUM": "0000005392", "DOCLIN": "0000000326"},
            {"RUNID": "0000001298", "EVTYP": "PP", "LINUM": "0000000009", "DOCNUM": "0000005392", "DOCLIN": "0000000001"},
        ]},
        "PPOIX": {"fields": PPOIX_F, "rows": [
            _poix("00000005", "00010", "0000000004", "/559", "724000.00-"),
            _poix("00000009", "00011", "0000000004", "/558", "13.97-"),
            _poix("00000197", "00013", "0000000004", "/561", "627.91"),
            _poix("00006637", "00014", "0000000004", "/563", "2587.53-"),
            _poix("80001145", "00005", "0000000347", "0029", "1090.00-"),
            # ruído: TSLIN de outra linha de posting -> não deve entrar
            _poix("00000005", "00002", "0000000009", "8000", "999999.00"),
        ]},
        "T52EL": {"fields": T52EL_F, "rows": [
            {"MOLGA": "19", "LGART": "/558", "ENDDA": "99991231", "SIGN": "-", "SYMKO": "S003", "SPPRC": ""},
            {"MOLGA": "19", "LGART": "/559", "ENDDA": "99991231", "SIGN": "-", "SYMKO": "S003", "SPPRC": ""},
            {"MOLGA": "19", "LGART": "/563", "ENDDA": "99991231", "SIGN": "-", "SYMKO": "S003", "SPPRC": ""},
            {"MOLGA": "19", "LGART": "/561", "ENDDA": "99991231", "SIGN": "+", "SYMKO": "S003", "SPPRC": ""},
            {"MOLGA": "19", "LGART": "0029", "ENDDA": "99991231", "SIGN": "+", "SYMKO": "S003", "SPPRC": ""},
        ]},
        "T52EK": {"fields": T52EK_F, "rows": [
            {"SYMKO": "S003", "KOART": "F", "U_MOMAG": "X", "NEG_POSTNG": ""},
        ]},
        "T030": {"fields": T030_F, "rows": [
            _t030("/558", "2", "0023120000"), _t030("/558", "1", "0023110000"),
            _t030("/559", "2", "0023120000"),
            _t030("/561", "2", "0023120000"),
            _t030("/563", "2", "0023120000"),
            _t030("0029", "2", "0023120000"),
            _t030("S003", "2", ""),
        ]},
    })


def test_link_builds_composition_and_residual():
    conn = _conn()
    payroll = analyze_payroll(conn, PARAMS)
    link = link_wage_types_to_posting_line(conn, PARAMS, payroll)

    assert link.resolved
    assert link.posting_doc_lines == [("0000005392", "0000000326")]
    assert link.posting_line_amount == Decimal("-727258.35")
    assert set(link.transfer_linums) == {"0000000004", "0000000347"}
    assert link.komok_set == ["S003"]

    # 5 linhas ligadas (a 6ª, TSLIN 9, é de outra linha de posting)
    assert link.ppoix_rows == 5
    assert set(link.by_wage_type) == {"/558", "/559", "/561", "/563", "0029"}
    assert link.reference_total == Decimal("-724013.97")     # /559 + /558
    # outras = /561 (+627.91) + /563 (-2587.53) + 0029 (-1090.00)
    assert link.other_total == Decimal("-3049.62")
    assert link.ppoix_total == Decimal("-727063.59")
    # resíduo = ppoix_total - posting_line_amount = -727063.59 - (-727258.35)
    assert link.residual_vs_posting == Decimal("194.76")


def test_link_ignores_ppoix_rows_of_other_posting_lines():
    conn = _conn()
    payroll = analyze_payroll(conn, PARAMS)
    link = link_wage_types_to_posting_line(conn, PARAMS, payroll)
    assert "8000" not in link.by_wage_type  # TSLIN 9 -> DOCLIN 1, não a conta alvo


def test_account_determination_confirms_target():
    conn = _conn()
    ad = resolve_account_determination(
        conn, PARAMS, symkos=["S003"], wage_types=["/558", "/559", "/561", "/563", "0029"], ktosl="HRF",
    )
    assert set(ad["wage_types_to_target"]) == {"/558", "/559", "/561", "/563", "0029"}
    assert any(r["KOART"] == "F" for r in ad["t52ek"])
    assert "Confirmado" in ad["conclusion"]


def test_link_sample_carries_docnum_doclin():
    conn = _conn()
    payroll = analyze_payroll(conn, PARAMS)
    link = link_wage_types_to_posting_line(conn, PARAMS, payroll)
    assert link.link_sample
    for row in link.link_sample:
        assert row["DOCNUM"] == "0000005392"
        assert row["DOCLIN"] == "0000000326"
        assert row["TSLIN"] in {"0000000004", "0000000347"}

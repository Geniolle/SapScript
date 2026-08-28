"""Fase 4.1 — rastreio PPOIX -> PPDIX -> PPDIT (sem SAP, via FakeConnection)."""

from dataclasses import replace
from decimal import Decimal

from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.wagetype_trace import (
    _signed,
    explain_amount_sign_path,
    trace_wagetype,
    write_trace_csv,
    write_trace_json,
)
from sap_payroll_analysis.wagetype_trace import PpditRow, PpoixRow
from sap_payroll_analysis.tests.fakes import FakeConnection

P = replace(DEFAULTS, empresa="1010", conta="23120000", primary_run="0000001298")

PPDHD_F = [("DOCNUM", "N", 10, ""), ("RUNID", "N", 10, "")]
PPDIT_F = [("DOCNUM", "N", 10, ""), ("DOCLIN", "N", 10, ""), ("BUKRS", "C", 4, ""),
           ("HKONT", "C", 10, ""), ("KTOSL", "C", 3, ""), ("WRBTR", "P", 15, ""),
           ("WAERS", "C", 5, ""), ("PERNR", "N", 8, ""), ("NEG_POSTNG", "C", 1, ""),
           ("ITTYP", "C", 1, ""), ("SGTXT", "C", 50, "")]
PPDIX_F = [("RUNID", "N", 10, ""), ("LINUM", "N", 10, ""), ("DOCNUM", "N", 10, ""), ("DOCLIN", "N", 10, "")]
PPOIX_F = [("RUNID", "N", 10, ""), ("PERNR", "N", 8, ""), ("SEQNO", "N", 5, ""),
           ("ACTSIGN", "C", 1, ""), ("POSTNUM", "N", 5, ""), ("SPPRC", "C", 1, ""),
           ("RTLINE", "N", 5, ""), ("KOART", "C", 2, ""), ("MOMAG", "C", 3, ""),
           ("KOMOK", "C", 4, ""), ("TSLIN", "N", 10, ""), ("LGART", "C", 4, ""),
           ("ANZHL", "P", 15, ""), ("BETRG", "P", 15, ""), ("WAERS", "C", 5, ""),
           ("SWRETROACT", "C", 1, ""), ("SWPER", "C", 1, ""), ("NEG_POSTNG", "C", 1, "")]
T52EL_F = [("MOLGA", "C", 2, ""), ("LGART", "C", 4, ""), ("ENDDA", "D", 8, ""),
           ("SIGN", "C", 1, ""), ("SYMKO", "C", 4, ""), ("SPPRC", "C", 1, "")]
T52EK_F = [("SYMKO", "C", 4, ""), ("KOART", "C", 2, ""), ("U_MOMAG", "C", 1, ""), ("NEG_POSTNG", "C", 1, "")]
T030_F = [("KTOPL", "C", 4, ""), ("KTOSL", "C", 3, ""), ("BWMOD", "C", 4, ""),
          ("KOMOK", "C", 2, ""), ("BKLAS", "C", 4, ""), ("KONTS", "C", 10, ""), ("KONTH", "C", 10, "")]


def _px(pernr, seqno, lgart, betrg, tslin, komok="S003", momag="2", postnum="001", rtline="001",
        actsign="A", neg=""):
    return {"RUNID": "0000001298", "PERNR": pernr, "SEQNO": seqno, "ACTSIGN": actsign,
            "POSTNUM": postnum, "SPPRC": "", "RTLINE": rtline, "KOART": "F", "MOMAG": momag,
            "KOMOK": komok, "TSLIN": tslin, "LGART": lgart, "ANZHL": "0.00", "BETRG": betrg,
            "WAERS": "EUR", "SWRETROACT": "", "SWPER": "", "NEG_POSTNG": neg}


def _base(extra_ppoix=None, ppdit_wrbtr="727258.35-", extra_ppdix=None):
    ppoix = [
        _px("00000005", "01199", "0029", "265.00-", "0000000004"),
        _px("00000005", "01199", "/559", "1382.40-", "0000000004"),
        _px("00000005", "01198", "/559", "823.70-", "0000000000"),  # não transferido
        _px("80001145", "00200", "0029", "825.00-", "0000000347", momag="3"),
    ] + (extra_ppoix or [])
    return {
        "PPDHD": {"fields": PPDHD_F, "rows": [{"DOCNUM": "0000005392", "RUNID": "0000001298"}]},
        "PPDIT": {"fields": PPDIT_F, "rows": [
            {"DOCNUM": "0000005392", "DOCLIN": "0000000326", "BUKRS": "1010", "HKONT": "0023120000",
             "KTOSL": "HRF", "WRBTR": ppdit_wrbtr, "WAERS": "EUR", "PERNR": "00000000",
             "NEG_POSTNG": "", "ITTYP": "", "SGTXT": ""},
            {"DOCNUM": "0000005392", "DOCLIN": "0000000001", "BUKRS": "1010", "HKONT": "0063200100",
             "KTOSL": "HRC", "WRBTR": "500000.00", "WAERS": "EUR", "PERNR": "00000000",
             "NEG_POSTNG": "", "ITTYP": "", "SGTXT": ""},
        ]},
        "PPDIX": {"fields": PPDIX_F, "rows": [
            {"RUNID": "0000001298", "LINUM": "0000000004", "DOCNUM": "0000005392", "DOCLIN": "0000000326"},
            {"RUNID": "0000001298", "LINUM": "0000000347", "DOCNUM": "0000005392", "DOCLIN": "0000000326"},
        ] + (extra_ppdix or [])},
        "PPOIX": {"fields": PPOIX_F, "rows": ppoix},
        "T52EL": {"fields": T52EL_F, "rows": [
            {"MOLGA": "19", "LGART": "0029", "ENDDA": "99991231", "SIGN": "+", "SYMKO": "S003", "SPPRC": ""},
            {"MOLGA": "19", "LGART": "/559", "ENDDA": "99991231", "SIGN": "-", "SYMKO": "S003", "SPPRC": ""},
        ]},
        "T52EK": {"fields": T52EK_F, "rows": [{"SYMKO": "S003", "KOART": "F", "U_MOMAG": "X", "NEG_POSTNG": ""}]},
        "T030": {"fields": T030_F, "rows": [
            {"KTOPL": "PCPT", "KTOSL": "HRF", "BWMOD": "0029", "KOMOK": "1", "BKLAS": "", "KONTS": "0023110000", "KONTH": "0023110000"},
            {"KTOPL": "PCPT", "KTOSL": "HRF", "BWMOD": "0029", "KOMOK": "2", "BKLAS": "", "KONTS": "0023120000", "KONTH": "0023120000"},
            {"KTOPL": "PCPT", "KTOSL": "HRF", "BWMOD": "/559", "KOMOK": "2", "BKLAS": "", "KONTS": "0023120000", "KONTH": "0023120000"},
        ]},
    }


# 5. sinal SAP "265.00-"
def test_signed_parses_trailing_minus():
    assert _signed("265.00-", "") == Decimal("-265.00")
    assert _signed("265.00", "X") == Decimal("-265.00")      # NEG_POSTNG inverte
    assert _signed("265.00-", "X") == Decimal("265.00")


# 1. um PPOIX -> um PPDIX -> um PPDIT
def test_single_chain_0029_reaches_target():
    tr = trace_wagetype(FakeConnection(_base()), P, pernr="00000005", lgart="0029")
    assert len(tr.ppoix) == 1
    assert tr.ppoix[0].tslin == "0000000004"
    assert tr.ppoix[0].ppdix_dest == [("0000005392", "0000000326")]
    assert tr.reaches_account is True
    assert tr.reaches_target_line is True
    assert tr.target_doc_lines == [("0000005392", "0000000326")]
    # determinação de contas
    assert "0029" in tr.account_determination["wage_types_to_target"]
    assert tr.conclusion["residual_class"] in {"UNEXPLAINED", "PARTIALLY_EXPLAINED", "EXPLAINED"}


# 2. múltiplos PPOIX no mesmo TSLIN (agregação da linha inteira, 2 LINUM)
def test_transfer_line_aggregation_covers_all_feeding_linums():
    tr = trace_wagetype(FakeConnection(_base()), P, pernr="00000005", lgart="0029")
    rec = tr.reconciliation
    assert set(rec["transfer_line_tslins"]) == {"0000000004", "0000000347"}
    # 0029 total na linha = -265 + -825
    assert tr.transfer_line_by_lgart["0029"]["sum"] == "-1090.00"
    # -265 -1382.40 -825 (o -823.70 não entra: TSLIN 0)
    assert rec["ppoix_sum"] == "-2472.40"
    assert rec["ppdit_wrbtr"] == "-727258.35"


# 3. 0029 e /559 no mesmo destino
def test_0029_and_559_same_posting_line():
    tr = trace_wagetype(FakeConnection(_base()), P, pernr="00000005", lgart="0029",
                        compare_lgart="/559")
    assert tr.compare is not None and tr.compare.lgart == "/559"
    assert tr.compare.reaches_target_line is True
    assert tr.conclusion["same_posting_line_as_compare"] is True
    # a linha /559 SEQNO 01198 não é transferida
    non = [r for r in tr.compare.ppoix if not r.ppdix_dest]
    assert non and non[0].tslin == "0000000000"


# 4. 0029 e /559 em destinos diferentes
def test_0029_and_559_different_lines():
    tables = _base()
    tables["PPDIX"]["rows"].append(
        {"RUNID": "0000001298", "LINUM": "0000000009", "DOCNUM": "0000005392", "DOCLIN": "0000000900"})
    tables["PPDIT"]["rows"].append(
        {"DOCNUM": "0000005392", "DOCLIN": "0000000900", "BUKRS": "1010", "HKONT": "0023125000",
         "KTOSL": "HRF", "WRBTR": "1000.00-", "WAERS": "EUR", "PERNR": "00000000",
         "NEG_POSTNG": "", "ITTYP": "", "SGTXT": ""})
    # muda o /559 01199 para outro TSLIN/LINUM
    for r in tables["PPOIX"]["rows"]:
        if r["LGART"] == "/559" and r["SEQNO"] == "01199":
            r["TSLIN"] = "0000000009"
    tr = trace_wagetype(FakeConnection(tables), P, pernr="00000005", lgart="0029", compare_lgart="/559")
    assert tr.reaches_target_line is True          # 0029 -> 326
    assert tr.compare.reaches_target_line is False  # /559 -> 900 (conta 23125000)
    assert tr.conclusion["same_posting_line_as_compare"] is False


# 6. residual 265,65  +  8. ausência de evidência -> UNEXPLAINED
#    9. nunca classificar automaticamente valor pequeno como arredondamento
def test_residual_unexplained_no_rounding_claim():
    # soma PPOIX linha (LINUM 4+347) = -265 -1382.40 -825 = -2472.40
    # para delta -265.65 -> PPDIT = -2472.40 + 265.65 = -2206.75
    tr = trace_wagetype(FakeConnection(_base(ppdit_wrbtr="2206.75-")), P,
                        pernr="00000005", lgart="0029")
    rec, inv = tr.reconciliation, tr.residual_investigation
    assert rec["delta"] == "-265.65"
    assert rec["traced_row_in_line"] == "-265.00"
    assert rec["leftover_if_traced_row_excluded"] == "-0.65"
    assert inv["single_row_equals_delta"] == []
    assert inv["single_row_equals_leftover"] == []
    assert inv["rows_with_sub_cent"] == []
    assert inv["classification"] == "UNEXPLAINED"
    assert "arredond" not in inv["explanation"].lower() or "possível" not in inv["explanation"].lower()
    assert "Hipótese (NÃO provada" in inv["explanation"]


# 7. procura de combinação 0,65 — encontrada mas marcada como coincidência
def test_arithmetic_combo_flagged_as_coincidence():
    # soma linha passa a -2472.40 -0.65 = -2473.05 ; PPDIT p/ delta -265.65 = -2207.40
    extra = [_px("00099999", "00001", "/561", "0.65-", "0000000004")]
    tr = trace_wagetype(FakeConnection(_base(extra_ppoix=extra, ppdit_wrbtr="2207.40-")), P,
                        pernr="00000005", lgart="0029")
    inv = tr.residual_investigation
    combos = (inv["arithmetic_combos_for_leftover"] + inv["arithmetic_combos_for_delta"]
              + inv["single_row_equals_leftover"])
    assert combos  # existe a linha -0,65
    assert "coincid" in inv["arithmetic_combos_note"].lower()
    assert "não" in inv["arithmetic_combos_note"].lower() and "prova" in inv["arithmetic_combos_note"].lower()


def test_explain_amount_sign_path_liability_credit():
    px = PpoixRow(lgart="0029", betrg_raw="265.00-", betrg=Decimal("-265.00"), actsign="A", neg_postng="")
    pd = PpditRow(hkont="0023120000", wrbtr=Decimal("-727258.35"), wrbtr_raw="727258.35-")
    sp = explain_amount_sign_path(px, pd)
    assert "crédito" in sp.accounting_effect and "23120000" in sp.accounting_effect
    sp2 = explain_amount_sign_path(px, None)
    assert "não transferido" in sp2.accounting_effect


def test_outputs_written(tmp_path):
    tr = trace_wagetype(FakeConnection(_base()), P, pernr="00000005", lgart="0029")
    j = write_trace_json(tr, tmp_path / "t.json")
    c = write_trace_csv(tr, tmp_path / "t.csv")
    assert j.exists() and c.exists()
    csv_txt = c.read_text(encoding="utf-8-sig")
    assert "PPOIX" in csv_txt and "PPDIT" in csv_txt and "TSLIN_BY_LGART" in csv_txt
    import json
    d = json.loads(j.read_text(encoding="utf-8"))
    assert d["input"]["lgart"] == "0029"
    assert d["reaches_target_line"] is True

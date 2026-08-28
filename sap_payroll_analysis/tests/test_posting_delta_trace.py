"""Fase 4.2 — testes de `posting_delta_trace` (sem SAP, via FakeConnection)."""

from __future__ import annotations

import json
from decimal import Decimal

import pytest

from sap_payroll_analysis import security
from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.posting_delta_trace import (
    DELTA_ORIGIN_CLASSES,
    analyze_momag,
    analyze_pernr_breakdown,
    analyze_zero_tslin,
    build_seqno_usage_map,
    classify_delta_origin,
    discover_posting_intermediates,
    find_delta_candidates,
    find_first_divergence,
    find_previous_runs,
    find_transferred_nontransferred_pairs,
    inspect_ppopx,
    reconcile_target_line,
    trace_posting_delta,
    write_posting_delta_csvs,
    write_posting_delta_json,
)
from sap_payroll_analysis.tests.fakes import FakeConnection

RUN = "0000009999"
DOC = "0000000042"
DLN = "0000000100"

# --- DDIC (DD02L / DD03L) --------------------------------------------------

_DD02L_ROWS = [
    {"TABNAME": t, "TABCLASS": "TRANSP"}
    for t in ("PPOIX", "PPDIX", "PPOPX", "PPDIT", "PPDST", "PPDHD")
] + [{"TABNAME": "PPZZZ_UNRELATED", "TABCLASS": "VIEW"}]

_DD03L_ROWS = (
    [{"TABNAME": "PPOPX", "FIELDNAME": n, "POSITION": str(i + 1), "LENG": "8",
      "DATATYPE": "NUMC" if n not in ("MANDT",) else "CLNT", "ROLLNAME": ""}
     for i, n in enumerate(["MANDT", "PERNR", "SEQNO", "RUNID", "POSTNUM", "TSLIN", "ACTSIGN"])]
    + [{"TABNAME": "PPOIX", "FIELDNAME": n, "POSITION": str(i + 1), "LENG": "10",
        "DATATYPE": "CURR" if n == "BETRG" else "CHAR", "ROLLNAME": ""}
       for i, n in enumerate(["RUNID", "PERNR", "SEQNO", "ACTSIGN", "POSTNUM", "SPPRC", "RTLINE",
                              "KOART", "MOMAG", "KOMOK", "TSLIN", "LGART", "BETRG",
                              "SWRETROACT", "SWPER", "NEG_POSTNG"])]
    + [{"TABNAME": "PPDIX", "FIELDNAME": n, "POSITION": str(i + 1), "LENG": "10",
        "DATATYPE": "CHAR", "ROLLNAME": ""}
       for i, n in enumerate(["CLIENT", "EVTYP", "RUNID", "LINUM", "DOCNUM", "DOCLIN"])]
    + [{"TABNAME": "PPDIT", "FIELDNAME": n, "POSITION": str(i + 1), "LENG": "10",
        "DATATYPE": "CURR" if n == "WRBTR" else "CHAR", "ROLLNAME": ""}
       for i, n in enumerate(["DOCNUM", "DOCLIN", "BUKRS", "HKONT", "KTOSL", "WRBTR", "WAERS",
                              "NEG_POSTNG", "ITTYP"])]
    + [{"TABNAME": "PPDST", "FIELDNAME": n, "POSITION": str(i + 1), "LENG": "10",
        "DATATYPE": "CURR" if n == "WRBTR" else "CHAR", "ROLLNAME": ""}
       for i, n in enumerate(["DOCNUM", "DOCLIN", "SEQNO", "WRBTR", "WAERS", "CURTP"])]
)


def _fld_meta(names):
    return [(n, "CHAR", 20, n) for n in names]


def _table(rows):
    if not rows:
        return {"fields": _fld_meta(["DUMMY"]), "rows": []}
    cols = list(rows[0].keys())
    return {"fields": _fld_meta(cols), "rows": rows}


def _ppoix(runid, pernr, seqno, lgart, betrg, tslin, postnum="00001", momag="2",
           komok="S003", rtline="00001"):
    return {"RUNID": runid, "PERNR": pernr, "SEQNO": seqno, "ACTSIGN": "A",
            "POSTNUM": postnum, "SPPRC": "", "RTLINE": rtline, "KOART": "F",
            "MOMAG": momag, "KOMOK": komok, "TSLIN": tslin, "LGART": lgart,
            "BETRG": betrg, "SWRETROACT": "", "SWPER": "", "NEG_POSTNG": ""}


def base_tables(*, extra_ppoix=None, extra_ppdit=None, extra_ppdhd=None, extra_ppdix=None,
                ppdst_rows=None, ppopx_rows=None):
    # linha alvo alimentada por dois LINUMs: 4 e 9
    ppdix = [
        {"CLIENT": "100", "EVTYP": "PP", "RUNID": RUN, "LINUM": "0000000004",
         "DOCNUM": DOC, "DOCLIN": DLN},
        {"CLIENT": "100", "EVTYP": "PP", "RUNID": RUN, "LINUM": "0000000009",
         "DOCNUM": DOC, "DOCLIN": DLN},
        # LINUM 7 alimenta OUTRA linha (não alvo)
        {"CLIENT": "100", "EVTYP": "PP", "RUNID": RUN, "LINUM": "0000000007",
         "DOCNUM": DOC, "DOCLIN": "0000000200"},
    ] + (extra_ppdix or [])

    ppoix = [
        _ppoix(RUN, "00000001", "01", "/559", "100.00-", "0000000004", pernr_pn := "00001"),
        _ppoix(RUN, "00000001", "01", "/559", "100.00-", "0000000004"),
        _ppoix(RUN, "00000002", "01", "/559", "100.00-", "0000000004"),
        _ppoix(RUN, "00000003", "01", "/559", "50.00-", "0000000009", momag="3"),
        # TSLIN 7 (linha não alvo)
        _ppoix(RUN, "00000004", "01", "0042", "20.00", "0000000007"),
        # TSLIN 0 — pares que se anulam + resíduo
        _ppoix(RUN, "00000001", "01", "8000", "10.00", "0000000000"),
        _ppoix(RUN, "00000001", "01", "8000", "10.00-", "0000000000"),
        _ppoix(RUN, "00000001", "01", "/559", "5.00-", "0000000000"),
        _ppoix(RUN, "00000002", "02", "/559", "3.00-", "0000000000"),
    ] + (extra_ppoix or [])

    ppdit = [
        {"DOCNUM": DOC, "DOCLIN": DLN, "BUKRS": "1010", "HKONT": "0023120000", "KTOSL": "HRF",
         "WRBTR": "347.00-", "WAERS": "EUR", "NEG_POSTNG": "", "ITTYP": ""},
        {"DOCNUM": DOC, "DOCLIN": "0000000200", "BUKRS": "1010", "HKONT": "0063200100",
         "KTOSL": "HRC", "WRBTR": "20.00", "WAERS": "EUR", "NEG_POSTNG": "", "ITTYP": ""},
    ] + (extra_ppdit or [])

    ppdhd = [
        {"RUNID": RUN, "DOCNUM": DOC, "BUKRS": "1010", "BUDAT": "20260630", "BLDAT": "20260630",
         "BLART": "HR", "XBLNR": "HRPAY00001"},
    ] + (extra_ppdhd or [])

    ppopx = ppopx_rows if ppopx_rows is not None else [
        {"MANDT": "100", "PERNR": "00000050", "SEQNO": "99", "RUNID": RUN, "POSTNUM": "00009",
         "TSLIN": "0000000000", "ACTSIGN": "P"},
    ]

    return {
        "DD02L": _table(_DD02L_ROWS),
        "DD03L": _table(_DD03L_ROWS),
        "DD04T": _table([]),
        "PPDIX": _table(ppdix),
        "PPOIX": _table(ppoix),
        "PPDIT": _table(ppdit),
        "PPDHD": _table(ppdhd),
        "PPOPX": _table(ppopx),
        "PPDST": _table(ppdst_rows or []),
    }


# ---------------------------------------------------------------------------

def test_trace_feeders_sums_and_delta():
    conn = FakeConnection(base_tables())
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    assert tr.feeder_linums == ["0000000004", "0000000009"]
    assert tr.ppoix_rows == 4
    assert tr.ppoix_sum == Decimal("-350.00")
    assert tr.ppdit_wrbtr == Decimal("-347.00")
    assert tr.delta == Decimal("-3.00")


def test_checkpoints_have_no_amount_for_ppopx_and_empty_ppdst():
    conn = FakeConnection(base_tables())
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    by_stage = {c.stage: c for c in tr.checkpoints}
    assert by_stage["1-PPOIX"].amount == Decimal("-350.00")
    assert by_stage["2-PPOPX"].amount is None
    assert by_stage["3-PPDST"].amount is None and by_stage["3-PPDST"].row_count == 0
    assert by_stage["4-PPDIT"].amount == Decimal("-347.00")


def test_first_divergence_is_between_ppoix_and_ppdit():
    conn = FakeConnection(base_tables())
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    fd = tr.first_divergence
    assert fd["result"] == "DIVERGES"
    assert fd["between"] == "1-PPOIX -> 4-PPDIT"
    assert fd["delta"] == "3.00"
    assert "2-PPOPX" in fd["skipped_stages_without_value"]


def test_discover_intermediates_money_tables():
    conn = FakeConnection(base_tables())
    res = discover_posting_intermediates(conn, DEFAULTS)
    assert set(res["money_bearing_tables"]) == {"PPOIX", "PPDIT", "PPDST"}
    assert res["tables"]["PPOPX"]["money_fields"] == []
    assert res["tables"]["PPDIX"]["money_fields"] == []


def test_inspect_ppopx_no_money_no_overlap():
    conn = FakeConnection(base_tables())
    line_rows = [_ppoix(RUN, "00000001", "01", "/559", "100.00-", "0000000004")]
    res = inspect_ppopx(conn, DEFAULTS, RUN, target_rows=line_rows)
    assert res["has_money_field"] is False
    assert res["rows_for_run"] == 1
    assert all(v["matched_rows"] == 0 for v in res["overlap_with_target_line"].values())
    assert res["conclusion"].startswith("[PROVED]")


def test_reconcile_target_line_per_linum():
    conn = FakeConnection(base_tables())
    res = reconcile_target_line(conn, DEFAULTS, RUN, DOC, DLN)
    assert res["feeder_linums"] == ["0000000004", "0000000009"]
    per = {p["linum"]: p for p in res["per_linum"]}
    assert per["0000000004"]["sum"] == "-300.00" and per["0000000004"]["rows"] == 3
    assert per["0000000009"]["sum"] == "-50.00"
    assert res["delta"] == "-3.00"
    assert res["no_amount_bearing_intermediate"] is True


def test_analyze_zero_tslin_totals_and_residual():
    conn = FakeConnection(base_tables())
    res = analyze_zero_tslin(conn, DEFAULTS, RUN)
    assert res["rows"] == 4
    assert res["sum"] == "-8.00"          # 10 -10 -5 -3
    assert "8000" in res["lgart_net_zero"]
    assert "/559" in res["lgart_with_residual"]
    assert res["lgart_with_residual"]["/559"]["sum"] == "-8.00"
    assert "sum_by_pernr" in res and "sum_by_postnum" in res


def test_transferred_nontransferred_pairs():
    conn = FakeConnection(base_tables())
    px = [r for r in base_tables()["PPOIX"]["rows"]]
    pairs = find_transferred_nontransferred_pairs(
        conn, DEFAULTS, RUN, run_ppoix=px, feeder_linums=["0000000004", "0000000009"])
    p1 = [p for p in pairs if p["pernr"] == "00000001" and p["lgart"] == "/559"]
    assert p1 and p1[0]["transfer_sum"] == "-200.00" and p1[0]["zero_sum"] == "-5.00"


def test_seqno_usage_map_first_seen_vs_reused():
    tables = base_tables()
    conn = FakeConnection(tables)
    px = tables["PPOIX"]["rows"]
    usage = build_seqno_usage_map(conn, DEFAULTS, RUN, run_ppoix=px, other_runs={RUN})
    # sem outros runs comparáveis -> tudo FIRST_SEEN
    assert usage["classification_counts"].get("FIRST_SEEN", 0) >= 1
    assert "REUSED" not in usage["classification_counts"]


def test_find_previous_runs_none_for_company():
    conn = FakeConnection(base_tables())
    res = find_previous_runs(conn, DEFAULTS, RUN)
    assert res["prior_runs_same_company"] == []
    assert "não é aplicável" in res["conclusion"]


def test_find_previous_runs_detects_prior():
    extra = [{"RUNID": "0000009000", "DOCNUM": "0000000041", "BUKRS": "1010",
              "BUDAT": "20260531", "BLDAT": "20260531", "BLART": "HR", "XBLNR": "X"}]
    conn = FakeConnection(base_tables(extra_ppdhd=extra))
    res = find_previous_runs(conn, DEFAULTS, RUN)
    assert res["prior_runs_same_company"] == ["0000009000"]


def test_classification_between_ppoix_and_ppdit():
    conn = FakeConnection(base_tables())
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    assert tr.classification["classification"] == "PROVED_BETWEEN_PPOIX_AND_PPDIT"
    assert tr.classification["classification"] in DELTA_ORIGIN_CLASSES


def test_classification_previous_run_netting():
    # run anterior 9000 (empresa 1010) lança exactamente -3,00 na conta alvo
    extra_hd = [{"RUNID": "0000009000", "DOCNUM": "0000000041", "BUKRS": "1010",
                 "BUDAT": "20260531", "BLDAT": "20260531", "BLART": "HR", "XBLNR": "X"}]
    extra_it = [{"DOCNUM": "0000000041", "DOCLIN": "0000000010", "BUKRS": "1010",
                 "HKONT": "0023120000", "KTOSL": "HRF", "WRBTR": "3.00-", "WAERS": "EUR",
                 "NEG_POSTNG": "", "ITTYP": ""}]
    conn = FakeConnection(base_tables(extra_ppdhd=extra_hd, extra_ppdit=extra_it))
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    assert tr.classification["classification"] == "PROVED_PREVIOUS_RUN_NETTING"


def test_classification_ppopx_netting():
    # PPOPX contém a chave (PERNR+SEQNO+POSTNUM+TSLIN) de uma linha alvo cujo
    # valor é exactamente o delta (-3,00): adicionamos uma linha alvo -3,00 e a
    # sua correspondência em PPOPX, e ajustamos a PPDIT para manter o delta.
    extra_px = [_ppoix(RUN, "00000009", "07", "/559", "3.00-", "0000000004",
                       postnum="00077")]
    ppopx = [{"MANDT": "100", "PERNR": "00000009", "SEQNO": "07", "RUNID": RUN,
              "POSTNUM": "00077", "TSLIN": "0000000004", "ACTSIGN": "P"}]
    # nova soma PPOIX linha alvo = -353,00 ; mantemos PPDIT -347,00 => delta -6,00?
    # para o teste do PPOPX-netting queremos delta == soma das linhas com match.
    # match sum = -3,00 => precisamos delta -3,00 => PPDIT = -350,00
    extra_it = []
    tables = base_tables(extra_ppoix=extra_px, ppopx_rows=ppopx)
    tables["PPDIT"]["rows"][0]["WRBTR"] = "350.00-"
    conn = FakeConnection(tables)
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    assert tr.delta == Decimal("-3.00")
    assert tr.classification["classification"] == "PROVED_AT_PPOPX"


def test_find_delta_candidates_single_row_is_candidate_not_proved():
    line_rows = [
        _ppoix(RUN, "00000001", "01", "/559", "100.00-", "0000000004"),
        _ppoix(RUN, "00000005", "01", "0029", "3.00-", "0000000004"),
    ]
    res = find_delta_candidates(line_rows, Decimal("-3.00"))
    kinds = [c["kind"] for c in res["candidates"]]
    assert "single_ppoix_row_equals_delta" in kinds
    assert all("[CANDIDATE]" in c["status"] for c in res["candidates"]
               if c["kind"] == "single_ppoix_row_equals_delta")


def test_analyze_momag_semantics_unknown():
    line_rows = [
        _ppoix(RUN, "00000001", "01", "/559", "100.00-", "0000000004", momag="2"),
        _ppoix(RUN, "00000003", "01", "/559", "50.00-", "0000000009", momag="3"),
    ]
    res = analyze_momag(None, DEFAULTS, line_rows)
    assert res["semantics"] == "UNKNOWN"
    assert set(res["by_momag"]) == {"2", "3"}


def test_pernr_breakdown_flags_multi_seqno():
    line_rows = [
        _ppoix(RUN, "00000007", "10", "/559", "100.00-", "0000000004"),
        _ppoix(RUN, "00000007", "11", "/559", "20.00-", "0000000004"),
        _ppoix(RUN, "00000008", "10", "/559", "30.00-", "0000000004"),
    ]
    res = analyze_pernr_breakdown(line_rows)
    assert res["n_pernr"] == 2
    multi = {r["pernr"] for r in res["multi_seqno_pernr"]}
    assert multi == {"00000007"}


def test_find_first_divergence_no_divergence():
    from sap_payroll_analysis.posting_delta_trace import PostingCheckpoint
    cps = [
        PostingCheckpoint(stage="1-PPOIX", source="PPOIX", amount=Decimal("-10.00")),
        PostingCheckpoint(stage="2-PPOPX", source="PPOPX", amount=None),
        PostingCheckpoint(stage="4-PPDIT", source="PPDIT", amount=Decimal("-10.00")),
    ]
    assert find_first_divergence(cps)["result"] == "NO_DIVERGENCE"


def test_write_json_and_csvs(tmp_path):
    conn = FakeConnection(base_tables())
    tr = trace_posting_delta(conn, DEFAULTS, docnum=DOC, doclin=DLN, run=RUN)
    jp = write_posting_delta_json(tr, tmp_path / "pd.json")
    data = json.loads(jp.read_text(encoding="utf-8"))
    assert data["delta"] == "-3.00"
    assert data["classification"]["classification"] in DELTA_ORIGIN_CLASSES
    csvs = write_posting_delta_csvs(tr, tmp_path, RUN, DOC, DLN)
    assert len(csvs) == 4 and all(p.exists() for p in csvs)


def test_ppdst_and_ppdsh_whitelisted():
    assert "PPDST" in security.READ_ONLY_TABLE_WHITELIST
    assert "PPDSH" in security.READ_ONLY_TABLE_WHITELIST

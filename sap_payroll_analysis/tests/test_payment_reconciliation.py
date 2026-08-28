"""Fase 5.0 — testes de `payment_reconciliation` (sem SAP, via FakeConnection)."""

from __future__ import annotations

import json
from decimal import Decimal

import pytest

from sap_payroll_analysis import security
from sap_payroll_analysis.config import DEFAULTS
from sap_payroll_analysis.payment_reconciliation import (
    aggregate_regu_by_employee,
    build_payroll_payment_expectations,
    classify_reconciliation,
    discover_payment_runs,
    inspect_regu_schema,
    match_payroll_to_regu,
    reconcile_payroll_payments,
    resolve_employee_identity,
    select_payment_run,
)
from sap_payroll_analysis.sap_connection import SapConnectionError, require_prefix
from sap_payroll_analysis.tests.fakes import FakeConnection

RUN = "0000001298"
COMPANY = "1010"
PERIOD = "202606"

_PPOIX_ORDER = ["RUNID", "PERNR", "SEQNO", "ACTSIGN", "POSTNUM", "SPPRC", "RTLINE",
                "KOART", "MOMAG", "KOMOK", "TSLIN", "LGART", "BETRG",
                "SWRETROACT", "SWPER", "NEG_POSTNG"]

_REGUH_FIELDS = ["ZBUKR", "LAUFD", "LAUFI", "LIFNR", "KUNNR", "EMPFG", "PERNR", "VBLNR",
                 "BELNR", "GJAHR", "WAERS", "RWBTR", "RZAWE", "ZNME1", "ZBNKN", "XBLNR", "ZALDT"]
_REGUP_FIELDS = ["ZBUKR", "LAUFD", "LAUFI", "LIFNR", "KUNNR", "EMPFG", "BELNR", "GJAHR",
                 "BUZEI", "WRBTR", "DMBTR", "WAERS", "SGTXT"]
_REGUV_FIELDS = ["LAUFD", "LAUFI", "ZBUKR", "XECHT", "WAERS"]


def _dd03l(mapping: dict[str, list[str]]) -> list[dict[str, str]]:
    rows = []
    for tab, flds in mapping.items():
        for i, n in enumerate(flds):
            rows.append({"TABNAME": tab, "FIELDNAME": n, "POSITION": str(i + 1),
                         "LENG": "10", "DATATYPE": "CURR" if n in ("RWBTR", "WRBTR", "DMBTR") else "CHAR",
                         "ROLLNAME": ""})
    return rows


def _fld_meta(names):
    return [(n, "CHAR", 20, n) for n in names]


def _table(rows):
    if not rows:
        return {"fields": _fld_meta(["DUMMY"]), "rows": []}
    return {"fields": _fld_meta(list(rows[0].keys())), "rows": rows}


def _ppoix(pernr, seqno, lgart, betrg, tslin, postnum="00001"):
    d = dict.fromkeys(_PPOIX_ORDER, "")
    d.update(RUNID=RUN, PERNR=pernr, SEQNO=seqno, ACTSIGN="A", POSTNUM=postnum,
             KOART="F", MOMAG="2", KOMOK="S003", TSLIN=tslin, LGART=lgart, BETRG=betrg)
    return d


def _reguh(laufd, laufi, empfg="", pernr="", lifnr="", rwbtr="0.00", vblnr="",
           zbukr=COMPANY, waers="EUR", name="", zaldt=""):
    d = dict.fromkeys(_REGUH_FIELDS, "")
    d.update(ZBUKR=zbukr, LAUFD=laufd, LAUFI=laufi, EMPFG=empfg, PERNR=pernr, LIFNR=lifnr,
             RWBTR=rwbtr, VBLNR=vblnr, WAERS=waers, RZAWE="U", ZNME1=name,
             ZALDT=zaldt or laufd, GJAHR="2026")
    return d


def base_tables(*, ppoix=None, reguh=None, reguv=None, reguh_has_pernr=True):
    reguh_fields = _REGUH_FIELDS if reguh_has_pernr else [f for f in _REGUH_FIELDS if f != "PERNR"]
    dd = _dd03l({"REGUH": reguh_fields, "REGUP": _REGUP_FIELDS, "REGUV": _REGUV_FIELDS})
    return {
        "DD03L": _table(dd),
        "DD04T": _table([]),
        "PPOIX": _table(ppoix if ppoix is not None else _default_ppoix()),
        "REGUH": _table(reguh if reguh is not None else []),
        "REGUP": _table([]),
        "REGUV": _table(reguv if reguv is not None else []),
        "REGUHM": _table([]),
        "REGUT": _table([]),
    }


def _default_ppoix():
    # 3 colaboradores, /559 corrente transferido (TSLIN!=0)
    return [
        _ppoix("00000005", "01199", "/559", "1382.40-", "0000000004"),
        _ppoix("00000005", "01199", "0029", "265.00-", "0000000004", "00005"),
        _ppoix("00000006", "01200", "/559", "2000.00-", "0000000004"),
        _ppoix("00000007", "01201", "/559", "1500.00-", "0000000004"),
        # ruído: /559 TSLIN=0 do 5 (versão anterior) — NÃO deve entrar no esperado
        _ppoix("00000005", "01198", "/559", "823.70-", "0000000000"),
    ]


# ---------------------------------------------------------------------------

def test_regu_tables_whitelisted():
    for t in ("REGUH", "REGUP", "REGUV", "REGUHM", "REGUT"):
        assert t in security.READ_ONLY_TABLE_WHITELIST


def test_require_prefix_aborts_without_r3(monkeypatch):
    for s in ("USER", "PASSWD", "ASHOST", "SYSNR", "CLIENT"):
        monkeypatch.delenv(f"SAP_R3_{s}", raising=False)
    with pytest.raises(SapConnectionError) as ei:
        require_prefix("SAP_R3_", purpose="payroll/payment reconciliation")
    assert "SAP_R3_* connection parameters required for payroll/payment reconciliation" in str(ei.value)


def test_require_prefix_ok_when_complete(monkeypatch):
    for s, v in (("USER", "U"), ("PASSWD", "P"), ("ASHOST", "1.2.3.4"),
                 ("SYSNR", "00"), ("CLIENT", "100")):
        monkeypatch.setenv(f"SAP_R3_{s}", v)
    assert require_prefix("SAP_R3_", purpose="x") == "SAP_R3_"


def test_inspect_regu_schema_roles():
    conn = FakeConnection(base_tables())
    sch = inspect_regu_schema(conn, DEFAULTS)
    reguh = sch["tables"]["REGUH"]
    assert reguh["exists"] and reguh["amount_field"] == "RWBTR"
    assert reguh["roles"]["company"] == "ZBUKR"
    assert reguh["roles"]["run_date"] == "LAUFD"
    assert reguh["roles"]["run_id"] == "LAUFI"
    assert reguh["has_pernr"] is True


def test_payment_run_discovery():
    reguh = [
        _reguh("20260628", "PAY001", empfg="00000005", rwbtr="1382.40"),
        _reguh("20260628", "PAY001", empfg="00000006", rwbtr="2000.00"),
        _reguh("20260628", "PAY001", empfg="00000007", rwbtr="1500.00"),
        _reguh("20260515", "OLD999", empfg="00000005", rwbtr="900.00"),   # fora da janela
    ]
    reguv = [{"LAUFD": "20260628", "LAUFI": "PAY001", "ZBUKR": COMPANY, "XECHT": "X", "WAERS": "EUR"}]
    conn = FakeConnection(base_tables(reguh=reguh, reguv=reguv))
    sch = inspect_regu_schema(conn, DEFAULTS)
    cands = discover_payment_runs(conn, DEFAULTS, company=COMPANY, period=PERIOD, schema=sch)
    keys = {(c.laufd, c.laufi) for c in cands}
    assert ("20260628", "PAY001") in keys
    assert ("20260515", "OLD999") not in keys       # fora da janela junho→15 julho
    c = next(c for c in cands if c.laufi == "PAY001")
    assert c.payment_count == 3 and c.total == Decimal("4882.40") and c.is_real == "X"


def test_multiple_candidates_ranked():
    reguh = (
        [_reguh("20260628", "PAY001", empfg=f"0000000{n}", rwbtr="100.00") for n in range(1, 4)]
        + [_reguh("20260702", "PAY002", empfg=f"0000000{n}", rwbtr="100.00") for n in range(1, 9)]
    )
    conn = FakeConnection(base_tables(reguh=reguh))
    sch = inspect_regu_schema(conn, DEFAULTS)
    cands = discover_payment_runs(conn, DEFAULTS, company=COMPANY, period=PERIOD, schema=sch)
    selected, ranked = select_payment_run(cands, period=PERIOD, payroll_employee_count=3)
    assert len(ranked) == 2
    assert selected["laufi"] == "PAY001"           # nº beneficiários ~ nº colaboradores
    assert ranked[0].confidence in ("HIGH_CONFIDENCE", "MEDIUM_CONFIDENCE")


def test_select_not_only_by_total():
    # PAY_T tem total mais perto da referência mas nº beneficiários errado;
    # PAY_G tem nº beneficiários certo. Deve ganhar PAY_G.
    reguh = (
        [_reguh("20260628", "PAY_G", empfg=f"0000000{n}", rwbtr="10.00") for n in (5, 6, 7)]
        + [_reguh("20260628", "PAY_T", empfg="00000009", rwbtr="4882.40")]
    )
    conn = FakeConnection(base_tables(reguh=reguh))
    sch = inspect_regu_schema(conn, DEFAULTS)
    cands = discover_payment_runs(conn, DEFAULTS, company=COMPANY, period=PERIOD, schema=sch)
    selected, _ = select_payment_run(
        cands, period=PERIOD, payroll_employee_count=3,
        payroll_reference_total=Decimal("4882.40"))
    assert selected["laufi"] == "PAY_G"


def test_identity_pernr_direct():
    reguh = [_reguh("20260628", "PAY001", pernr="00000005", empfg="", rwbtr="1382.40"),
             _reguh("20260628", "PAY001", pernr="00000006", rwbtr="2000.00")]
    conn = FakeConnection(base_tables(reguh=reguh))
    sch = inspect_regu_schema(conn, DEFAULTS)
    roles = sch["tables"]["REGUH"]["roles"]
    ident = resolve_employee_identity(conn, DEFAULTS, regu_rows=reguh, roles=roles,
                                      payroll_pernrs={"00000005", "00000006", "00000007"})
    assert ident["method"] == "PERNR_FIELD" and ident["confidence"] == "PROVED"


def test_identity_empfg_is_pernr():
    reguh = [_reguh("20260628", "PAY001", empfg="00000005", rwbtr="1382.40"),
             _reguh("20260628", "PAY001", empfg="00000006", rwbtr="2000.00")]
    conn = FakeConnection(base_tables(reguh=reguh, reguh_has_pernr=False))
    sch = inspect_regu_schema(conn, DEFAULTS)
    roles = sch["tables"]["REGUH"]["roles"]
    ident = resolve_employee_identity(conn, DEFAULTS, regu_rows=reguh, roles=roles,
                                      payroll_pernrs={"00000005", "00000006"})
    assert ident["method"] == "EMPFG_IS_PERNR"


def test_identity_via_lifnr_is_hypothesis():
    reguh = [_reguh("20260628", "PAY001", lifnr="0000100005", empfg="", rwbtr="1382.40")]
    conn = FakeConnection(base_tables(reguh=reguh, reguh_has_pernr=False))
    sch = inspect_regu_schema(conn, DEFAULTS)
    roles = sch["tables"]["REGUH"]["roles"]
    ident = resolve_employee_identity(conn, DEFAULTS, regu_rows=reguh, roles=roles,
                                      payroll_pernrs={"00000005"})
    assert ident["method"] == "LIFNR_UNRESOLVED" and ident["confidence"] == "HYPOTHESIS"
    assert "0000100005" in ident["unresolved"]


def test_expectations_use_current_559_not_gross():
    conn = FakeConnection(base_tables())
    exps = build_payroll_payment_expectations(conn, DEFAULTS, run=RUN)
    e5 = next(e for e in exps if e.pernr == "00000005")
    assert e5.expected_payment_amount == Decimal("1382.40")   # /559 corrente, NÃO 1647,40
    # PERNR 5 tem uma versão anterior (/559 TSLIN=0) -> OBSERVED, não PROVED
    assert e5.confidence == "OBSERVED"
    assert e5.related_amounts["0029"] == "-265.00"
    # colaborador sem versão anterior -> PROVED
    e6 = next(e for e in exps if e.pernr == "00000006")
    assert e6.confidence == "PROVED" and e6.expected_payment_amount == Decimal("2000.00")


def test_tslin0_559_not_duplicated_and_classified_previous():
    conn = FakeConnection(base_tables())
    exps = build_payroll_payment_expectations(conn, DEFAULTS, run=RUN)
    e5 = next(e for e in exps if e.pernr == "00000005")
    classes = {c["class"] for c in e5.line_classes}
    assert "PREVIOUS_VERSION" in classes           # a linha /559 TSLIN=0
    assert "CURRENT_PAYMENT" in classes
    # o valor esperado é só a corrente (1382,40), a TSLIN=0 (823,70) não soma
    assert e5.expected_payment_amount == Decimal("1382.40")


def test_multiple_559_seqno_is_candidate_with_retro_reference():
    ppoix = [
        _ppoix("00000009", "01100", "/559", "500.00-", "0000000004"),   # SEQNO antigo, transferido
        _ppoix("00000009", "01101", "/559", "1200.00-", "0000000004"),  # SEQNO recente
    ]
    conn = FakeConnection(base_tables(ppoix=ppoix))
    exps = build_payroll_payment_expectations(conn, DEFAULTS, run=RUN)
    e9 = next(e for e in exps if e.pernr == "00000009")
    assert e9.confidence == "CANDIDATE"
    assert e9.expected_payment_amount == Decimal("1200.00")            # só o SEQNO mais recente
    assert any(c["class"] == "RETRO_REFERENCE" for c in e9.line_classes)


def test_multiple_payments_per_employee_kept_and_totalled():
    reguh = [_reguh("20260628", "PAY001", empfg="00000005", rwbtr="682.40", vblnr="D1"),
             _reguh("20260628", "PAY001", empfg="00000005", rwbtr="700.00", vblnr="D2")]
    ident = {"method": "EMPFG_IS_PERNR", "mapping": {}}
    from sap_payroll_analysis.payment_reconciliation import read_regu_payments
    conn = FakeConnection(base_tables(reguh=reguh))
    sch = inspect_regu_schema(conn, DEFAULTS)
    pays = read_regu_payments(conn, DEFAULTS, laufd="20260628", laufi="PAY001",
                              company=COMPANY, schema=sch)
    agg = aggregate_regu_by_employee(pays, ident)
    rec = agg["by_pernr"]["00000005"]
    assert rec["count"] == 2 and Decimal(rec["total"]) == Decimal("1382.40")


def test_reconcile_exact_match(monkeypatch):
    _prep_r3_env(monkeypatch)
    reguh = [_reguh("20260628", "PAY001", empfg="00000005", rwbtr="1382.40"),
             _reguh("20260628", "PAY001", empfg="00000006", rwbtr="2000.00"),
             _reguh("20260628", "PAY001", empfg="00000007", rwbtr="1500.00")]
    reguv = [{"LAUFD": "20260628", "LAUFI": "PAY001", "ZBUKR": COMPANY, "XECHT": "X", "WAERS": "EUR"}]
    conn = FakeConnection(base_tables(reguh=reguh, reguv=reguv))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon.totals["difference"] == "0.00"
    assert recon.classification["classification"] == "EXACT_MATCH"
    assert all(l.status == "EXACT_MATCH" for l in recon.reconciliation)


def test_reconcile_difference(monkeypatch):
    _prep_r3_env(monkeypatch)
    reguh = [_reguh("20260628", "PAY001", empfg="00000005", rwbtr="1382.40"),
             _reguh("20260628", "PAY001", empfg="00000006", rwbtr="1999.00"),   # -1,00
             _reguh("20260628", "PAY001", empfg="00000007", rwbtr="1500.00")]
    conn = FakeConnection(base_tables(reguh=reguh))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon.classification["classification"] == "DIFFERENCE"
    d6 = next(l for l in recon.reconciliation if l.pernr == "00000006")
    assert d6.status == "DIFFERENCE" and d6.difference == Decimal("1.00")


def test_reconcile_rh_only(monkeypatch):
    _prep_r3_env(monkeypatch)
    reguh = [_reguh("20260628", "PAY001", empfg="00000005", rwbtr="1382.40"),
             _reguh("20260628", "PAY001", empfg="00000006", rwbtr="2000.00")]
    conn = FakeConnection(base_tables(reguh=reguh))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    l7 = next(l for l in recon.reconciliation if l.pernr == "00000007")
    assert l7.status == "RH_ONLY"
    assert recon.classification["classification"] in ("DIFFERENCE", "PARTIAL")


def test_reconcile_regu_only(monkeypatch):
    _prep_r3_env(monkeypatch)
    reguh = [_reguh("20260628", "PAY001", empfg="00000005", rwbtr="1382.40"),
             _reguh("20260628", "PAY001", empfg="00000006", rwbtr="2000.00"),
             _reguh("20260628", "PAY001", empfg="00000007", rwbtr="1500.00"),
             _reguh("20260628", "PAY001", empfg="00009999", rwbtr="123.45")]
    conn = FakeConnection(base_tables(reguh=reguh))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    lx = next(l for l in recon.reconciliation if l.pernr == "00009999")
    assert lx.status == "REGU_ONLY"


def test_reconcile_ambiguous_identity_is_partial(monkeypatch):
    _prep_r3_env(monkeypatch)
    # EMPFG não numérico e sem PERNR -> identidade UNKNOWN -> PARTIAL
    reguh = [_reguh("20260628", "PAY001", empfg="ABC", rwbtr="1382.40")]
    conn = FakeConnection(base_tables(reguh=reguh, reguh_has_pernr=False))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon.employee_identity_mapping["method"] in ("UNKNOWN", "LIFNR_UNRESOLVED")
    assert recon.classification["classification"] == "PARTIAL"


def test_totals_close_and_diverge(monkeypatch):
    _prep_r3_env(monkeypatch)
    ok = [_reguh("20260628", "PAY001", empfg=p, rwbtr=a) for p, a in
          (("00000005", "1382.40"), ("00000006", "2000.00"), ("00000007", "1500.00"))]
    recon_ok = reconcile_payroll_payments(
        FakeConnection(base_tables(reguh=ok)), DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon_ok.totals["difference"] == "0.00"

    bad = [_reguh("20260628", "PAY001", empfg=p, rwbtr=a) for p, a in
           (("00000005", "1382.40"), ("00000006", "2000.00"), ("00000007", "1400.00"))]
    recon_bad = reconcile_payroll_payments(
        FakeConnection(base_tables(reguh=bad)), DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon_bad.totals["difference"] == "100.00"


def test_427_flagged_candidate_not_explained(monkeypatch):
    _prep_r3_env(monkeypatch)
    # diferença de total = exactamente 427,74 -> CANDIDATE, nunca "explicado"
    reguh = [_reguh("20260628", "PAY001", empfg=p, rwbtr=a) for p, a in
             (("00000005", "1382.40"), ("00000006", "2000.00"), ("00000007", "1072.26"))]
    conn = FakeConnection(base_tables(reguh=reguh))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon.totals["difference"] == "427.74"
    kinds = [d for d in recon.differences if d.get("kind")]
    assert any("427" in d["status"] and "[CANDIDATE]" in d["status"] for d in kinds)
    assert recon.classification["classification"] == "DIFFERENCE"   # não "EXACT" nem "explicado"


def test_json_and_csv_output(tmp_path, monkeypatch):
    _prep_r3_env(monkeypatch)
    reguh = [_reguh("20260628", "PAY001", empfg=p, rwbtr=a) for p, a in
             (("00000005", "1382.40"), ("00000006", "2000.00"), ("00000007", "1500.00"))]
    conn = FakeConnection(base_tables(reguh=reguh))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    from sap_payroll_analysis.payment_reconciliation import (
        write_reconciliation_csv, write_reconciliation_json,
    )
    jp = write_reconciliation_json(recon, tmp_path / "r.json")
    cp = write_reconciliation_csv(recon, tmp_path / "r.csv")
    data = json.loads(jp.read_text(encoding="utf-8"))
    assert data["totals"]["difference"] == "0.00"
    assert "PERNR;RH_EXPECTED;REGU_PAID" in cp.read_text(encoding="utf-8-sig")


def test_select_payment_run_set_reconstructs_split():
    """O pagamento do payroll dividido por 2 LAUFI no mesmo dia é reconstruído;
    um 3.º run (despesas) que introduz divergência fica de fora."""
    from sap_payroll_analysis.payment_reconciliation import (
        _rows_to_payments, resolve_employee_identity, select_payment_run_set,
    )
    reguh = [
        _reguh("20260625", "SAL1", pernr="00000005", rwbtr="1382.40"),
        _reguh("20260625", "SAL1", pernr="00000006", rwbtr="2000.00"),
        _reguh("20260625", "SAL2", pernr="00000007", rwbtr="1500.00"),
        _reguh("20260626", "EXP1", pernr="00000005", rwbtr="50.00"),   # despesas -> divergência
    ]
    conn = FakeConnection(base_tables(reguh=reguh))
    sch = inspect_regu_schema(conn, DEFAULTS)
    roles = sch["tables"]["REGUH"]["roles"]
    exps = build_payroll_payment_expectations(conn, DEFAULTS, run=RUN)
    ident = resolve_employee_identity(conn, DEFAULTS, regu_rows=reguh, roles=roles,
                                      payroll_pernrs={e.pernr for e in exps})
    pays = _rows_to_payments(reguh, roles, COMPANY)
    rs = select_payment_run_set(exps, pays, ident,
                                primary={"laufd": "20260625", "laufi": "SAL1"})
    chosen = {tuple(x) for x in rs["run_ids"]}
    assert ("20260625", "SAL1") in chosen and ("20260625", "SAL2") in chosen
    assert ("20260626", "EXP1") not in chosen          # introduz divergência -> excluído
    assert rs["exact_pernr"] == 3 and rs["diff_pernr"] == 0
    assert rs["confidence"] == "HIGH_CONFIDENCE"


def test_reconcile_multi_run_exact_match(monkeypatch):
    """reconcile end-to-end: 2 salary runs + 1 expense run -> EXACT_MATCH usando só os 2."""
    _prep_r3_env(monkeypatch)
    reguh = [
        _reguh("20260625", "SAL1", pernr="00000005", rwbtr="1382.40"),
        _reguh("20260625", "SAL1", pernr="00000006", rwbtr="2000.00"),
        _reguh("20260625", "SAL2", pernr="00000007", rwbtr="1500.00"),
        _reguh("20260626", "EXP1", pernr="00000009", rwbtr="9999.00"),  # PERNR fora do payroll
    ]
    conn = FakeConnection(base_tables(reguh=reguh))
    recon = reconcile_payroll_payments(conn, DEFAULTS, run=RUN, company=COMPANY, period=PERIOD)
    assert recon.classification["classification"] == "EXACT_MATCH"
    assert recon.totals["difference"] == "0.00"
    assert recon.totals["payment_runs_used"] == 2
    assert set(recon.totals["payment_run_ids"]) == {"SAL1", "SAL2"}
    assert all(l.status == "EXACT_MATCH" for l in recon.reconciliation)


def test_run_set_keeps_anchor_even_with_difference():
    """O run primário (âncora) entra sempre no conjunto, mesmo com divergência."""
    from sap_payroll_analysis.payment_reconciliation import (
        _rows_to_payments, resolve_employee_identity, select_payment_run_set,
    )
    reguh = [_reguh("20260625", "SAL1", pernr="00000005", rwbtr="1000.00"),  # /559=1382,40
             _reguh("20260625", "SAL1", pernr="00000006", rwbtr="2000.00")]
    conn = FakeConnection(base_tables(reguh=reguh))
    sch = inspect_regu_schema(conn, DEFAULTS)
    roles = sch["tables"]["REGUH"]["roles"]
    exps = build_payroll_payment_expectations(conn, DEFAULTS, run=RUN)
    ident = resolve_employee_identity(conn, DEFAULTS, regu_rows=reguh, roles=roles,
                                      payroll_pernrs={e.pernr for e in exps})
    pays = _rows_to_payments(reguh, roles, COMPANY)
    rs = select_payment_run_set(exps, pays, ident,
                                primary={"laufd": "20260625", "laufi": "SAL1"})
    assert [tuple(x) for x in rs["run_ids"]] == [("20260625", "SAL1")]
    assert rs["diff_pernr"] == 1


def _prep_r3_env(monkeypatch):
    for s, v in (("USER", "U"), ("PASSWD", "P"), ("ASHOST", "10.1.1.101"),
                 ("SYSNR", "00"), ("CLIENT", "100")):
        monkeypatch.setenv(f"SAP_R3_{s}", v)

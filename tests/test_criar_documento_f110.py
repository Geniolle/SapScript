from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
F110_DIR = next(p for p in (ROOT / "Processos").iterdir() if p.is_dir() and "Sim" in p.name)
if str(F110_DIR) not in sys.path:
    sys.path.insert(0, str(F110_DIR))

import criar_documento_teste_f110 as f110  # noqa: E402


def test_next_reference_advances_sequence_from_bkpf(monkeypatch):
    captured = {}

    def fake_read_table_with_fallbacks(conn, table_name, field_sets, options=None, rowcount=10):
        captured["conn"] = conn
        captured["table_name"] = table_name
        captured["field_sets"] = field_sets
        captured["options"] = options
        captured["rowcount"] = rowcount
        return (
            [
                {"XBLNR": "UAT-F110-TEST01"},
                {"XBLNR": "UAT-F110-TEST02"},
                {"XBLNR": "UAT-F110-TESTXX"},
            ],
            "XBLNR",
        )

    monkeypatch.setattr(f110, "read_table_with_fallbacks", fake_read_table_with_fallbacks)

    runner = f110.FiDocumentPoster()
    runner.conn = object()

    assert runner._next_reference("2010", "UAT-F110-TEST") == "UAT-F110-TEST03"
    assert captured["table_name"] == "BKPF"
    assert captured["field_sets"] == [["XBLNR"], ["BUKRS", "XBLNR"]]
    assert captured["options"] == [
        "BUKRS = '2010'",
        "AND XBLNR LIKE 'UAT-F110-TEST%'",
    ]
    assert captured["rowcount"] == 500


def test_build_document_header_uses_reference_as_xblnr_source():
    payload = f110.DocumentInput(
        system_key="QAD",
        company_code="2010",
        vendor="10000040",
        gl_account="12010741",
        amount=f110._normalize_amount("88,88"),
        reference="UAT-F110-TEST03",
    )

    header = f110.FiDocumentPoster._build_document_header(payload, "SAPUSER")

    assert header["REF_DOC_NO"] == "UAT-F110-TEST03"


def test_build_tables_includes_vendor_withholding_tax_information():
    payload = f110.DocumentInput(
        system_key="QAD",
        company_code="2010",
        vendor="10000040",
        gl_account="12010741",
        amount=f110._normalize_amount("88,88"),
        reference="UAT-F110-TEST03",
    )

    tables = f110.FiDocumentPoster._build_tables(
        payload,
        payment_terms="",
        company_currency="EUR",
        withholding_tax_rows=[
            {
                "ITEMNO_ACC": "0000000001",
                "WT_TYPE": "IR",
                "WT_CODE": "01",
                "BAS_AMT_TC": "88.88",
            }
        ],
    )

    assert tables["ACCOUNTWT"] == [
        {
            "ITEMNO_ACC": "0000000001",
            "WT_TYPE": "IR",
            "WT_CODE": "01",
            "BAS_AMT_TC": "88.88",
        }
    ]
    assert tables["ACCOUNTPAYABLE"][0]["W_TAX_CODE"] == "01"

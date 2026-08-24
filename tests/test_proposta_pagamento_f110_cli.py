from __future__ import annotations

import sys
from datetime import date, timedelta
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
F110_DIR = ROOT / "Processos" / "UAT Simulação"
if str(F110_DIR) not in sys.path:
    sys.path.insert(0, str(F110_DIR))

import proposta_pagamento_f110 as f110  # noqa: E402


def test_build_parser_defaults_to_proposal_only():
    args = f110.build_parser().parse_args([])
    assert args.proposal_only is True
    assert args.posting_date == date.today().strftime("%Y%m%d")
    assert args.run_date == date.today().strftime("%Y%m%d")
    assert args.docs_entered_up_to == (date.today() + timedelta(days=1)).strftime("%Y%m%d")
    assert args.identification == "AUTO"


def test_build_parser_can_disable_proposal_only():
    args = f110.build_parser().parse_args(["--execute-payment"])
    assert args.proposal_only is False


def test_build_selinfo_uses_real_rff110s_field_names():
    payload = f110.ProposalInput(
        system_key="QAD",
        company_code="2010",
        vendor="10000040",
        document_number="6050000002",
        fiscal_year="2026",
        run_date="20260822",
        identification="UAT1",
        posting_date="20260821",
        docs_entered_up_to="20260822",
        payment_method="S",
        proposal_only=True,
        jobclass="C",
        wait_seconds=120,
    )
    runner = f110.F110ProposalRunner()
    selinfo = runner._build_selinfo(payload)
    assert [row["SELNAME"] for row in selinfo] == [
        "SEL_BUKR",
        "SEL_KRED",
        "PAR_ZWE",
        "PAR_XVL",
        "PAR_LFD",
        "PAR_LFID",
        "PAR_NEDA",
        "PAR_BUDA",
        "PAR_GRDA",
        "PAR_XFA",
        "PAR_XZE",
        "PAR_XBL",
        "PAR_MITD",
        "PAR_TEX1",
        "PAR_LIS1",
    ]
    assert selinfo[0]["LOW"] == "2010"
    assert selinfo[1]["LOW"] == "0010000040"
    assert selinfo[1]["HIGH"] == ""
    assert selinfo[4]["LOW"] == "20260822"
    assert selinfo[5]["LOW"] == "UAT1"
    assert selinfo[6]["LOW"] == "20260822"
    assert selinfo[9]["LOW"] == "X"
    assert selinfo[10]["LOW"] == "X"
    assert selinfo[11]["LOW"] == "X"
    assert selinfo[12]["LOW"] == "X"
    assert selinfo[13]["LOW"] == "BKPF-BELNR"
    assert selinfo[14]["LOW"] == "6050000002"


def test_next_identification_picks_first_free_sequence():
    class FakeConn:
        def call(self, function_name, **kwargs):
            assert function_name == "RFC_READ_TABLE"
            assert kwargs["QUERY_TABLE"] == "REGUV"
            return {
                "FIELDS": [{"FIELDNAME": "LAUFI"}],
                "DATA": [],
            }

    runner = f110.F110ProposalRunner()
    runner.conn = FakeConn()
    assert runner._next_identification("20260822") == "UAT01"


def test_next_identification_skips_existing_sequence():
    class FakeConn:
        def call(self, function_name, **kwargs):
            assert function_name == "RFC_READ_TABLE"
            assert kwargs["QUERY_TABLE"] == "REGUV"
            return {
                "FIELDS": [{"FIELDNAME": "LAUFI"}],
                "DATA": [
                    {"WA": "UAT01"},
                    {"WA": "UAT02"},
                ],
            }

    runner = f110.F110ProposalRunner()
    runner.conn = FakeConn()
    assert runner._next_identification("20260822") == "UAT03"


@pytest.mark.parametrize(
    ("value", "default", "expected"),
    [
        (None, True, True),
        (None, False, False),
        (True, False, True),
        (False, True, False),
        ("sim", False, True),
        ("nao", True, False),
        ("false", True, False),
        ("x", False, True),
    ],
)
def test_coerce_bool_handles_cli_and_programmatic_inputs(value, default, expected):
    assert f110._coerce_bool(value, default=default) is expected

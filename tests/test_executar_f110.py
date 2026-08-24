from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest


ROOT = Path(__file__).resolve().parents[1]
F110_DIR = ROOT / "Processos" / "UAT Simulação"
if str(F110_DIR) not in sys.path:
    sys.path.insert(0, str(F110_DIR))

import executar_f110 as f110  # noqa: E402


def _stage_result(**kwargs):
    return SimpleNamespace(**kwargs)


def test_executar_creates_document_by_default(monkeypatch):
    calls: list[tuple[str, dict[str, object]]] = []

    def fake_create(**kwargs):
        calls.append(("create", kwargs))
        assert kwargs["system_key"] == f110.DEFAULT_SYSTEM_KEY
        assert kwargs["company_code"] == f110.DEFAULT_COMPANY_CODE
        assert kwargs["vendor"] == f110.DEFAULT_VENDOR
        assert kwargs["gl_account"] == f110.DEFAULT_GL_ACCOUNT
        assert kwargs["amount"] == f110.DEFAULT_AMOUNT
        assert kwargs["currency"] == f110.DEFAULT_CURRENCY
        assert kwargs["document_date"] == f110._today_yyyymmdd()
        assert kwargs["posting_date"] == f110._today_yyyymmdd()
        assert kwargs["payment_method"] == f110.DEFAULT_PAYMENT_METHOD
        assert kwargs["doc_type"] == f110.DEFAULT_DOC_TYPE
        assert kwargs["reference"] == f110.DEFAULT_REFERENCE
        assert kwargs["header_text"] == f110.DEFAULT_HEADER_TEXT
        assert kwargs["item_text"] == f110.DEFAULT_ITEM_TEXT
        assert kwargs["check_only"] is False
        return _stage_result(
            status="posted",
            system_id="S4Q",
            posted_belnr="6000000001",
            posted_gjahr="2026",
            reasons=[],
        )

    def fake_simulate(**kwargs):
        raise AssertionError("simulacao nao devia ser chamada no passo padrao")

    def fake_proposal(**kwargs):
        raise AssertionError("proposta nao devia ser chamada no passo padrao")

    monkeypatch.setattr(f110, "executar_criar_documento", fake_create)
    monkeypatch.setattr(f110, "executar_simulacao_f110", fake_simulate)
    monkeypatch.setattr(f110, "executar_proposta_pagamento", fake_proposal)

    result = f110.executar()

    assert result.status == "created"
    assert result.created_document_number == "6000000001"
    assert result.system_id == "S4Q"
    assert [name for name, _ in calls] == ["create"]


def test_executar_runs_three_stages_in_order_with_full_step(monkeypatch):
    calls: list[tuple[str, dict[str, object]]] = []

    def fake_create(**kwargs):
        calls.append(("create", kwargs))
        assert kwargs["system_key"] == f110.DEFAULT_SYSTEM_KEY
        assert kwargs["company_code"] == f110.DEFAULT_COMPANY_CODE
        assert kwargs["vendor"] == f110.DEFAULT_VENDOR
        assert kwargs["gl_account"] == f110.DEFAULT_GL_ACCOUNT
        assert kwargs["amount"] == f110.DEFAULT_AMOUNT
        assert kwargs["currency"] == f110.DEFAULT_CURRENCY
        assert kwargs["document_date"] == f110._today_yyyymmdd()
        assert kwargs["posting_date"] == f110._today_yyyymmdd()
        assert kwargs["payment_method"] == f110.DEFAULT_PAYMENT_METHOD
        assert kwargs["doc_type"] == f110.DEFAULT_DOC_TYPE
        assert kwargs["reference"] == f110.DEFAULT_REFERENCE
        assert kwargs["header_text"] == f110.DEFAULT_HEADER_TEXT
        assert kwargs["item_text"] == f110.DEFAULT_ITEM_TEXT
        assert kwargs["check_only"] is False
        return _stage_result(
            status="posted",
            system_id="S4Q",
            posted_belnr="6000000001",
            posted_gjahr="2026",
            reasons=[],
        )

    def fake_simulate(**kwargs):
        calls.append(("simulate", kwargs))
        assert kwargs["document_number"] == "6000000001"
        assert kwargs["fiscal_year"] == "2026"
        return _stage_result(status="eligible", eligible=True, reasons=[])

    def fake_proposal(**kwargs):
        calls.append(("proposal", kwargs))
        assert kwargs["document_number"] == "6000000001"
        assert kwargs["proposal_only"] is True
        return _stage_result(status="finished", scheduled=True, finished=True, reasons=[])

    monkeypatch.setattr(f110, "executar_criar_documento", fake_create)
    monkeypatch.setattr(f110, "executar_simulacao_f110", fake_simulate)
    monkeypatch.setattr(f110, "executar_proposta_pagamento", fake_proposal)

    result = f110.executar(step="full")

    assert result.status == "finished"
    assert result.created_document_number == "6000000001"
    assert result.system_id == "S4Q"
    assert [name for name, _ in calls] == ["create", "simulate", "proposal"]


def test_executar_blocks_when_simulation_is_not_eligible(monkeypatch):
    calls: list[str] = []

    def fake_create(**kwargs):
        calls.append("create")
        return _stage_result(
            status="posted",
            system_id="S4Q",
            posted_belnr="6000000002",
            posted_gjahr="2026",
            reasons=[],
        )

    def fake_simulate(**kwargs):
        calls.append("simulate")
        return _stage_result(status="blocked", eligible=False, reasons=["nao elegivel"])

    def fake_proposal(**kwargs):
        calls.append("proposal")
        raise AssertionError("proposal nao devia ser chamada quando a simulacao bloqueia")

    monkeypatch.setattr(f110, "executar_criar_documento", fake_create)
    monkeypatch.setattr(f110, "executar_simulacao_f110", fake_simulate)
    monkeypatch.setattr(f110, "executar_proposta_pagamento", fake_proposal)

    result = f110.executar(step="full")

    assert result.status == "blocked"
    assert result.simulation_status == "blocked"
    assert calls == ["create", "simulate"]


def test_main_uses_default_arguments(monkeypatch):
    captured: dict[str, object] = {}

    def fake_executar(**kwargs):
        captured.update(kwargs)
        return _stage_result(status="finished")

    monkeypatch.setattr(f110, "executar", fake_executar)

    assert f110.main([]) == 0
    assert captured["system_key"] == f110.DEFAULT_SYSTEM_KEY
    assert captured["company_code"] == f110.DEFAULT_COMPANY_CODE
    assert captured["vendor"] == f110.DEFAULT_VENDOR
    assert captured["gl_account"] == f110.DEFAULT_GL_ACCOUNT
    assert captured["proposal_only"] is True
    assert captured["check_only"] is False
    assert captured["step"] == f110.DEFAULT_STEP


def test_build_parser_supports_execute_payment_and_check_only():
    args = f110.build_parser().parse_args(["--execute-payment", "--check-only"])
    assert args.proposal_only is False
    assert args.check_only is True
    assert args.step == f110.DEFAULT_STEP


def test_build_parser_supports_full_step():
    args = f110.build_parser().parse_args(["--step", "full"])
    assert args.step == "full"


@pytest.mark.parametrize(
    ("value", "default", "expected"),
    [
        (None, True, True),
        (None, False, False),
        ("sim", False, True),
        ("nao", True, False),
        ("x", False, True),
        (False, True, False),
    ],
)
def test_to_bool_handles_mixed_inputs(value, default, expected):
    assert f110._to_bool(value, default=default) is expected

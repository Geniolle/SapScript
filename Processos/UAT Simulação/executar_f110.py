# -*- coding: utf-8 -*-
"""
Orquestrador unico para o fluxo UAT da F110.

Sequencia:
1. Criar documento FI de teste;
2. Simular elegibilidade RFF110S;
3. Executar a proposta F110 ou o lancamento de pagamento.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass, field
from datetime import date
from typing import Any

from criar_documento_teste_f110 import executar as executar_criar_documento
from proposta_pagamento_f110 import (
    DEFAULT_DOC_DATE,
    DEFAULT_IDENTIFICATION,
    DEFAULT_PAYMENT_METHOD,
    DEFAULT_RUN_DATE,
    executar as executar_proposta_pagamento,
)
from rff110s_uat_orchestrator import executar as executar_simulacao_f110


# =============================================================================
# (1) CONSTANTES E MODELOS
# =============================================================================

DEFAULT_SYSTEM_KEY = "QAD"
DEFAULT_COMPANY_CODE = "2010"
DEFAULT_VENDOR = "10000040"
DEFAULT_GL_ACCOUNT = "12010741"
DEFAULT_AMOUNT = "88,88"
DEFAULT_CURRENCY = "EUR"
DEFAULT_DOC_TYPE = "KR"
DEFAULT_REFERENCE = "UAT-F110-TEST"
DEFAULT_HEADER_TEXT = "UAT TESTE F110"
DEFAULT_ITEM_TEXT = "UAT F110 TESTE"
DEFAULT_STEP = "create-document"


def _today_yyyymmdd() -> str:
    return date.today().strftime("%Y%m%d")


def _to_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value or "").strip().lower()
    if not text:
        return default
    if text in {"1", "true", "yes", "on", "sim", "s", "x"}:
        return True
    if text in {"0", "false", "no", "off", "nao", "não", "n"}:
        return False
    return default


@dataclass
class F110FlowResult:
    status: str
    system_id: str
    company_code: str
    vendor: str
    document_number: str
    fiscal_year: str
    proposal_only: bool
    created_document_number: str = ""
    created_fiscal_year: str = ""
    simulation_status: str = ""
    proposal_status: str = ""
    reasons: list[str] = field(default_factory=list)
    document_result: Any = None
    simulation_result: Any = None
    proposal_result: Any = None


# =============================================================================
# (2) ORQUESTRACAO DO FLUXO
# =============================================================================

def _stage_reason(prefix: str, result: Any) -> str:
    status = str(getattr(result, "status", "") or "").strip()
    if not status:
        return prefix
    return f"{prefix} [{status}]"


def _print_summary(result: F110FlowResult) -> None:
    lines: list[str] = []
    lines.append("=" * 84)
    lines.append("UAT FLUXO F110")
    lines.append("=" * 84)
    lines.append(f"Sistema        : {result.system_id or '?'}")
    lines.append(f"Empresa        : {result.company_code}")
    lines.append(f"Fornecedor     : {result.vendor}")
    lines.append(f"Documento      : {result.document_number or '?'}")
    lines.append(f"Exercicio      : {result.fiscal_year}")
    lines.append(f"Modo           : {'PROPOSTA' if result.proposal_only else 'PAGAMENTO'}")
    lines.append(f"Estado         : {result.status.upper()}")
    if result.created_document_number:
        lines.append(f"Documento novo : {result.created_document_number}")
    if result.created_fiscal_year:
        lines.append(f"Exercicio novo : {result.created_fiscal_year}")
    if result.simulation_status:
        lines.append(f"Simulacao      : {result.simulation_status.upper()}")
    if result.proposal_status:
        lines.append(f"Proposta/F110  : {result.proposal_status.upper()}")
    lines.append("")
    lines.append("Motivos")
    lines.append("-" * 84)
    if result.reasons:
        for index, reason in enumerate(result.reasons, start=1):
            lines.append(f"{index}. {reason}")
    else:
        lines.append("Sem observacoes adicionais.")
    print("\n".join(lines))


def executar(**kwargs: Any) -> F110FlowResult:
    system_key = str(kwargs.get("system_key") or kwargs.get("sap_system") or DEFAULT_SYSTEM_KEY).strip()
    company_code = str(kwargs.get("company_code") or DEFAULT_COMPANY_CODE).strip()
    vendor = str(kwargs.get("vendor") or DEFAULT_VENDOR).strip()
    gl_account = str(kwargs.get("gl_account") or DEFAULT_GL_ACCOUNT).strip()
    amount = str(kwargs.get("amount") or DEFAULT_AMOUNT).strip()
    currency = str(kwargs.get("currency") or DEFAULT_CURRENCY).strip().upper() or DEFAULT_CURRENCY
    document_date = str(kwargs.get("document_date") or _today_yyyymmdd()).strip()
    posting_date = str(kwargs.get("posting_date") or _today_yyyymmdd()).strip()
    baseline_date = str(kwargs.get("baseline_date") or "").strip()
    payment_terms = str(kwargs.get("payment_terms") or "").strip().upper()
    payment_method = str(kwargs.get("payment_method") or DEFAULT_PAYMENT_METHOD).strip().upper()
    doc_type = str(kwargs.get("doc_type") or DEFAULT_DOC_TYPE).strip().upper()
    reference = str(kwargs.get("reference") or DEFAULT_REFERENCE).strip()
    header_text = str(kwargs.get("header_text") or DEFAULT_HEADER_TEXT).strip()
    item_text = str(kwargs.get("item_text") or DEFAULT_ITEM_TEXT).strip()
    proposal_only = _to_bool(kwargs.get("proposal_only"), default=True)
    if _to_bool(kwargs.get("execute_payment"), default=False):
        proposal_only = False
    force_payment = _to_bool(kwargs.get("force_payment"), default=False)
    step = str(kwargs.get("step") or DEFAULT_STEP).strip().lower()
    run_date = str(kwargs.get("run_date") or DEFAULT_RUN_DATE).strip()
    identification = str(kwargs.get("identification") or DEFAULT_IDENTIFICATION).strip()
    docs_entered_up_to = str(kwargs.get("docs_entered_up_to") or DEFAULT_DOC_DATE).strip()
    wait_seconds = int(kwargs.get("wait_seconds") or 120)
    check_only = _to_bool(kwargs.get("check_only"), default=False)

    reasons: list[str] = []

    # -------------------------------------------------------------------------
    # (2.1) CRIACAO DO DOCUMENTO DE TESTE
    # -------------------------------------------------------------------------
    document_result = executar_criar_documento(
        system_key=system_key,
        company_code=company_code,
        vendor=vendor,
        gl_account=gl_account,
        amount=amount,
        currency=currency,
        document_date=document_date,
        posting_date=posting_date,
        baseline_date=baseline_date,
        payment_terms=payment_terms,
        payment_method=payment_method,
        doc_type=doc_type,
        reference=reference,
        header_text=header_text,
        item_text=item_text,
        check_only=check_only,
    )

    created_document_number = str(getattr(document_result, "posted_belnr", "") or "").strip()
    created_fiscal_year = str(getattr(document_result, "posted_gjahr", "") or "").strip() or posting_date[:4]
    if not created_document_number:
        created_document_number = str(kwargs.get("document_number") or "").strip()

    if getattr(document_result, "status", "") != "posted":
        reasons.append(_stage_reason("Criacao do documento nao concluiu com sucesso", document_result))
        result = F110FlowResult(
            status="blocked",
            system_id=str(getattr(document_result, "system_id", "") or "").strip(),
            company_code=company_code,
            vendor=vendor,
            document_number=created_document_number,
            fiscal_year=created_fiscal_year,
            proposal_only=proposal_only,
            created_document_number=created_document_number,
            created_fiscal_year=created_fiscal_year,
            reasons=reasons + list(getattr(document_result, "reasons", []) or []),
            document_result=document_result,
        )
        _print_summary(result)
        return result

    if step in {"create-document", "create_doc", "create"} or check_only:
        result = F110FlowResult(
            status="checked" if check_only else "created",
            system_id=str(getattr(document_result, "system_id", "") or "").strip(),
            company_code=company_code,
            vendor=vendor,
            document_number=created_document_number,
            fiscal_year=created_fiscal_year,
            proposal_only=proposal_only,
            created_document_number=created_document_number,
            created_fiscal_year=created_fiscal_year,
            reasons=list(getattr(document_result, "reasons", []) or []),
            document_result=document_result,
        )
        _print_summary(result)
        return result

    # -------------------------------------------------------------------------
    # (2.2) SIMULACAO RFF110S
    # -------------------------------------------------------------------------
    simulation_result = executar_simulacao_f110(
        system_key=system_key,
        company_code=company_code,
        vendor=vendor,
        document_number=created_document_number,
        fiscal_year=created_fiscal_year,
        proposal_date=posting_date,
        payment_method=payment_method,
    )

    simulation_status = str(getattr(simulation_result, "status", "") or "").strip()
    if step in {"simulate", "simulation", "rff110s"}:
        result = F110FlowResult(
            status="simulated",
            system_id=str(getattr(document_result, "system_id", "") or "").strip(),
            company_code=company_code,
            vendor=vendor,
            document_number=created_document_number,
            fiscal_year=created_fiscal_year,
            proposal_only=proposal_only,
            created_document_number=created_document_number,
            created_fiscal_year=created_fiscal_year,
            simulation_status=simulation_status,
            reasons=list(getattr(document_result, "reasons", []) or [])
            + list(getattr(simulation_result, "reasons", []) or []),
            document_result=document_result,
            simulation_result=simulation_result,
        )
        _print_summary(result)
        return result

    if not getattr(simulation_result, "eligible", False) and not force_payment:
        reasons.append(_stage_reason("Simulacao nao indicou elegibilidade para prosseguir", simulation_result))
        reasons.extend(list(getattr(simulation_result, "reasons", []) or []))
        result = F110FlowResult(
            status="blocked",
            system_id=str(getattr(document_result, "system_id", "") or "").strip(),
            company_code=company_code,
            vendor=vendor,
            document_number=created_document_number,
            fiscal_year=created_fiscal_year,
            proposal_only=proposal_only,
            created_document_number=created_document_number,
            created_fiscal_year=created_fiscal_year,
            simulation_status=simulation_status,
            reasons=reasons,
            document_result=document_result,
            simulation_result=simulation_result,
        )
        _print_summary(result)
        return result

    # -------------------------------------------------------------------------
    # (2.3) PROPOSTA / PAGAMENTO F110
    # -------------------------------------------------------------------------
    proposal_result = executar_proposta_pagamento(
        system_key=system_key,
        company_code=company_code,
        vendor=vendor,
        document_number=created_document_number,
        fiscal_year=created_fiscal_year,
        run_date=run_date,
        identification=identification,
        posting_date=posting_date,
        docs_entered_up_to=docs_entered_up_to or posting_date,
        payment_method=payment_method,
        proposal_only=proposal_only,
        wait_seconds=wait_seconds,
    )

    proposal_status = str(getattr(proposal_result, "status", "") or "").strip()
    final_status = "finished" if proposal_status in {"finished", "posted"} else proposal_status or "scheduled"
    if final_status == "scheduled" and not proposal_only:
        final_status = "scheduled"

    reasons.extend(list(getattr(document_result, "reasons", []) or []))
    reasons.extend(list(getattr(simulation_result, "reasons", []) or []))
    reasons.extend(list(getattr(proposal_result, "reasons", []) or []))

    result = F110FlowResult(
        status=final_status,
        system_id=str(getattr(document_result, "system_id", "") or "").strip(),
        company_code=company_code,
        vendor=vendor,
        document_number=created_document_number,
        fiscal_year=created_fiscal_year,
        proposal_only=proposal_only,
        created_document_number=created_document_number,
        created_fiscal_year=created_fiscal_year,
        simulation_status=simulation_status,
        proposal_status=proposal_status,
        reasons=reasons,
        document_result=document_result,
        simulation_result=simulation_result,
        proposal_result=proposal_result,
    )
    _print_summary(result)
    return result


# =============================================================================
# (3) CLI / ENTRYPOINT
# =============================================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Executar o fluxo UAT completo da F110.")
    parser.add_argument("--sap-system", default=DEFAULT_SYSTEM_KEY, help="Chave de sistema SAP. Padrao: QAD")
    parser.add_argument("--company-code", default=DEFAULT_COMPANY_CODE, help="Codigo da empresa (BUKRS)")
    parser.add_argument("--vendor", default=DEFAULT_VENDOR, help="Numero do fornecedor")
    parser.add_argument("--gl-account", default=DEFAULT_GL_ACCOUNT, help="Conta contabil de contrapartida")
    parser.add_argument("--amount", default=DEFAULT_AMOUNT, help="Valor do documento")
    parser.add_argument("--currency", default=DEFAULT_CURRENCY, help="Moeda do documento")
    parser.add_argument("--document-date", default=_today_yyyymmdd(), help="Data do documento em YYYYMMDD")
    parser.add_argument("--posting-date", default=_today_yyyymmdd(), help="Data de lancamento em YYYYMMDD")
    parser.add_argument("--baseline-date", default="", help="Data base de pagamento em YYYYMMDD")
    parser.add_argument("--payment-terms", default="", help="Condicao de pagamento do fornecedor")
    parser.add_argument("--payment-method", default=DEFAULT_PAYMENT_METHOD, help="Metodo de pagamento")
    parser.add_argument("--doc-type", default=DEFAULT_DOC_TYPE, help="Tipo de documento SAP")
    parser.add_argument("--reference", default=DEFAULT_REFERENCE, help="Referencia externa do documento")
    parser.add_argument("--header-text", default=DEFAULT_HEADER_TEXT, help="Texto do cabecalho")
    parser.add_argument("--item-text", default=DEFAULT_ITEM_TEXT, help="Texto das linhas")
    parser.add_argument("--run-date", default=DEFAULT_RUN_DATE, help="Data da execucao do job em YYYYMMDD")
    parser.add_argument("--identification", default=DEFAULT_IDENTIFICATION, help="Identificacao da F110")
    parser.add_argument("--docs-entered-up-to", default=DEFAULT_DOC_DATE, help="Data limite dos documentos em YYYYMMDD")
    parser.add_argument("--wait-seconds", type=int, default=120, help="Tempo maximo de espera pelo job")
    parser.add_argument("--check-only", action="store_true", help="Executa apenas a criacao do documento, sem continuar")
    parser.add_argument(
        "--step",
        default=DEFAULT_STEP,
        choices=("create-document", "simulate", "proposal", "full"),
        help="Passo do fluxo F110 a executar",
    )
    parser.add_argument(
        "--proposal-only",
        dest="proposal_only",
        action="store_true",
        help="Executar apenas a proposta F110",
    )
    parser.add_argument(
        "--execute-payment",
        dest="proposal_only",
        action="store_false",
        help="Executar o lancamento de pagamento",
    )
    parser.set_defaults(proposal_only=True)
    parser.add_argument("--force-payment", action="store_true", help="Ignorar bloqueio da simulacao e prosseguir")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    result = executar(
        system_key=args.sap_system,
        company_code=args.company_code,
        vendor=args.vendor,
        gl_account=args.gl_account,
        amount=args.amount,
        currency=args.currency,
        document_date=args.document_date,
        posting_date=args.posting_date,
        baseline_date=args.baseline_date,
        payment_terms=args.payment_terms,
        payment_method=args.payment_method,
        doc_type=args.doc_type,
        reference=args.reference,
        header_text=args.header_text,
        item_text=args.item_text,
        run_date=args.run_date,
        identification=args.identification,
        docs_entered_up_to=args.docs_entered_up_to,
        wait_seconds=args.wait_seconds,
        check_only=args.check_only,
        step=args.step,
        proposal_only=args.proposal_only,
        force_payment=args.force_payment,
    )
    return 0 if result.status in {"finished", "scheduled", "checked", "created", "simulated"} else 1


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import dataclasses
import json
import os
from datetime import date, timedelta
from typing import Any

import requests


class F110ProposalJobError(RuntimeError):
    pass


def default_next_due_date() -> str:
    return (date.today() + timedelta(days=1)).isoformat()


def build_f110_proposal_payload(params: dict[str, Any]) -> dict[str, Any]:
    payload = dict(params or {})
    payload["environment"] = str(payload.get("environment") or "QAD").strip().upper()
    payload["operation_type"] = str(payload.get("operation_type") or "cobranca").strip().lower()
    payload["company_code"] = str(payload.get("company_code") or "").strip().upper()
    payload["payment_method"] = str(payload.get("payment_method") or "").strip().upper()
    payload["account_number"] = str(payload.get("account_number") or "").strip().upper()
    payload["posting_date"] = str(payload.get("posting_date") or "").strip()
    payload["next_due_date"] = str(payload.get("next_due_date") or "").strip() or default_next_due_date()
    payload["document_number"] = str(payload.get("document_number") or "").strip().upper()
    return payload


def update_job_params_via_api(job_id: str, new_params: dict[str, Any]) -> dict[str, Any]:
    api_base_url = (
        os.getenv("SAP_API_BASE_URL", "").strip().rstrip("/")
        or os.getenv("API_BASE_URL", "").strip().rstrip("/")
    )
    if not api_base_url:
        raise F110ProposalJobError("API base URL não definido para atualizar o job.")

    worker_token = os.getenv("WORKER_TOKEN", "change-me")
    response = requests.post(
        f"{api_base_url}/api/jobs/{job_id}/params",
        headers={"X-Worker-Token": worker_token},
        json={"params": new_params},
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def run_f110_proposal_job(
    *,
    job_id: str,
    params: dict[str, Any],
    run_f110_proposal: Any,
) -> tuple[str, str]:
    payload = build_f110_proposal_payload(params)

    result = run_f110_proposal(
        payload["environment"],
        payload["operation_type"],
        company_code=payload["company_code"],
        payment_method=payload["payment_method"],
        account_number=payload["account_number"],
        posting_date=payload["posting_date"],
        next_due_date=payload["next_due_date"],
        document_number=payload["document_number"],
    )

    result_payload = dataclasses.asdict(result)
    result_json = json.dumps(result_payload, ensure_ascii=False)

    try:
        update_job_params_via_api(job_id, {"f110_proposal_result": result_payload})
    except Exception as exc:
        raise F110ProposalJobError(
            f"Não foi possível gravar o resultado F110 no job: {exc}"
        ) from exc

    log = (
        "Proposta F110 executada pelo worker Windows.\n"
        f"Ambiente: {payload['environment']}\n"
        f"Operação: {payload['operation_type']}\n"
        f"Empresa: {payload['company_code']}\n"
        f"Conta: {payload['account_number']}\n"
        f"Forma de pagamento: {payload['payment_method']}\n"
        f"NEXT: {payload['next_due_date']}\n"
        f"Estado: {result.status}\n"
        f"Mensagem: {result.message}"
    )
    return result_json, log

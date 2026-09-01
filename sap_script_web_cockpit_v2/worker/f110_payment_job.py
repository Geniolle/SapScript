from __future__ import annotations

import dataclasses
import json
from typing import Any

from .f110_proposal_job import build_f110_proposal_payload as build_f110_payment_payload
from .f110_proposal_job import F110ProposalJobError, update_job_params_via_api


def run_f110_payment_job(
    *,
    job_id: str,
    params: dict[str, Any],
    run_f110_payment: Any,
) -> tuple[str, str]:
    payload = build_f110_payment_payload(params)

    result = run_f110_payment(
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
        update_job_params_via_api(job_id, {"f110_payment_result": result_payload})
    except Exception as exc:
        raise F110ProposalJobError(
            f"Não foi possível gravar o resultado do pagamento F110 no job: {exc}"
        ) from exc

    log = (
        "Pagamento F110 executado pelo worker Windows.\n"
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

from __future__ import annotations

import json
import os
import sys
from datetime import date, timedelta
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import load_project_env
from sap_rfc.f110_service import run_f110_payment, run_f110_proposal
from sap_script_web_cockpit_v2.worker.f110_proposal_job import build_f110_proposal_payload


def _ascii(obj: object) -> str:
    return json.dumps(obj, ensure_ascii=False, indent=2).encode("ascii", "backslashreplace").decode()


def main() -> int:
    load_project_env(REPO_ROOT)

    document_number = str(
        (sys.argv[1] if len(sys.argv) > 1 else "")
        or os.getenv("SAP_FI_TEST_DOCUMENT_NUMBER", "6050000075")
        or ""
    ).strip()
    if not document_number:
        raise RuntimeError("SAP_FI_TEST_DOCUMENT_NUMBER não está definido.")

    proposal_payload = {
        "environment": "QAD",
        "operation_type": "pagamento",
        "company_code": "2010",
        "payment_method": "S",
        "account_number": str(os.getenv("SAP_FI_VENDOR_ACCOUNT", "0010000040") or "").strip(),
        "posting_date": date.today().isoformat(),
        "next_due_date": (date.today() + timedelta(days=1)).isoformat(),
        "document_number": document_number,
    }

    if not proposal_payload["account_number"]:
        raise RuntimeError("Conta de fornecedor indisponível.")

    proposal = run_f110_proposal(
        proposal_payload["environment"],
        proposal_payload["operation_type"],
        company_code=proposal_payload["company_code"],
        payment_method=proposal_payload["payment_method"],
        account_number=proposal_payload["account_number"],
        posting_date=proposal_payload["posting_date"],
        next_due_date=proposal_payload["next_due_date"],
        document_number=document_number,
    )

    payment_input = build_f110_proposal_payload(
        {
            "source_payload": proposal.payload,
            "document_number": document_number,
        }
    )
    payment = run_f110_payment(
        payment_input["environment"],
        payment_input["operation_type"],
        company_code=payment_input["company_code"],
        payment_method=payment_input["payment_method"],
        account_number=payment_input["account_number"],
        posting_date=payment_input["posting_date"],
        next_due_date=payment_input["next_due_date"],
        document_number=payment_input["document_number"],
        run_id=proposal.run_id,
    )

    print(_ascii({
        "document_number": document_number,
        "proposal": {
            "ok": proposal.ok,
            "status": proposal.status,
            "run_id": proposal.run_id,
            "job_name": proposal.job_name,
            "job_count": proposal.job_count,
            "job_status": proposal.job_status,
            "message": proposal.message,
            "included": proposal.document_included_in_proposal,
        },
        "payment": {
            "ok": payment.ok,
            "status": payment.status,
            "run_id": payment.run_id,
            "job_name": payment.job_name,
            "job_count": payment.job_count,
            "job_status": payment.job_status,
            "message": payment.message,
            "included": payment.document_included_in_proposal,
        },
    }))

    return 0 if proposal.ok and payment.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())

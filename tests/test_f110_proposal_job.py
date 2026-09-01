from __future__ import annotations

from unittest import TestCase

from sap_script_web_cockpit_v2.worker.f110_proposal_job import build_f110_proposal_payload


class F110ProposalJobTests(TestCase):
    def test_build_f110_proposal_payload_merges_source_payload(self) -> None:
        payload = build_f110_proposal_payload(
            {
                "source_payload": {
                    "environment": "QAD",
                    "operation_type": "pagamento",
                    "company_code": "2010",
                    "payment_method": "S",
                    "account_number": "0010000040",
                    "posting_date": "2026-09-01",
                    "next_due_date": "2026-09-02",
                    "document_number": "6050000074",
                },
                "document_number": "6050000074",
            }
        )

        self.assertEqual(payload["environment"], "QAD")
        self.assertEqual(payload["operation_type"], "pagamento")
        self.assertEqual(payload["company_code"], "2010")
        self.assertEqual(payload["payment_method"], "S")
        self.assertEqual(payload["account_number"], "0010000040")
        self.assertEqual(payload["posting_date"], "2026-09-01")
        self.assertEqual(payload["next_due_date"], "2026-09-02")
        self.assertEqual(payload["document_number"], "6050000074")

from __future__ import annotations

import dataclasses
import os
from dataclasses import dataclass
from unittest import TestCase, mock

from sap_script_web_cockpit_v2.worker.fi_default_document_job import (
    build_fi_default_document_payload,
    run_fi_default_document_job,
    update_job_params_via_api,
)


@dataclass
class _Result:
    status: str
    message: str
    ok: bool = True


class FiDefaultDocumentJobTests(TestCase):
    def test_build_defaults(self) -> None:
        environment, branch, payload = build_fi_default_document_payload({})

        self.assertEqual(environment, "QAD")
        self.assertEqual(branch, "cliente")
        self.assertEqual(payload["data_mode"], "default")
        self.assertEqual(payload["environment"], "QAD")
        self.assertEqual(payload["branch"], "cliente")

    @mock.patch("sap_script_web_cockpit_v2.worker.fi_default_document_job.requests.post")
    def test_update_job_params_via_api_uses_api_payload(self, post_mock: mock.Mock) -> None:
        response = mock.Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = {"ok": True}
        post_mock.return_value = response

        with mock.patch.dict(os.environ, {"API_BASE_URL": "http://api.local", "WORKER_TOKEN": "tok"}, clear=False):
            result = update_job_params_via_api("job-1", {"fi_document_result": {"status": "SUCESSO"}})

        self.assertEqual(result, {"ok": True})
        post_mock.assert_called_once()
        args, kwargs = post_mock.call_args
        self.assertEqual(args[0], "http://api.local/api/jobs/job-1/params")
        self.assertEqual(kwargs["headers"], {"X-Worker-Token": "tok"})
        self.assertEqual(kwargs["json"], {"params": {"fi_document_result": {"status": "SUCESSO"}}})

    def test_run_job_updates_and_formats_log(self) -> None:
        fake_result = _Result(status="SUCESSO", message="Documento lançado")

        with mock.patch(
            "sap_script_web_cockpit_v2.worker.fi_default_document_job.update_job_params_via_api"
        ) as update_mock:
            result_json, log = run_fi_default_document_job(
                job_id="job-2",
                params={"environment": "QAD", "branch": "cliente", "payload": {"data_mode": "default"}},
                post_fi_document=mock.Mock(return_value=fake_result),
            )

        update_mock.assert_called_once()
        self.assertIn('"status": "SUCESSO"', result_json)
        self.assertIn("Ambiente: QAD", log)
        self.assertIn("Branch: cliente", log)


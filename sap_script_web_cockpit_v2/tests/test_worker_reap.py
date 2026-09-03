"""
Testes da limpeza de jobs orfaos (Fase 4).

`reap_orphan_running_jobs` / `POST /api/worker/jobs/reap-orphans`: no arranque,
o worker marca como failed os jobs presos em 'running' com o seu nome.

    python -m unittest tests.test_worker_reap
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

_COCKPIT_DIR = Path(__file__).resolve().parents[1]
_REPO_ROOT = Path(__file__).resolve().parents[2]
for _p in (str(_COCKPIT_DIR), str(_REPO_ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

_TMPDIR = tempfile.mkdtemp(prefix="worker_reap_")
os.environ["DATA_DIR"] = _TMPDIR

import web_api.store as store  # noqa: E402
from web_api import main  # noqa: E402


def _body(response) -> dict:
    return json.loads(bytes(response.body).decode("utf-8"))


class WorkerReapTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        store.DATA_DIR = Path(_TMPDIR)
        store.DB_PATH = store.DATA_DIR / "sap_script_jobs.sqlite3"
        store.init_db()

    def _running_job(self, worker: str) -> str:
        """Cria um job e forca-o a 'running' com este worker (deterministico,
        sem depender da ordem de claim_next_job)."""
        job = store.create_job("select_excel_file", {})
        with store.get_connection() as conn:
            conn.execute(
                "UPDATE jobs SET state = 'running', worker_name = ? WHERE id = ?",
                (worker, job["id"]),
            )
            conn.commit()
        row = store.get_job(job["id"])
        self.assertEqual(row["state"], "running")
        self.assertEqual(row["worker_name"], worker)
        return job["id"]

    def test_reap_marks_only_this_worker_running_jobs(self) -> None:
        mine = self._running_job("WORKER_A")
        other = self._running_job("WORKER_B")

        reaped = store.reap_orphan_running_jobs("WORKER_A")

        self.assertEqual(reaped, [mine])
        self.assertEqual(store.get_job(mine)["state"], "failed")
        self.assertIn("orfao", store.get_job(mine)["status"].lower())
        # o job de outro worker fica intacto
        self.assertEqual(store.get_job(other)["state"], "running")

    def test_reap_ignores_pending_and_finished(self) -> None:
        pending_id = store.create_job("select_excel_file", {})["id"]  # fica pending
        done = self._running_job("WORKER_C")
        store.complete_job(done, "succeeded", "ok", "")

        reaped = store.reap_orphan_running_jobs("WORKER_C")

        self.assertEqual(reaped, [])
        self.assertEqual(store.get_job(pending_id)["state"], "pending")
        self.assertEqual(store.get_job(done)["state"], "succeeded")

    def test_reap_empty_worker_name_is_noop(self) -> None:
        self.assertEqual(store.reap_orphan_running_jobs(""), [])
        self.assertEqual(store.reap_orphan_running_jobs("   "), [])

    def test_endpoint_requires_worker_token(self) -> None:
        from fastapi import HTTPException

        with self.assertRaises(HTTPException) as ctx:
            main.api_worker_reap_orphans(worker_name="WORKER_X", x_worker_token="errado")
        self.assertEqual(ctx.exception.status_code, 401)

    def test_endpoint_reaps_and_reports(self) -> None:
        token = main.WORKER_TOKEN
        job_id = self._running_job("WORKER_EP")
        out = main.api_worker_reap_orphans(worker_name="WORKER_EP", x_worker_token=token)
        self.assertEqual(out["count"], 1)
        self.assertEqual(out["reaped"], [job_id])
        self.assertEqual(store.get_job(job_id)["state"], "failed")


if __name__ == "__main__":
    unittest.main()

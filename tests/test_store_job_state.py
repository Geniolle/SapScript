from __future__ import annotations

import tempfile
from datetime import datetime, timezone
from pathlib import Path
from unittest import TestCase, mock

from sap_script_web_cockpit_v2.web_api import store


def _ticket(key: str, summary: str = "Ticket", status: str = "Open") -> dict[str, str]:
    now = datetime.now(timezone.utc).isoformat(timespec="seconds")
    return {
        "key": key,
        "summary": summary,
        "status": status,
        "assignee": "",
        "created_at": now,
        "updated_at": now,
        "priority": "Medium",
        "ticket_type": "Bug",
        "creator": "tester",
        "project": "SAP",
        "team": "Platform",
        "stream": "Core",
        "process": "Sync",
        "time_to_resolution": "",
        "supplier": "",
        "linked_keys": [],
        "resolved_at": "",
    }


class StoreJobStateTests(TestCase):
    def setUp(self) -> None:
        self._temp_dir = tempfile.TemporaryDirectory(ignore_cleanup_errors=True)
        self.addCleanup(self._temp_dir.cleanup)
        data_dir = Path(self._temp_dir.name)
        db_path = data_dir / "sap_script_jobs.sqlite3"
        self._patchers = [
            mock.patch.object(store, "DATA_DIR", data_dir),
            mock.patch.object(store, "DB_PATH", db_path),
        ]
        for patcher in self._patchers:
            patcher.start()
            self.addCleanup(patcher.stop)

        store.init_db()

    def test_complete_job_is_idempotent_and_rejects_late_conflicts(self) -> None:
        job = store.create_job("demo", {"foo": "bar"})
        claimed = store.claim_next_job("worker-a")

        self.assertIsNotNone(claimed)
        self.assertEqual(claimed["id"], job["id"])
        self.assertEqual(claimed["state"], "running")

        finished = store.complete_job(job["id"], "succeeded", "OK", "primeiro log")
        self.assertEqual(finished["state"], "succeeded")
        self.assertEqual(finished["status"], "OK")

        duplicate = store.complete_job(job["id"], "succeeded", "OK", "segundo log")
        self.assertEqual(duplicate["state"], "succeeded")
        self.assertEqual(duplicate["status"], "OK")
        self.assertEqual(duplicate["log"], finished["log"])

    def test_cancel_job_is_idempotent_and_blocks_late_completion(self) -> None:
        job = store.create_job("demo", {"foo": "bar"})
        store.claim_next_job("worker-b")

        cancelled = store.cancel_job(job["id"])
        self.assertEqual(cancelled["state"], "failed")
        self.assertEqual(cancelled["status"], "Cancelado pelo utilizador")

        duplicate_cancel = store.cancel_job(job["id"])
        self.assertEqual(duplicate_cancel["state"], "failed")
        self.assertEqual(duplicate_cancel["status"], "Cancelado pelo utilizador")

        with self.assertRaises(ValueError):
            store.complete_job(job["id"], "succeeded", "OK", "log tardio")

    def test_save_jira_tickets_does_not_prune_by_default(self) -> None:
        store.save_jira_tickets_to_db([_ticket("JIRA-1"), _ticket("JIRA-2")])

        store.save_jira_tickets_to_db([_ticket("JIRA-2")])

        with store.get_connection() as conn:
            rows = conn.execute(
                "SELECT key FROM jira_tickets ORDER BY key"
            ).fetchall()
        self.assertEqual([row["key"] for row in rows], ["JIRA-1", "JIRA-2"])

        store.save_jira_tickets_to_db([_ticket("JIRA-2")], prune_missing=True)

        with store.get_connection() as conn:
            rows = conn.execute(
                "SELECT key FROM jira_tickets ORDER BY key"
            ).fetchall()
        self.assertEqual([row["key"] for row in rows], ["JIRA-2"])

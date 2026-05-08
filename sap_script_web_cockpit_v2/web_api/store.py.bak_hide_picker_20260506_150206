import json
import os
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DATA_DIR = Path(os.getenv("DATA_DIR", "/data"))
DB_PATH = DATA_DIR / "sap_script_jobs.sqlite3"


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def get_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                id TEXT PRIMARY KEY,
                task TEXT NOT NULL,
                params_json TEXT NOT NULL,
                state TEXT NOT NULL,
                status TEXT NOT NULL,
                log TEXT NOT NULL,
                worker_name TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.commit()


def row_to_job(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "task": row["task"],
        "params": json.loads(row["params_json"] or "{}"),
        "state": row["state"],
        "status": row["status"],
        "log": row["log"],
        "worker_name": row["worker_name"],
        "created_at": row["created_at"],
        "updated_at": row["updated_at"],
    }


def create_job(task: str, params: dict[str, Any]) -> dict[str, Any]:
    job_id = str(uuid.uuid4())
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO jobs (id, task, params_json, state, status, log, worker_name, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                job_id,
                task,
                json.dumps(params, ensure_ascii=False),
                "pending",
                "A aguardar execução",
                "",
                "",
                now,
                now,
            ),
        )
        conn.commit()
    job = get_job(job_id)
    if not job:
        raise RuntimeError("Job criado mas não encontrado na base de dados.")
    return job


def get_job(job_id: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
    return row_to_job(row) if row else None


def list_jobs(limit: int = 50) -> list[dict[str, Any]]:
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?",
            (int(limit),),
        ).fetchall()
    return [row_to_job(row) for row in rows]


def claim_next_job(worker_name: str) -> dict[str, Any] | None:
    now = utc_now()
    with get_connection() as conn:
        conn.execute("BEGIN IMMEDIATE")
        row = conn.execute(
            "SELECT * FROM jobs WHERE state = 'pending' ORDER BY created_at ASC LIMIT 1"
        ).fetchone()
        if not row:
            conn.commit()
            return None
        conn.execute(
            """
            UPDATE jobs
            SET state = 'running', status = ?, worker_name = ?, updated_at = ?
            WHERE id = ?
            """,
            ("Em execução no worker Windows", worker_name, now, row["id"]),
        )
        conn.commit()
    return get_job(row["id"])


def complete_job(job_id: str, state: str, status: str, log: str) -> dict[str, Any]:
    if state not in {"succeeded", "failed"}:
        raise ValueError("Estado final inválido.")
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE jobs
            SET state = ?, status = ?, log = ?, updated_at = ?
            WHERE id = ?
            """,
            (state, status, log, now, job_id),
        )
        conn.commit()
    job = get_job(job_id)
    if not job:
        raise RuntimeError("Job concluído mas não encontrado na base de dados.")
    return job

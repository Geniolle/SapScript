import json
import os
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

DATA_DIR = Path(os.getenv("DATA_DIR", "/data"))
DB_PATH = DATA_DIR / "sap_script_jobs.sqlite3"

INTERNAL_TASKS = {
    "select_excel_file",
    "sap_search_requests",
}


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
        try:
            conn.execute("ALTER TABLE jobs ADD COLUMN is_archived INTEGER DEFAULT 0")
        except sqlite3.OperationalError:
            pass
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
        "is_archived": bool(row["is_archived"]),
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


def list_jobs(limit: int = 50, include_internal: bool = False) -> list[dict[str, Any]]:
    """
    Lista a fila/histórico visível.

    Tasks técnicas, como select_excel_file, são usadas apenas para comunicar
    com o worker Windows e não devem aparecer visualmente na Fila / Histórico.
    """
    with get_connection() as conn:
        if include_internal:
            rows = conn.execute(
                "SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?",
                (int(limit),),
            ).fetchall()
        else:
            placeholders = ",".join("?" for _ in INTERNAL_TASKS)
            rows = conn.execute(
                f"""
                SELECT *
                FROM jobs
                WHERE task NOT IN ({placeholders})
                ORDER BY created_at DESC
                LIMIT ?
                """,
                (*INTERNAL_TASKS, int(limit)),
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
    log = (log or "").strip()

    with get_connection() as conn:
        row = conn.execute("SELECT log, task FROM jobs WHERE id = ?", (job_id,)).fetchone()
        if row:
            current_log = (row["log"] or "").strip()
            task = row["task"]
            
            if current_log:
                if task == "sap_cockpit":
                    # Se for sap_cockpit e terminou com sucesso, o log do streaming já está completo.
                    # Só adicionamos se o novo log for um erro/traceback que não esteja no log atual.
                    if state == "failed" and log and log not in current_log:
                        new_log = current_log + "\n\n" + log
                    else:
                        new_log = current_log
                else:
                    # Para outras tarefas, se o log enviado já começa ou contém o atual,
                    # usamos o novo log completo, senão fazemos append.
                    if log:
                        if log.startswith(current_log) or current_log in log:
                            new_log = log
                        else:
                            new_log = current_log + "\n\n" + log
                    else:
                        new_log = current_log
            else:
                new_log = log
        else:
            new_log = log

        conn.execute(
            """
            UPDATE jobs
            SET state = ?, status = ?, log = ?, updated_at = ?
            WHERE id = ?
            """,
            (state, status, new_log, now, job_id),
        )
        conn.commit()

    job = get_job(job_id)

    if not job:
        raise RuntimeError("Job concluído mas não encontrado na base de dados.")

    return job

def cancel_job(job_id: str) -> dict[str, Any]:
    now = utc_now()

    with get_connection() as conn:
        conn.execute(
            """
            UPDATE jobs
            SET state = 'failed', status = 'Cancelado pelo utilizador', log = 'O pedido foi cancelado manualmente via interface web.', updated_at = ?
            WHERE id = ? AND state IN ('pending', 'running')
            """,
            (now, job_id),
        )
        conn.commit()

    job = get_job(job_id)

    if not job:
        raise RuntimeError("Job cancelado mas não encontrado na base de dados.")

    return job

def append_job_log(job_id: str, log_line: str) -> dict[str, Any]:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE jobs
            SET log = log || '\n' || ?, updated_at = ?
            WHERE id = ?
            """,
            (log_line, now, job_id),
        )
        conn.commit()

    job = get_job(job_id)
    if not job:
        raise RuntimeError("Job não encontrado para append log.")
    return job


def archive_job(job_id: str) -> dict[str, Any]:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE jobs
            SET is_archived = 1, updated_at = ?
            WHERE id = ?
            """,
            (now, job_id),
        )
        conn.commit()

    job = get_job(job_id)
    if not job:
        raise RuntimeError("Job arquivado mas não encontrado na base de dados.")
    return job


def update_job_params(job_id: str, new_params: dict[str, Any]) -> dict[str, Any]:
    now = utc_now()
    with get_connection() as conn:
        row = conn.execute("SELECT params_json FROM jobs WHERE id = ?", (job_id,)).fetchone()
        if not row:
            raise RuntimeError("Job não encontrado para atualizar params.")
        params = json.loads(row["params_json"] or "{}")
        params.update(new_params)
        conn.execute(
            "UPDATE jobs SET params_json = ?, updated_at = ? WHERE id = ?",
            (json.dumps(params, ensure_ascii=False), now, job_id),
        )
        conn.commit()

    job = get_job(job_id)
    if not job:
        raise RuntimeError("Job atualizado mas não encontrado na base de dados.")
    return job

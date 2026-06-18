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
    "sap_agent_analysis",
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

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS jira_tickets (
                key TEXT PRIMARY KEY,
                summary TEXT NOT NULL,
                status TEXT NOT NULL,
                assignee TEXT,
                created_at TEXT,
                updated_at TEXT,
                last_sync_at TEXT NOT NULL,
                priority TEXT,
                ticket_type TEXT,
                creator TEXT
            )
            """
        )
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN priority TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN ticket_type TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN creator TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN project TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN team TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN stream TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN process TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN time_to_resolution TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN supplier TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN linked_keys TEXT")
        except sqlite3.OperationalError:
            pass
        try:
            conn.execute("ALTER TABLE jira_tickets ADD COLUMN resolved_at TEXT")
        except sqlite3.OperationalError:
            pass

        # Index creation for JIRA tickets performance
        try:
            conn.execute("CREATE INDEX IF NOT EXISTS idx_jira_tickets_status ON jira_tickets(status)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_jira_tickets_updated_at ON jira_tickets(updated_at)")
        except sqlite3.OperationalError:
            pass

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS jira_auto_trigger_log (
                id TEXT PRIMARY KEY,
                triggered_at TEXT NOT NULL,
                ticket_key TEXT NOT NULL,
                ticket_summary TEXT,
                job_id TEXT,
                status TEXT NOT NULL,
                reason TEXT
            )
            """
        )

        # ---------------------------------------------------------------------------
        # Agent Context Rules — tabela de parâmetros de contexto para o Agente SAP
        # ---------------------------------------------------------------------------
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS agent_context_rules (
                id TEXT PRIMARY KEY,
                campo TEXT NOT NULL,
                valor TEXT NOT NULL,
                transacao_sap TEXT,
                notas TEXT,
                tags TEXT,
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


def list_jobs(limit: int = 50, include_internal: bool = False, include_archived: bool = False) -> list[dict[str, Any]]:
    """
    Lista a fila/histórico visível.

    Tasks técnicas, como select_excel_file, são usadas apenas para comunicar
    com o worker Windows e não devem aparecer visualmente na Fila / Histórico.
    """
    with get_connection() as conn:
        if include_internal:
            if include_archived:
                rows = conn.execute(
                    "SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?",
                    (int(limit),),
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM jobs WHERE is_archived = 0 ORDER BY created_at DESC LIMIT ?",
                    (int(limit),),
                ).fetchall()
        else:
            placeholders = ",".join("?" for _ in INTERNAL_TASKS)
            if include_archived:
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
            else:
                rows = conn.execute(
                    f"""
                    SELECT *
                    FROM jobs
                    WHERE task NOT IN ({placeholders}) AND is_archived = 0
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

def unarchive_job(job_id: str) -> dict[str, Any]:
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE jobs
            SET is_archived = 0, updated_at = ?
            WHERE id = ?
            """,
            (now, job_id),
        )
        conn.commit()

    job = get_job(job_id)
    if not job:
        raise RuntimeError("Job desarquivado mas não encontrado na base de dados.")
    return job

def delete_job(job_id: str) -> None:
    with get_connection() as conn:
        conn.execute("DELETE FROM jobs WHERE id = ?", (job_id,))
        conn.commit()


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


def save_jira_ticket_batch_only(tickets: list[dict[str, Any]]) -> None:
    now = utc_now()
    with get_connection() as conn:
        for t in tickets:
            conn.execute(
                """
                INSERT INTO jira_tickets (key, summary, status, assignee, created_at, updated_at, last_sync_at, priority, ticket_type, creator, project, team, stream, process, time_to_resolution, supplier, linked_keys, resolved_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    summary = excluded.summary,
                    status = excluded.status,
                    assignee = excluded.assignee,
                    created_at = excluded.created_at,
                    updated_at = excluded.updated_at,
                    last_sync_at = excluded.last_sync_at,
                    priority = excluded.priority,
                    ticket_type = excluded.ticket_type,
                    creator = excluded.creator,
                    project = excluded.project,
                    team = excluded.team,
                    stream = excluded.stream,
                    process = excluded.process,
                    time_to_resolution = excluded.time_to_resolution,
                    supplier = excluded.supplier,
                    linked_keys = excluded.linked_keys,
                    resolved_at = excluded.resolved_at
                """,
                (
                    t["key"],
                    t["summary"],
                    t["status"],
                    t["assignee"],
                    t["created_at"],
                    t["updated_at"],
                    now,
                    t.get("priority"),
                    t.get("ticket_type"),
                    t.get("creator"),
                    t.get("project"),
                    t.get("team"),
                    t.get("stream"),
                    t.get("process"),
                    t.get("time_to_resolution"),
                    t.get("supplier"),
                    json.dumps(t.get("linked_keys", [])),
                    t.get("resolved_at"),
                ),
            )
        conn.commit()


def save_jira_tickets_to_db(tickets: list[dict[str, Any]]) -> None:
    now = utc_now()
    active_keys = [t["key"] for t in tickets]

    with get_connection() as conn:
        for t in tickets:
            conn.execute(
                """
                INSERT INTO jira_tickets (key, summary, status, assignee, created_at, updated_at, last_sync_at, priority, ticket_type, creator, project, team, stream, process, time_to_resolution, supplier, linked_keys, resolved_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(key) DO UPDATE SET
                    summary = excluded.summary,
                    status = excluded.status,
                    assignee = excluded.assignee,
                    created_at = excluded.created_at,
                    updated_at = excluded.updated_at,
                    last_sync_at = excluded.last_sync_at,
                    priority = excluded.priority,
                    ticket_type = excluded.ticket_type,
                    creator = excluded.creator,
                    project = excluded.project,
                    team = excluded.team,
                    stream = excluded.stream,
                    process = excluded.process,
                    time_to_resolution = excluded.time_to_resolution,
                    supplier = excluded.supplier,
                    linked_keys = excluded.linked_keys,
                    resolved_at = excluded.resolved_at
                """,
                (
                    t["key"],
                    t["summary"],
                    t["status"],
                    t["assignee"],
                    t["created_at"],
                    t["updated_at"],
                    now,
                    t.get("priority"),
                    t.get("ticket_type"),
                    t.get("creator"),
                    t.get("project"),
                    t.get("team"),
                    t.get("stream"),
                    t.get("process"),
                    t.get("time_to_resolution"),
                    t.get("supplier"),
                    json.dumps(t.get("linked_keys", [])),
                    t.get("resolved_at"),
                ),
            )

        # Only delete OPEN tickets that are not in the active sync set
        if active_keys:
            placeholders = ",".join("?" for _ in active_keys)
            conn.execute(
                f"""
                DELETE FROM jira_tickets 
                WHERE key NOT IN ({placeholders})
                  AND lower(status) NOT IN ('done', 'closed', 'concluído', 'resolvido', 'fechado', 'fechada', 'cancelled')
                """,
                (*active_keys,),
            )
        else:
            conn.execute(
                """
                DELETE FROM jira_tickets 
                WHERE lower(status) NOT IN ('done', 'closed', 'concluído', 'resolvido', 'fechado', 'fechada', 'cancelled')
                """
            )
        conn.commit()


def list_jira_tickets(limit: int = 50, exclude_closed: bool = True) -> list[dict[str, Any]]:
    with get_connection() as conn:
        if exclude_closed:
            query = """
                SELECT * FROM jira_tickets
                WHERE lower(status) NOT IN ('done', 'concluído', 'resolvido', 'fechada', 'closed', 'cancelled', 'fechado')
                ORDER BY updated_at DESC
                LIMIT ?
            """
        else:
            query = """
                SELECT * FROM jira_tickets
                ORDER BY
                    CASE WHEN lower(status) IN ('done', 'concluído', 'resolvido', 'fechada', 'closed', 'cancelled', 'fechado') THEN 1 ELSE 0 END,
                    updated_at DESC
                LIMIT ?
            """
        rows = conn.execute(query, (limit,)).fetchall()
    return [
        {
            "key": row["key"],
            "summary": row["summary"],
            "status": row["status"],
            "assignee": row["assignee"],
            "created_at": row["created_at"],
            "updated_at": row["updated_at"],
            "last_sync_at": row["last_sync_at"],
            "priority": row["priority"],
            "ticket_type": row["ticket_type"],
            "creator": row["creator"],
            "project": row["project"],
            "team": row["team"],
            "stream": row["stream"],
            "process": row["process"],
            "time_to_resolution": row["time_to_resolution"] if "time_to_resolution" in row.keys() else "",
            "supplier": row["supplier"] if "supplier" in row.keys() else "",
            "linked_keys": json.loads(row["linked_keys"]) if "linked_keys" in row.keys() and row["linked_keys"] else [],
            "resolved_at": row["resolved_at"] if "resolved_at" in row.keys() else "",
        }
        for row in rows
    ]


def update_jira_ticket_assignee(key: str, assignee: str) -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE jira_tickets SET assignee = ?, last_sync_at = ? WHERE key = ?",
            (assignee, utc_now(), key),
        )
        conn.commit()


def update_jira_ticket_type_db(key: str, ticket_type: str) -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE jira_tickets SET ticket_type = ?, last_sync_at = ? WHERE key = ?",
            (ticket_type, utc_now(), key),
        )
        conn.commit()


def update_jira_ticket_status_db(key: str, status: str) -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE jira_tickets SET status = ?, last_sync_at = ? WHERE key = ?",
            (status, utc_now(), key),
        )
        conn.commit()


def update_jira_ticket_supplier_db(key: str, supplier: str) -> None:
    with get_connection() as conn:
        conn.execute(
            "UPDATE jira_tickets SET supplier = ?, last_sync_at = ? WHERE key = ?",
            (supplier, utc_now(), key),
        )
        conn.commit()


# ---------------------------------------------------------------------------
# Auto-Trigger Log
# ---------------------------------------------------------------------------

def log_auto_trigger_entry(
    ticket_key: str,
    ticket_summary: str,
    job_id: str | None,
    status: str,
    reason: str = "",
) -> None:
    """Regista uma entrada no log do auto-trigger."""
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO jira_auto_trigger_log
                (id, triggered_at, ticket_key, ticket_summary, job_id, status, reason)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(uuid.uuid4()),
                utc_now(),
                ticket_key,
                ticket_summary,
                job_id,
                status,
                reason,
            ),
        )
        conn.commit()


def list_auto_trigger_log(limit: int = 50) -> list[dict[str, Any]]:
    """Lista as entradas mais recentes do log do auto-trigger."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT * FROM jira_auto_trigger_log
            ORDER BY triggered_at DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [
        {
            "id": row["id"],
            "triggered_at": row["triggered_at"],
            "ticket_key": row["ticket_key"],
            "ticket_summary": row["ticket_summary"],
            "job_id": row["job_id"],
            "status": row["status"],
            "reason": row["reason"],
        }
        for row in rows
    ]


def has_active_job_for_ticket(ticket_key: str, updated_at: str) -> bool:
    """
    Verifica se já existe um job ativo (pending ou running) para o ticket.
    Usa ticket_key + updated_at como chave de idempotência.
    Retorna True se o ticket já foi processado e não deve ser re-acionado.
    """
    # Verifica jobs ativos (pending/running) com o mesmo ticket_key
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT id FROM jobs
            WHERE state IN ('pending', 'running')
              AND params_json LIKE ?
            LIMIT 1
            """,
            (f'%"jira_key": "{ticket_key}"%',),
        ).fetchone()
        if row:
            return True

        # Verifica se já existe entrada de sucesso no log para esta versão do ticket
        row = conn.execute(
            """
            SELECT id FROM jira_auto_trigger_log
            WHERE ticket_key = ?
              AND status = 'triggered'
              AND reason = ?
            LIMIT 1
            """,
            (ticket_key, updated_at),
        ).fetchone()
        return row is not None


def clear_auto_trigger_log() -> None:
    """Limpa todo o histórico de execuções do auto-trigger."""
    with get_connection() as conn:
        conn.execute("DELETE FROM jira_auto_trigger_log")
        conn.commit()


def delete_auto_trigger_log_entry(entry_id: str) -> None:
    """Elimina uma entrada específica do histórico do auto-trigger."""
    with get_connection() as conn:
        conn.execute("DELETE FROM jira_auto_trigger_log WHERE id = ?", (entry_id,))
        conn.commit()


def get_latest_sap_agent_analysis(ticket_key: str) -> dict[str, Any] | None:
    """Retorna o resultado mais recente e com sucesso da análise do Agente SAP para o ticket indicado."""
    with get_connection() as conn:
        row = conn.execute(
            """
            SELECT * FROM jobs
            WHERE task = 'sap_agent_analysis'
              AND state = 'succeeded'
              AND params_json LIKE ?
            ORDER BY created_at DESC
            LIMIT 1
            """,
            (f'%"ticket_key": "{ticket_key}"%',),
        ).fetchone()
        return row_to_job(row) if row else None


# ---------------------------------------------------------------------------
# Agent Context Rules CRUD
# ---------------------------------------------------------------------------

def _row_to_rule(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id":            row["id"],
        "campo":         row["campo"],
        "valor":         row["valor"],
        "transacao_sap": row["transacao_sap"] or "",
        "notas":         row["notas"] or "",
        "tags":          row["tags"] or "",
        "created_at":    row["created_at"],
        "updated_at":    row["updated_at"],
    }


def create_agent_rule(
    campo: str,
    valor: str,
    transacao_sap: str = "",
    notas: str = "",
    tags: str = "",
) -> dict[str, Any]:
    """Cria uma nova regra de contexto para o Agente SAP."""
    rule_id = str(uuid.uuid4())
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO agent_context_rules
                (id, campo, valor, transacao_sap, notas, tags, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (rule_id, campo.strip(), valor.strip(), transacao_sap.strip(),
             notas.strip(), tags.strip(), now, now),
        )
        conn.commit()
    return get_agent_rule(rule_id)


def get_agent_rule(rule_id: str) -> dict[str, Any] | None:
    with get_connection() as conn:
        row = conn.execute(
            "SELECT * FROM agent_context_rules WHERE id = ?", (rule_id,)
        ).fetchone()
    return _row_to_rule(row) if row else None


def list_agent_rules() -> list[dict[str, Any]]:
    """Lista todas as regras de contexto ordenadas por campo e valor."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT * FROM agent_context_rules ORDER BY campo, valor"
        ).fetchall()
    return [_row_to_rule(r) for r in rows]


def update_agent_rule(
    rule_id: str,
    campo: str,
    valor: str,
    transacao_sap: str = "",
    notas: str = "",
    tags: str = "",
) -> dict[str, Any] | None:
    """Actualiza uma regra de contexto existente."""
    now = utc_now()
    with get_connection() as conn:
        conn.execute(
            """
            UPDATE agent_context_rules
            SET campo = ?, valor = ?, transacao_sap = ?, notas = ?, tags = ?, updated_at = ?
            WHERE id = ?
            """,
            (campo.strip(), valor.strip(), transacao_sap.strip(),
             notas.strip(), tags.strip(), now, rule_id),
        )
        conn.commit()
    return get_agent_rule(rule_id)


def delete_agent_rule(rule_id: str) -> None:
    """Elimina uma regra de contexto."""
    with get_connection() as conn:
        conn.execute(
            "DELETE FROM agent_context_rules WHERE id = ?", (rule_id,)
        )
        conn.commit()


def get_agent_rules_for_ticket(
    processo: str = "",
    ticket_type: str = "",
    stream: str = "",
) -> list[dict[str, Any]]:
    """
    Retorna as regras de contexto que correspondem ao ticket.
    Faz match em qualquer combinação de campo+valor que coincida
    com os metadados fornecidos.
    """
    all_rules = list_agent_rules()
    matches: list[dict[str, Any]] = []

    field_map = {
        "IT SALSA - Categoria SAP": processo,
        "Tipo de Ticket":           ticket_type,
        "Stream":                   stream,
    }

    for rule in all_rules:
        field_value = field_map.get(rule["campo"], "")
        if field_value and rule["valor"].strip().lower() == field_value.strip().lower():
            matches.append(rule)

    return matches

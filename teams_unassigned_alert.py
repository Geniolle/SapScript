from __future__ import annotations

import json
import os
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv

load_dotenv()

BOOL_TRUE = {"1", "true", "yes", "on", "sim", "s"}


def _env_bool(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return str(raw).strip().lower() in BOOL_TRUE


def _env_int(name: str, default: int) -> int:
    try:
        value = int(str(os.getenv(name, str(default))).strip())
        return max(value, 30)
    except Exception:
        return default


def _load_state(state_path: Path) -> dict[str, Any]:
    if not state_path.exists():
        return {"notified": {}}
    try:
        with open(state_path, "r", encoding="utf-8") as file_obj:
            data = json.load(file_obj)
        if isinstance(data, dict):
            data.setdefault("notified", {})
            return data
    except Exception as exc:
        print(f"[TEAMS ALERT] Não foi possível ler estado {state_path}: {exc}")
    return {"notified": {}}


def _save_state(state_path: Path, state: dict[str, Any]) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    with open(state_path, "w", encoding="utf-8") as file_obj:
        json.dump(state, file_obj, ensure_ascii=False, indent=2)


def _jira_config() -> tuple[str, str, str, str]:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    return jira_base, jira_api_path, jira_email, jira_token


def _option_value(value: Any) -> str:
    if isinstance(value, dict):
        return str(value.get("value") or value.get("name") or "").strip()
    if isinstance(value, str):
        return value.strip()
    return ""


def _build_jql() -> str:
    override = os.getenv("TEAMS_UNASSIGNED_JQL", "").strip()
    if override:
        return override

    team_value = os.getenv("TEAMS_UNASSIGNED_TEAM", "Core Systems").strip() or "Core Systems"
    team_jql_field = os.getenv("TEAMS_UNASSIGNED_TEAM_JQL_FIELD", "cf[15839]").strip() or "cf[15839]"
    include_done = _env_bool("TEAMS_UNASSIGNED_INCLUDE_DONE", False)

    clauses = [
        f'{team_jql_field} = "{team_value}"',
        "assignee IS EMPTY",
    ]
    if not include_done:
        clauses.append("statusCategory != Done")
    return " AND ".join(clauses)


def fetch_unassigned_team_tickets() -> list[dict[str, str]]:
    jira_base, jira_api_path, jira_email, jira_token = _jira_config()
    if not jira_base or not jira_email or not jira_token:
        print("[TEAMS ALERT] Credenciais JIRA não configuradas.")
        return []

    team_field_id = os.getenv("TEAMS_UNASSIGNED_TEAM_FIELD_ID", "customfield_15839").strip() or "customfield_15839"
    jira_ui_base = os.getenv("JIRA_UI_BASE_URL", jira_base).strip().rstrip("/") or jira_base
    jql = _build_jql()

    url = f"{jira_base}/{jira_api_path}/search/jql"
    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    payload: dict[str, Any] = {
        "jql": jql,
        "fields": [
            "summary",
            "status",
            "assignee",
            "created",
            "updated",
            "project",
            "priority",
            team_field_id,
        ],
        "maxResults": 100,
    }

    tickets: list[dict[str, str]] = []
    next_page_token = None

    while True:
        if next_page_token:
            payload["nextPageToken"] = next_page_token
        else:
            payload.pop("nextPageToken", None)

        try:
            response = requests.post(url, auth=auth, headers=headers, json=payload, timeout=20)
            response.raise_for_status()
        except Exception as exc:
            print(f"[TEAMS ALERT] Erro ao consultar JIRA: {exc}")
            return tickets

        data = response.json()
        for issue in data.get("issues", []) or []:
            fields = issue.get("fields", {}) or {}
            status_data = fields.get("status") or {}
            project_data = fields.get("project") or {}
            priority_data = fields.get("priority") or {}

            tickets.append(
                {
                    "key": str(issue.get("key") or "").strip().upper(),
                    "summary": str(fields.get("summary") or "").strip(),
                    "status": str(status_data.get("name") or "").strip(),
                    "project": str(project_data.get("name") or project_data.get("key") or "").strip(),
                    "priority": str(priority_data.get("name") or "").strip(),
                    "team": _option_value(fields.get(team_field_id)),
                    "created_at": str(fields.get("created") or "").strip(),
                    "updated_at": str(fields.get("updated") or "").strip(),
                    "url": f"{jira_ui_base}/browse/{str(issue.get('key') or '').strip().upper()}",
                }
            )

        next_page_token = data.get("nextPageToken")
        is_last = data.get("isLast", True)
        if not next_page_token or is_last:
            break

    return tickets


def _ticket_identity(ticket: dict[str, str]) -> str:
    updated_at = str(ticket.get("updated_at") or "").strip()
    return updated_at or str(ticket.get("created_at") or "").strip() or "sem-data"


def _new_or_updated_tickets(tickets: list[dict[str, str]], state: dict[str, Any]) -> list[dict[str, str]]:
    notified = state.setdefault("notified", {})
    result = []
    for ticket in tickets:
        key = ticket.get("key", "")
        if not key:
            continue
        identity = _ticket_identity(ticket)
        if notified.get(key) == identity:
            continue
        result.append(ticket)
    return result


def _format_teams_message(tickets: list[dict[str, str]]) -> str:
    team_name = os.getenv("TEAMS_UNASSIGNED_TEAM", "Core Systems").strip() or "Core Systems"
    now = datetime.now(timezone.utc).astimezone().strftime("%d/%m/%Y %H:%M")

    if len(tickets) == 1:
        title = f"⚠️ Existe 1 ticket da equipa {team_name} sem responsável"
    else:
        title = f"⚠️ Existem {len(tickets)} tickets da equipa {team_name} sem responsável"

    lines = [title, "", f"Verificação: {now}", ""]
    for ticket in tickets:
        details = []
        if ticket.get("status"):
            details.append(f"Estado: {ticket['status']}")
        if ticket.get("priority"):
            details.append(f"Prioridade: {ticket['priority']}")
        if ticket.get("project"):
            details.append(f"Projeto: {ticket['project']}")
        detail_text = " | ".join(details)
        if detail_text:
            detail_text = f" — {detail_text}"
        lines.append(f"- {ticket['key']} — {ticket.get('summary', '')}{detail_text}")
        if ticket.get("url"):
            lines.append(f"  {ticket['url']}")

    lines.append("")
    lines.append("Critério: Team = Core Systems e Responsável vazio.")
    return "\n".join(lines)


def send_teams_message(text: str) -> bool:
    webhook_url = os.getenv("TEAMS_WEBHOOK_URL", "").strip()
    if not webhook_url:
        print("[TEAMS ALERT] TEAMS_WEBHOOK_URL não configurado. Mensagem não enviada.")
        print(text)
        return False

    dry_run = _env_bool("TEAMS_UNASSIGNED_DRY_RUN", False)
    if dry_run:
        print("[TEAMS ALERT] DRY RUN ativo. Mensagem que seria enviada:")
        print(text)
        return True

    try:
        response = requests.post(webhook_url, json={"text": text}, timeout=20)
        if response.status_code not in (200, 202):
            print(f"[TEAMS ALERT] Erro Teams HTTP {response.status_code}: {response.text}")
            return False
        return True
    except Exception as exc:
        print(f"[TEAMS ALERT] Erro ao enviar mensagem Teams: {exc}")
        return False


def run_once() -> dict[str, int]:
    state_path = Path(os.getenv("TEAMS_UNASSIGNED_STATE_PATH", "cache/teams_unassigned_alert_state.json")).resolve()
    state = _load_state(state_path)

    tickets = fetch_unassigned_team_tickets()
    pending = _new_or_updated_tickets(tickets, state)

    if not pending:
        print(f"[TEAMS ALERT] Nenhum ticket novo/alterado sem responsável. Total atual encontrado: {len(tickets)}")
        return {"found": len(tickets), "notified": 0}

    message = _format_teams_message(pending)
    sent = send_teams_message(message)
    if sent:
        notified = state.setdefault("notified", {})
        for ticket in pending:
            notified[ticket["key"]] = _ticket_identity(ticket)
        state["last_run_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        state["last_found_count"] = len(tickets)
        state["last_notified_count"] = len(pending)
        _save_state(state_path, state)
        print(f"[TEAMS ALERT] Alerta enviado para {len(pending)} ticket(s).")
    else:
        print("[TEAMS ALERT] Alerta não enviado; estado não foi atualizado.")

    return {"found": len(tickets), "notified": len(pending) if sent else 0}


def main() -> None:
    if not _env_bool("TEAMS_UNASSIGNED_ALERT_ENABLED", True):
        print("[TEAMS ALERT] Monitor desativado por TEAMS_UNASSIGNED_ALERT_ENABLED=false.")
        return

    run_once_only = _env_bool("TEAMS_UNASSIGNED_RUN_ONCE", False)
    poll_seconds = _env_int("TEAMS_UNASSIGNED_POLL_SECONDS", 300)

    if run_once_only:
        run_once()
        return

    print(f"[TEAMS ALERT] Monitor iniciado. Intervalo: {poll_seconds}s")
    while True:
        try:
            run_once()
        except Exception as exc:
            print(f"[TEAMS ALERT] Erro inesperado no ciclo: {exc}")
        time.sleep(poll_seconds)


if __name__ == "__main__":
    main()

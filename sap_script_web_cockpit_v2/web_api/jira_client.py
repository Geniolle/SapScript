import os
import re
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv()


def _safe_filename(filename: str) -> str:
    """Sanitiza o nome do ficheiro removendo caracteres inválidos."""
    sanitized = re.sub(r'[<>:"/\\|?*]+', "_", str(filename or "").strip())
    return sanitized or "anexo_sem_nome"


def download_ticket_attachments_to_dir(
    ticket_key: str,
    output_dir_container: str,
    output_dir_windows: str,
    only_xlsx: bool = True,
    overwrite: bool = False,
) -> list[str]:
    """
    Descarrega os anexos de um ticket JIRA para uma pasta local.

    Args:
        ticket_key:           Chave do ticket (ex: 'IZ-56680')
        output_dir_container: Caminho base dentro do container (ex: '/data/jira')
        output_dir_windows:   Caminho base no Windows para o worker (ex: r'C:\\Jira')
        only_xlsx:            Se True, descarrega apenas ficheiros .xlsx
        overwrite:            Se True, sobrescreve ficheiros já existentes

    Returns:
        Lista de paths Windows dos ficheiros descarregados (.xlsx mais recente primeiro).
    """
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print(f"[DOWNLOAD] Credenciais JIRA não configuradas para {ticket_key}.")
        return []

    auth = (jira_email, jira_token)
    headers = {"Accept": "application/json"}

    # 1. Criar pasta no container: /data/jira/{TICKET_KEY}/
    ticket_key_upper = ticket_key.strip().upper()
    issue_folder = Path(output_dir_container) / ticket_key_upper
    issue_folder.mkdir(parents=True, exist_ok=True)

    # 2. Buscar anexos do ticket
    url = f"{jira_base}/{jira_api_path}/issue/{ticket_key_upper}"
    try:
        res = requests.get(url, params={"fields": "attachment"}, auth=auth,
                           headers=headers, timeout=30)
        res.raise_for_status()
    except Exception as e:
        print(f"[DOWNLOAD] Erro ao obter anexos de {ticket_key_upper}: {e}")
        return []

    attachments = res.json().get("fields", {}).get("attachment", []) or []
    if not attachments:
        print(f"[DOWNLOAD] {ticket_key_upper}: sem anexos.")
        return []

    downloaded_windows = []

    for att in attachments:
        filename = _safe_filename(att.get("filename", "anexo"))
        if only_xlsx and not filename.lower().endswith(".xlsx"):
            continue

        content_url = att.get("content")
        if not content_url:
            continue

        target = issue_folder / filename
        if target.exists() and not overwrite:
            print(f"[DOWNLOAD] {ticket_key_upper}: já existe -> {filename}")
            # Mesmo que já exista, inclui na lista
            win_path = str(Path(output_dir_windows) / ticket_key_upper / filename)
            downloaded_windows.append(win_path)
            continue

        try:
            with requests.get(content_url, auth=auth, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(target, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
            print(f"[DOWNLOAD] {ticket_key_upper}: descarregado -> {filename}")
            win_path = str(Path(output_dir_windows) / ticket_key_upper / filename)
            downloaded_windows.append(win_path)
        except Exception as e:
            print(f"[DOWNLOAD] {ticket_key_upper}: erro ao descarregar {filename}: {e}")

    # Ordena por data de modificação (mais recente primeiro) dentro do container
    def _mtime(win_path: str) -> float:
        container_path = issue_folder / Path(win_path).name
        try:
            return container_path.stat().st_mtime
        except Exception:
            return 0.0

    downloaded_windows.sort(key=_mtime, reverse=True)
    return downloaded_windows


def _parse_issue(issue: dict) -> dict:
    """Converte um objeto issue da API JIRA num dict normalizado."""
    fields = issue.get("fields", {})
    assignee_data = fields.get("assignee") or {}
    raw_assignee = (
        assignee_data.get("displayName")
        or assignee_data.get("emailAddress")
        or ""
    )
    assignee_parts = raw_assignee.strip().split()
    if len(assignee_parts) > 2:
        assignee_name = f"{assignee_parts[0]} {assignee_parts[-1]}"
    else:
        assignee_name = raw_assignee

    status_data = fields.get("status") or {}
    status_name = status_data.get("name") or "Unknown"

    reporter_data = fields.get("reporter") or {}
    raw_creator = reporter_data.get("displayName") or ""
    creator_parts = raw_creator.strip().split()
    if len(creator_parts) > 2:
        creator_name = f"{creator_parts[0]} {creator_parts[-1]}"
    else:
        creator_name = raw_creator

    # Project field
    project_data = fields.get("project") or {}
    project_name = project_data.get("name") or ""

    # Priority option field
    priority_data = fields.get("customfield_15815")
    priority_name = ""
    if isinstance(priority_data, dict):
        priority_name = priority_data.get("value") or ""
    elif isinstance(priority_data, str):
        priority_name = priority_data

    # Ticket type option field
    type_data = fields.get("customfield_15810")
    type_name = ""
    if isinstance(type_data, dict):
        type_name = type_data.get("value") or ""
    elif isinstance(type_data, str):
        type_name = type_data

    # Team field (customfield_15839)
    team_data = fields.get("customfield_15839")
    team_name = ""
    if isinstance(team_data, dict):
        team_name = team_data.get("value") or ""
    elif isinstance(team_data, str):
        team_name = team_data

    # Stream field (customfield_15260)
    stream_data = fields.get("customfield_15260")
    stream_name = ""
    if isinstance(stream_data, dict):
        stream_name = stream_data.get("value") or ""
    elif isinstance(stream_data, str):
        stream_name = stream_data

    # Process field (customfield_15845)
    process_data = fields.get("customfield_15845")
    process_name = ""
    if isinstance(process_data, dict):
        process_name = process_data.get("value") or ""
    elif isinstance(process_data, str):
        process_name = process_data

    # Time to resolution field (customfield_14560)
    sla_data = fields.get("customfield_14560")
    time_to_resolution = ""
    if isinstance(sla_data, dict):
        ongoing = sla_data.get("ongoingCycle")
        if isinstance(ongoing, dict):
            breached = ongoing.get("breached", False)
            rem = ongoing.get("remainingTime")
            friendly = ""
            if isinstance(rem, dict):
                friendly = rem.get("friendly", "")
            if breached:
                if friendly:
                    time_to_resolution = friendly if friendly.startswith("-") else f"-{friendly}"
                else:
                    time_to_resolution = "Excedido"
            else:
                time_to_resolution = friendly or "Pendente"
        else:
            completed = sla_data.get("completedCycles")
            if isinstance(completed, list) and len(completed) > 0:
                last_cycle = completed[-1]
                if isinstance(last_cycle, dict):
                    breached = last_cycle.get("breached", False)
                    elapsed = last_cycle.get("elapsedTime")
                    friendly_elapsed = ""
                    if isinstance(elapsed, dict):
                        friendly_elapsed = elapsed.get("friendly", "")
                    status = "Resolvido"
                    if breached:
                        status = "Resolvido com atraso"
                    if friendly_elapsed:
                        time_to_resolution = f"{status} ({friendly_elapsed})"
                    else:
                        time_to_resolution = status

    # Supplier option field (customfield_14595)
    supplier_data = fields.get("customfield_14595")
    supplier_name = ""
    if isinstance(supplier_data, dict):
        supplier_name = supplier_data.get("value") or ""
    elif isinstance(supplier_data, str):
        supplier_name = supplier_data

    linked_keys: list[str] = []
    done_statuses = ["DONE", "CONCLU", "RESOLV", "FECHADO", "FECHADA", "CLOSED"]
    
    for link in fields.get("issuelinks", []) or []:
        # "inwardIssue" = ticket que linka ESTE; "outwardIssue" = ticket que ESTE linka
        for direction in ("inwardIssue", "outwardIssue"):
            linked_issue = link.get(direction)
            if linked_issue and linked_issue.get("key"):
                issue_fields = linked_issue.get("fields") or {}
                status_obj = issue_fields.get("status") or {}
                link_status = status_obj.get("name", "").upper()
                
                # Check if the linked issue status is in the done_statuses
                if not any(ds in link_status for ds in done_statuses):
                    linked_keys.append(linked_issue["key"])

    return {
        "key": issue.get("key"),
        "summary": fields.get("summary") or "",
        "status": status_name,
        "assignee": assignee_name,
        "created_at": fields.get("created") or "",
        "updated_at": fields.get("updated") or "",
        "priority": priority_name,
        "ticket_type": type_name,
        "creator": creator_name,
        "project": project_name,
        "team": team_name,
        "stream": stream_name,
        "process": process_name,
        "time_to_resolution": time_to_resolution,
        "supplier": supplier_name,
        "linked_keys": linked_keys,
    }


def _fetch_single_issue(
    key: str,
    jira_base: str,
    jira_api_path: str,
    auth: tuple,
    headers: dict,
) -> dict | None:
    """Busca e parseia um único issue JIRA pelo seu key."""
    url = f"{jira_base}/{jira_api_path}/issue/{key.strip().upper()}"
    params = {
        "fields": ",".join([
            "summary", "status", "assignee", "created", "updated",
            "reporter", "project", "customfield_15815", "customfield_15810",
            "customfield_15839", "customfield_15260", "customfield_15845",
            "customfield_14560", "customfield_14595", "issuelinks",
        ])
    }
    try:
        res = requests.get(url, auth=auth, headers=headers, params=params, timeout=15)
        res.raise_for_status()
        return _parse_issue(res.json())
    except Exception as e:
        print(f"[JIRA SYNC] Erro ao buscar ticket linkado {key}: {e}")
        return None


def fetch_jira_tickets_from_api() -> list[dict]:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print(
            "Jira integration credentials are not configured in environment variables."
        )
        return []

    # Query para buscar tickets abertos atribuídos ao usuário logado
    jql = os.getenv(
        "JIRA_SYNC_JQL", "assignee = currentUser() AND statusCategory != Done"
    )

    url = f"{jira_base}/{jira_api_path}/search/jql"
    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    payload = {
        "jql": jql,
        "fields": [
            "summary",
            "status",
            "assignee",
            "created",
            "updated",
            "reporter",
            "project",
            "customfield_15815",
            "customfield_15810",
            "customfield_15839",
            "customfield_15260",
            "customfield_15845",
            "customfield_14560",
            "customfield_14595",
            "issuelinks",
        ],
        "maxResults": 100,
    }

    tickets = []
    next_page_token = None

    while True:
        if next_page_token:
            payload["nextPageToken"] = next_page_token
        else:
            payload.pop("nextPageToken", None)

        try:
            response = requests.post(
                url, auth=auth, headers=headers, json=payload, timeout=15
            )
            response.raise_for_status()
        except Exception as e:
            print(f"Jira API connection error: {e}")
            if tickets:
                break
            raise

        data = response.json()
        issues = data.get("issues", [])

        for issue in issues:
            parsed_ticket = _parse_issue(issue)
            
            # Regra de exclusão: Não selecionar resultados do projeto IZ com status DONE/Resolvido/Fechada
            if parsed_ticket["key"].upper().startswith("IZ-"):
                status_name = parsed_ticket.get("status", "").upper()
                if any(done_status in status_name for done_status in ["DONE", "CONCLU", "RESOLV", "FECHADO", "FECHADA", "CLOSED"]):
                    continue
                    
            tickets.append(parsed_ticket)

        next_page_token = data.get("nextPageToken")
        is_last = data.get("isLast", True)
        if not next_page_token or is_last:
            break

    # ─────────────────────────────────────────────────────────────────────────

    return tickets


def assign_jira_ticket(key: str, assignee_name: str) -> bool:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print("Jira integration credentials are not configured in environment variables.")
        return False

    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    try:
        # If it's a special string like 'Sem responsável', we want to unassign (set assignee to None)
        if not assignee_name or assignee_name.lower() in ("sem responsável", "unassigned", "none", ""):
            assign_url = f"{jira_base}/{jira_api_path}/issue/{key}/assignee"
            res = requests.put(assign_url, auth=auth, headers=headers, json={"accountId": None}, timeout=15)
            return res.ok

        # Search for assignee_name to find their accountId
        search_url = f"{jira_base}/{jira_api_path}/user/search"
        params = {"query": assignee_name, "maxResults": 1}
        res = requests.get(search_url, auth=auth, headers=headers, params=params, timeout=15)
        res.raise_for_status()
        users = res.json()
        if not users:
            print(f"User {assignee_name} not found in search.")
            return False

        account_id = users[0].get("accountId")
        if not account_id:
            print("No accountId found for user.")
            return False

        # Assign issue
        assign_url = f"{jira_base}/{jira_api_path}/issue/{key}/assignee"
        res = requests.put(assign_url, auth=auth, headers=headers, json={"accountId": account_id}, timeout=15)
        res.raise_for_status()
        return True
    except Exception as e:
        print(f"Error assigning ticket {key} to {assignee_name}: {e}")
        return False


def update_jira_ticket_type(key: str, ticket_type: str) -> bool:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print("Jira integration credentials are not configured in environment variables.")
        return False

    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    url = f"{jira_base}/{jira_api_path}/issue/{key}"
    payload = {
        "fields": {
            "customfield_15810": {
                "value": ticket_type
            } if ticket_type else None
        }
    }

    try:
        res = requests.put(url, auth=auth, headers=headers, json=payload, timeout=15)
        res.raise_for_status()
        return True
    except Exception as e:
        print(f"Error updating ticket type for {key} to {ticket_type}: {e}")
        return False


def get_jira_issue_transitions(key: str) -> list[dict]:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        return []

    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    url = f"{jira_base}/{jira_api_path}/issue/{key}/transitions"
    try:
        res = requests.get(url, auth=auth, headers=headers, timeout=15)
        res.raise_for_status()
        data = res.json()
        transitions = data.get("transitions", [])
        return [{"id": t.get("id"), "name": t.get("name")} for t in transitions]
    except Exception as e:
        print(f"Error fetching transitions for ticket {key}: {e}")
        return []


def transition_jira_issue(key: str, transition_id: str) -> bool:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        return False

    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    url = f"{jira_base}/{jira_api_path}/issue/{key}/transitions"
    payload = {
        "transition": {
            "id": transition_id
        }
    }

    try:
        res = requests.post(url, auth=auth, headers=headers, json=payload, timeout=15)
        res.raise_for_status()
        return True
    except Exception as e:
        print(f"Error transitioning ticket {key} with transition {transition_id}: {e}")
        return False


def fetch_auto_trigger_tickets(
    assignee_name: str = "",
    status_name: str = "In Review",
    supplier_value: str = "Evolutive",
) -> list[dict]:
    """
    Consulta a API do JIRA em busca de tickets elegíveis para auto-trigger SAP.

    Critérios (configuráveis via .env):
      - status    : JIRA_AUTO_TRIGGER_STATUS   (default: "In Review")
      - supplier  : JIRA_AUTO_TRIGGER_SUPPLIER (default: "Evolutive")
      - assignee  : JIRA_AUTO_TRIGGER_ASSIGNEE (default: "Clayton Lopes")
        → filtrado em Python por displayName (pós-fetch)

    Retorna lista de dicts com: key, summary, status, assignee, process,
    supplier, updated_at, priority.
    """
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print("Jira auto-trigger: credenciais não configuradas.")
        return []

    # Parâmetros configuráveis via env
    env_status = os.getenv("JIRA_AUTO_TRIGGER_STATUS", status_name).strip()
    env_supplier = os.getenv("JIRA_AUTO_TRIGGER_SUPPLIER", supplier_value).strip()
    env_assignee = os.getenv("JIRA_AUTO_TRIGGER_ASSIGNEE", assignee_name or "Clayton Lopes").strip()

    # JQL: status + supplier (assignee filtrado em Python)
    jql = f'status = "{env_status}" AND cf[14595] = "{env_supplier}"'

    url = f"{jira_base}/{jira_api_path}/search/jql"
    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    payload = {
        "jql": jql,
        "fields": [
            "summary",
            "status",
            "assignee",
            "updated",
            "reporter",
            "customfield_15845",  # process
            "customfield_14595",  # supplier
            "customfield_15815",  # priority
        ],
        "maxResults": 100,
    }

    tickets = []
    next_page_token = None

    while True:
        if next_page_token:
            payload["nextPageToken"] = next_page_token
        else:
            payload.pop("nextPageToken", None)

        try:
            response = requests.post(
                url, auth=auth, headers=headers, json=payload, timeout=15
            )
            response.raise_for_status()
        except Exception as e:
            print(f"Jira auto-trigger API error: {e}")
            break

        data = response.json()
        issues = data.get("issues", [])

        for issue in issues:
            fields = issue.get("fields", {})

            assignee_data = fields.get("assignee") or {}
            raw_assignee = (
                assignee_data.get("displayName")
                or assignee_data.get("emailAddress")
                or ""
            )

            # Filtro de assignee em Python (comparação case-insensitive)
            if env_assignee and env_assignee.lower() not in raw_assignee.lower():
                continue

            # Abreviar nome se necessário
            assignee_parts = raw_assignee.strip().split()
            if len(assignee_parts) > 2:
                assignee_display = f"{assignee_parts[0]} {assignee_parts[-1]}"
            else:
                assignee_display = raw_assignee

            status_data = fields.get("status") or {}
            status_val = status_data.get("name") or ""

            process_data = fields.get("customfield_15845")
            process_name = ""
            if isinstance(process_data, dict):
                process_name = process_data.get("value") or ""
            elif isinstance(process_data, str):
                process_name = process_data

            supplier_data = fields.get("customfield_14595")
            supplier_name = ""
            if isinstance(supplier_data, dict):
                supplier_name = supplier_data.get("value") or ""
            elif isinstance(supplier_data, str):
                supplier_name = supplier_data

            priority_data = fields.get("customfield_15815")
            priority_name = ""
            if isinstance(priority_data, dict):
                priority_name = priority_data.get("value") or ""
            elif isinstance(priority_data, str):
                priority_name = priority_data

            tickets.append({
                "key": issue.get("key"),
                "summary": fields.get("summary") or "",
                "status": status_val,
                "assignee": assignee_display,
                "process": process_name,
                "supplier": supplier_name,
                "priority": priority_name,
                "updated_at": fields.get("updated") or "",
            })

        next_page_token = data.get("nextPageToken")
        is_last = data.get("isLast", True)
        if not next_page_token or is_last:
            break

    return tickets


def update_jira_ticket_supplier(key: str, supplier: str) -> bool:
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print("Jira integration credentials are not configured in environment variables.")
        return False

    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json"
    }

    url = f"{jira_base}/{jira_api_path}/issue/{key}"
    payload = {
        "fields": {
            "customfield_14595": {
                "value": supplier
            } if supplier else None
        }
    }

    try:
        res = requests.put(url, auth=auth, headers=headers, json=payload, timeout=15)
        res.raise_for_status()
        return True
    except Exception as e:
        print(f"Error updating supplier for {key} to {supplier}: {e}")
        return False


def _parse_jira_adf(value) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        if "text" in value:
            return str(value["text"])
        return "\n".join(_parse_jira_adf(item) for item in value.get("content", []))
    if isinstance(value, list):
        return "\n".join(_parse_jira_adf(item) for item in value)
    return str(value)


def add_jira_comment(key: str, comment_text: str) -> bool:
    """
    Adiciona um comentário público (Reply to customer) a um ticket JIRA.

    Args:
        key:          Chave do ticket (ex: 'IZ-56680')
        comment_text: Texto do comentário em formato plano

    Returns:
        True se bem-sucedido, False caso contrário.
    """
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print(f"[COMMENT] Credenciais JIRA não configuradas para {key}.")
        return False

    auth = (jira_email, jira_token)
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/json",
    }

    url = f"{jira_base}/{jira_api_path}/issue/{key.upper().strip()}/comment"

    # Jira Cloud REST API v3 usa Atlassian Document Format (ADF)
    payload = {
        "body": {
            "type": "doc",
            "version": 1,
            "content": [
                {
                    "type": "paragraph",
                    "content": [
                        {
                            "type": "text",
                            "text": comment_text,
                        }
                    ],
                }
            ],
        }
    }

    try:
        res = requests.post(url, auth=auth, headers=headers, json=payload, timeout=15)
        res.raise_for_status()
        print(f"[COMMENT] Comentário adicionado ao ticket {key}.")
        return True
    except Exception as e:
        print(f"[COMMENT] Erro ao adicionar comentário ao ticket {key}: {e}")
        return False


def fetch_ticket_details(ticket_key: str) -> dict:
    """
    Busca os detalhes de um único ticket (Summary, Description, Comments) do JIRA,
    convertendo campos no formato ADF (Atlassian Document Format) para texto simples.
    """
    jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
    jira_email = os.getenv("JIRA_EMAIL", "").strip()
    jira_token = os.getenv("JIRA_TOKEN", "").strip()
    jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

    if not jira_base or not jira_email or not jira_token:
        print(f"[CHAT DETAILS] Credenciais JIRA não configuradas para {ticket_key}.")
        return {"summary": "", "description": "", "comments": [], "categoria_sap": ""}

    auth = (jira_email, jira_token)
    headers = {"Accept": "application/json"}
    url = f"{jira_base}/{jira_api_path}/issue/{ticket_key.upper().strip()}"

    try:
        res = requests.get(
            url,
            auth=auth,
            headers=headers,
            params={"fields": "summary,description,comment,customfield_15845"},
            timeout=15,
        )
        res.raise_for_status()
        data = res.json()
        fields = data.get("fields", {})

        # Converter descrição ADF para texto simples
        desc_raw = fields.get("description")
        description = _parse_jira_adf(desc_raw)

        # Converter comentários ADF para lista de textos simples
        comments_raw = fields.get("comment", {}).get("comments", [])
        comments = []
        for c in comments_raw:
            author = c.get("author", {}).get("displayName", "User")
            body = _parse_jira_adf(c.get("body"))
            comments.append(f"{author}: {body}")

        # IT SALSA - Categoria SAP (customfield_15845)
        categoria_raw = fields.get("customfield_15845")
        if isinstance(categoria_raw, dict):
            categoria_sap = categoria_raw.get("value") or ""
        elif isinstance(categoria_raw, str):
            categoria_sap = categoria_raw
        else:
            categoria_sap = ""

        return {
            "summary": str(fields.get("summary") or ""),
            "description": description,
            "comments": comments,
            "categoria_sap": categoria_sap,
        }
    except Exception as e:
        print(f"[CHAT DETAILS] Erro ao obter detalhes de {ticket_key}: {e}")
        return {"summary": "", "description": "", "comments": [], "categoria_sap": ""}



import os
import requests
from dotenv import load_dotenv

load_dotenv()


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
            fields = issue.get("fields", {})
            assignee_data = fields.get("assignee") or {}
            assignee_name = (
                assignee_data.get("displayName")
                or assignee_data.get("emailAddress")
                or ""
            )

            status_data = fields.get("status") or {}
            status_name = status_data.get("name") or "Unknown"

            reporter_data = fields.get("reporter") or {}
            creator_name = reporter_data.get("displayName") or ""

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
                        time_to_resolution = f"Excedido (-{friendly})" if friendly else "Excedido"
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

            tickets.append({
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
            })

        next_page_token = data.get("nextPageToken")
        is_last = data.get("isLast", True)
        if not next_page_token or is_last:
            break

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


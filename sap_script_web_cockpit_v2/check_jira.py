import os
import requests
from dotenv import load_dotenv

load_dotenv(dotenv_path="../.env")

jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
jira_email = os.getenv("JIRA_EMAIL", "").strip()
jira_token = os.getenv("JIRA_TOKEN", "").strip()
jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")

auth = (jira_email, jira_token)
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

# 1. Total resolved tickets before 2026
jql1 = '(project = "IT - Salsa Jeans" OR project = "SAP - Desenvolvimento") AND statusCategory = Done AND resolved < "2026-01-01"'
payload1 = {"jql": jql1, "maxResults": 1}

# 2. Total resolved tickets in 2026
jql2 = '(project = "IT - Salsa Jeans" OR project = "SAP - Desenvolvimento") AND statusCategory = Done AND resolved >= "2026-01-01"'
payload2 = {"jql": jql2, "maxResults": 1}

url = f"{jira_base}/{jira_api_path}/search/jql"

def count_jira_tickets(jql):
    payload = {
        "jql": jql,
        "fields": ["id"],
        "maxResults": 100
    }
    count = 0
    next_page_token = None
    while True:
        if next_page_token:
            payload["nextPageToken"] = next_page_token
        else:
            payload.pop("nextPageToken", None)
        try:
            res = requests.post(url, auth=auth, headers=headers, json=payload, timeout=15)
            res.raise_for_status()
            data = res.json()
            issues = data.get("issues", [])
            count += len(issues)
            next_page_token = data.get("nextPageToken")
            is_last = data.get("isLast", True)
            if not next_page_token or is_last:
                break
        except Exception as e:
            print(f"Error for JQL {jql}: {e}")
            break
    return count

print("Starting counts...")
count_before = count_jira_tickets(jql1)
print(f"Total resolved JIRA tickets before 2026: {count_before}")
count_after = count_jira_tickets(jql2)
print(f"Total resolved JIRA tickets in 2026: {count_after}")

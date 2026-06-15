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

url = f"{jira_base}/{jira_api_path}/search/jql"

def get_total_for_jql(jql):
    payload = {
        "jql": jql,
        "maxResults": 1,
        "fields": ["id"]
    }
    try:
        res = requests.post(url, auth=auth, headers=headers, json=payload, timeout=15)
        # Print error details if status code is 400
        if res.status_code == 400:
            print("Bad Request Details:", res.text)
        res.raise_for_status()
        # In case the response doesn't have "total" but has "isLast"/"nextPageToken", let's print keys
        data = res.json()
        return data.get("total") or f"No total in keys: {list(data.keys())}"
    except Exception as e:
        return f"Error: {e}"

base_jql = '(project = "IT - Salsa Jeans" OR project = "SAP - Desenvolvimento") AND statusCategory = Done'
print("Total resolved in JIRA:", get_total_for_jql(base_jql))

for year in [2023, 2024, 2025, 2026]:
    jql_year = f'{base_jql} AND resolved >= "{year}-01-01" AND resolved <= "{year}-12-31"'
    print(f"Total resolved in JIRA for {year}: {get_total_for_jql(jql_year)}")

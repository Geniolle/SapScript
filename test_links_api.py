import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv
import requests

sys.path.append('c:/workspace/sap-script/sap_script_web_cockpit_v2/web_api')
load_dotenv('c:/workspace/sap-script/.env')

jira_base = os.getenv("JIRA_DADOS_COMP_HASH", "").strip().rstrip("/")
jira_api_path = os.getenv("JIRA_DADOS_HASH", "rest/api/3").strip().strip("/")
auth = (os.getenv("JIRA_EMAIL"), os.getenv("JIRA_TOKEN"))

jql = "project IN ('IT - Salsa Jeans', 'SAP S4 HANA - DADOS') AND status NOT IN ('Concluído', 'Done', 'Resolvido') ORDER BY updated DESC"
url = f"{jira_base}/{jira_api_path}/search"
params = {
    "jql": jql,
    "maxResults": 10,
    "fields": "issuelinks"
}
res = requests.get(url, auth=auth, headers={"Accept": "application/json"}, params=params)
res.raise_for_status()

for issue in res.json().get("issues", []):
    links = issue.get("fields", {}).get("issuelinks", [])
    if links:
        print(f"Links for {issue['key']}:")
        print(json.dumps(links, indent=2))
        break

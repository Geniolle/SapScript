import sys
sys.path.append('c:/workspace/sap-script/sap_script_web_cockpit_v2/web_api')
from jira_client import fetch_jira_tickets_from_api
tickets = fetch_jira_tickets_from_api(limit=5)


import sys
sys.path.append('c:/workspace/sap-script/sap_script_web_cockpit_v2/web_api')
from jira_client import fetch_jira_tickets_from_api

tickets = fetch_jira_tickets_from_api()
for t in tickets:
    if 'IZ-52980' in t['key']:
        print("IZ-52980:", t)
    if t.get('linked_keys'):
        print(t['key'], "has linked_keys:", t['linked_keys'])

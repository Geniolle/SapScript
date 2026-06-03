import sqlite3
import json
conn = sqlite3.connect('/data/sap_script_jobs.sqlite3')
conn.execute("DELETE FROM jobs WHERE task='sap_agent_analysis'")
conn.commit()
print("Jobs deleted")

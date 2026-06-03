import sqlite3
import json

conn = sqlite3.connect('c:/data/sap_script_jobs.sqlite3')
conn.row_factory = sqlite3.Row

# Get IZ-52980
row = conn.execute("SELECT key, linked_keys FROM jira_tickets WHERE key = 'IZ-52980'").fetchone()
if row:
    print(f"IZ-52980 found! linked_keys={row['linked_keys']}")
else:
    print("IZ-52980 not found in db.")
    
# Get some other tickets with linked keys to see if they saved
rows = conn.execute("SELECT key, linked_keys FROM jira_tickets WHERE linked_keys IS NOT NULL AND linked_keys != '[]' LIMIT 5").fetchall()
print("Tickets with links:")
for r in rows:
    print(dict(r))

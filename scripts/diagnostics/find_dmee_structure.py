import os
import json
from pathlib import Path
import dotenv
from pyrfc import Connection

dotenv.load_dotenv(Path(__file__).resolve().parents[2] / ".env")

params = {
    "user": os.environ["SAP_QAD_USER"],
    "passwd": os.environ["SAP_QAD_PASSWD"],
    "ashost": os.environ["SAP_QAD_ASHOST"],
    "sysnr": os.environ["SAP_QAD_SYSNR"],
    "client": os.environ["SAP_QAD_CLIENT"],
    "lang": os.getenv("SAP_QAD_LANG", "PT"),
}

conn = Connection(**params)

# Check DMEEX% in DD02L
res_dmeex = conn.call("RFC_READ_TABLE", QUERY_TABLE="DD02L", FIELDS=[{"FIELDNAME": "TABNAME"}], OPTIONS=[{"TEXT": "TABNAME LIKE 'DMEEX%' AND AS4LOCAL = 'A'"}], ROWCOUNT=200)
tables_dmeex = [r["WA"].strip() for r in res_dmeex.get("DATA", [])]
print("Tabelas DMEEX:", tables_dmeex)

# Check fields of DMEE_TREE
f_tree = conn.call("DDIF_FIELDINFO_GET", TABNAME="DMEE_TREE")
print("\nCampos de DMEE_TREE:", [f["FIELDNAME"] for f in f_tree.get("DFIES_TAB", [])])

# Search trees matching *SEPA* or *CGI* in DMEE_TREE
res_tree = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE", DELIMITER="|", FIELDS=[{"FIELDNAME": "TREE_TYPE"}, {"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "PARENT_ID"}, {"FIELDNAME": "DMEEX"}, {"FIELDNAME": "EXTENSIBLE"}], OPTIONS=[{"TEXT": "TREE_ID LIKE '%SEPA%' OR TREE_ID LIKE '%CGI%'"}], ROWCOUNT=100)
print("\nDMEE_TREE matches:")
for r in res_tree.get("DATA", []):
    print("  ", r["WA"])

# Search DMEE_TREE_HEAD for versions
res_head = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_HEAD", DELIMITER="|", FIELDS=[{"FIELDNAME": "TREE_TYPE"}, {"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "VERSION"}, {"FIELDNAME": "FIRSTNODE_ID"}], OPTIONS=[{"TEXT": "TREE_ID LIKE '%SEPA%' OR TREE_ID LIKE '%CGI%'"}], ROWCOUNT=100)
print("\nDMEE_TREE_HEAD matches:")
for r in res_head.get("DATA", []):
    print("  ", r["WA"])

conn.close()


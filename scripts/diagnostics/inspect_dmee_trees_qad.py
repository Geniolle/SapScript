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

# 1. Inspect fields of DMEE_TREE_COND
cond_fields = conn.call("DDIF_FIELDINFO_GET", TABNAME="DMEE_TREE_COND")
print("=== CAMPOS DE DMEE_TREE_COND ===")
print([f["FIELDNAME"] for f in cond_fields.get("DFIES_TAB", [])])

# 2. Inspect fields of DMEE_TREE_NODE
node_fields = conn.call("DDIF_FIELDINFO_GET", TABNAME="DMEE_TREE_NODE")
print("\n=== CAMPOS DE DMEE_TREE_NODE ===")
print([f["FIELDNAME"] for f in node_fields.get("DFIES_TAB", [])])

# 3. Search trees in DMEE_TREE
print("\n=== PROCURAR ARVORES ESPECIFICAS EM DMEE_TREE ===")
target_trees = [
    "Z_SEPA_CT",
    "Z_PT_CGI_XML_CT_V9",
    "CGI_CT_V9",
    "PT_CGI_XML_CT_V9",
    "CGI_XML_CT_V9",
    "ZSEPA_CT",
]

for t in target_trees:
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE", DELIMITER="|",
                    FIELDS=[{"FIELDNAME": "TREE_TYPE"}, {"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "PARENT_ID"}, {"FIELDNAME": "DMEEX"}, {"FIELDNAME": "EXTENSIBLE"}],
                    OPTIONS=[{"TEXT": f"TREE_ID = '{t}'"}])
    rows = [r["WA"].split("|") for r in res.get("DATA", [])]
    print(f"Tree '{t}':", rows)

# 4. Search any tree with V9
print("\n=== ARVORES COM V9 NO NOME ===")
res_v9 = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE", DELIMITER="|",
                   FIELDS=[{"FIELDNAME": "TREE_TYPE"}, {"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "PARENT_ID"}, {"FIELDNAME": "DMEEX"}],
                   OPTIONS=[{"TEXT": "TREE_ID LIKE '%V9%'"}])
for r in res_v9.get("DATA", []):
    print("  ", r["WA"])

# 5. Search any Z* tree in PAYM
print("\n=== ARVORES Z* EM PAYM ===")
res_z = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE", DELIMITER="|",
                  FIELDS=[{"FIELDNAME": "TREE_TYPE"}, {"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "PARENT_ID"}, {"FIELDNAME": "DMEEX"}],
                  OPTIONS=[{"TEXT": "TREE_TYPE = 'PAYM' AND TREE_ID LIKE 'Z%'"}])
for r in res_z.get("DATA", []):
    print("  ", r["WA"])

conn.close()


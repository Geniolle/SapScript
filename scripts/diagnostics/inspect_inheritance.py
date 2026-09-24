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

# Check DMEE_TREE_NODE_R
f_node_r = conn.call("DDIF_FIELDINFO_GET", TABNAME="DMEE_TREE_NODE_R")
print("Campos DMEE_TREE_NODE_R:", [f["FIELDNAME"] for f in f_node_r.get("DFIES_TAB", [])])

# Check if DMEE_TREE_NODE_R has entries for our trees
trees = ["Z_PT_CGI_XML_CT_V9", "CGI_CT_V9", "PT_CGI_XML_CT_V9"]
for t in trees:
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE_R", DELIMITER="|",
                    FIELDS=[{"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "VERSION"}],
                    OPTIONS=[{"TEXT": f"TREE_TYPE = 'PAYM' AND TREE_ID = '{t}'"}],
                    ROWCOUNT=10)
    print(f"DMEE_TREE_NODE_R for {t}: {len(res.get('DATA', []))} rows")

# Compare NODE_IDs of CGI_CT_V9 vs Z_PT_CGI_XML_CT_V9
def get_node_ids(tree_id, ver):
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE", DELIMITER="|",
                    FIELDS=[{"FIELDNAME": "NODE_ID"}, {"FIELDNAME": "TECH_NAME"}, {"FIELDNAME": "EX_STATUS"}],
                    OPTIONS=[{"TEXT": f"TREE_TYPE = 'PAYM' AND TREE_ID = '{tree_id}' AND VERSION = '{ver}'"}],
                    ROWCOUNT=1000)
    return {r["WA"].split("|")[0].strip(): (r["WA"].split("|")[1].strip(), r["WA"].split("|")[2].strip() if len(r["WA"].split("|")) > 2 else "") for r in res.get("DATA", [])}

cgi_nodes = get_node_ids("CGI_CT_V9", "000")
z_v9_000 = get_node_ids("Z_PT_CGI_XML_CT_V9", "000")
z_v9_999 = get_node_ids("Z_PT_CGI_XML_CT_V9", "999")

print(f"\nCGI_CT_V9 ver 000 node count: {len(cgi_nodes)}")
print(f"Z_PT_CGI_XML_CT_V9 ver 000 node count: {len(z_v9_000)}")
print(f"Z_PT_CGI_XML_CT_V9 ver 999 node count: {len(z_v9_999)}")

common_000 = set(cgi_nodes.keys()) & set(z_v9_000.keys())
extra_z_000 = set(z_v9_000.keys()) - set(cgi_nodes.keys())
missing_z_000 = set(cgi_nodes.keys()) - set(z_v9_000.keys())

print(f"\nCommon node IDs between CGI_CT_V9 and Z_PT_CGI_XML_CT_V9 (000): {len(common_000)}")
print(f"Extra node IDs in Z_PT_CGI_XML_CT_V9 (000): {len(extra_z_000)}")
print(f"Node IDs only in CGI_CT_V9: {len(missing_z_000)}")

print("\nExtra nodes in Z_PT_CGI_XML_CT_V9 (000):")
for nid in list(extra_z_000)[:20]:
    print(f"  {nid}: tech_name={z_v9_000[nid][0]} ex_status={z_v9_000[nid][1]}")

# Also compare Z 000 vs Z 999
print(f"\nDiff between Z_PT_CGI_XML_CT_V9 ver 000 and ver 999:")
in_000_not_999 = set(z_v9_000.keys()) - set(z_v9_999.keys())
in_999_not_000 = set(z_v9_999.keys()) - set(z_v9_000.keys())
print(f"In 000 but not in 999 ({len(in_000_not_999)}):", [f"{k}:{z_v9_000[k][0]}" for k in list(in_000_not_999)[:20]])
print(f"In 999 but not in 000 ({len(in_999_not_000)}):", [f"{k}:{z_v9_999[k][0]}" for k in list(in_999_not_000)[:20]])

conn.close()


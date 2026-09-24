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

def read_nodes(tree_id, ver):
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE", DELIMITER="|",
                    FIELDS=[
                        {"FIELDNAME": "NODE_ID"},
                        {"FIELDNAME": "TECH_NAME"},
                        {"FIELDNAME": "PARENT_ID"},
                        {"FIELDNAME": "NODE_TYPE"},
                        {"FIELDNAME": "EX_STATUS"},
                        {"FIELDNAME": "MP_SC_TAB"},
                        {"FIELDNAME": "MP_SC_FLD"},
                        {"FIELDNAME": "MP_CONST"},
                        {"FIELDNAME": "MP_EXIT_FUNC"},
                        {"FIELDNAME": "ATOM_HANDL"},
                    ],
                    OPTIONS=[
                        {"TEXT": "TREE_TYPE = 'PAYM' "},
                        {"TEXT": f"AND TREE_ID = '{tree_id}' "},
                        {"TEXT": f"AND VERSION = '{ver}' "},
                    ],
                    ROWCOUNT=1000,
                    GET_SORTED="X")
    data = {}
    for r in res.get("DATA", []):
        parts = [p.strip() for p in r["WA"].split("|")]
        data[parts[0]] = {
            "node_id": parts[0],
            "tech_name": parts[1],
            "parent_id": parts[2],
            "node_type": parts[3],
            "ex_status": parts[4],
            "tab": parts[5],
            "fld": parts[6],
            "const": parts[7],
            "exit": parts[8],
            "atom": parts[9],
        }
    return data

def read_redefined(tree_id, ver):
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE_R", DELIMITER="|",
                    FIELDS=[
                        {"FIELDNAME": "NODE_ID"},
                        {"FIELDNAME": "NODE_REDEFINED"},
                        {"FIELDNAME": "NODE_DEACTIVATED"},
                    ],
                    OPTIONS=[
                        {"TEXT": "TREE_TYPE = 'PAYM' "},
                        {"TEXT": f"AND TREE_ID = '{tree_id}' "},
                        {"TEXT": f"AND VERSION = '{ver}' "},
                    ],
                    ROWCOUNT=1000,
                    GET_SORTED="X")
    data = {}
    for r in res.get("DATA", []):
        parts = [p.strip() for p in r["WA"].split("|")]
        data[parts[0]] = {
            "redefined": parts[1] == "X",
            "deactivated": parts[2] == "X"
        }
    return data

cgi_nodes = read_nodes("CGI_CT_V9", "000")
z_nodes_000 = read_nodes("Z_PT_CGI_XML_CT_V9", "000")
z_redef_000 = read_redefined("Z_PT_CGI_XML_CT_V9", "000")

print(f"CGI_CT_V9 (000): {len(cgi_nodes)} nodes")
print(f"Z_PT_CGI_XML_CT_V9 (000): {len(z_nodes_000)} nodes")

# Check overlap
common = set(cgi_nodes.keys()) & set(z_nodes_000.keys())
added_in_z = set(z_nodes_000.keys()) - set(cgi_nodes.keys())
only_in_cgi = set(cgi_nodes.keys()) - set(z_nodes_000.keys())

print(f"Common node IDs: {len(common)}")
print(f"Added in Z_PT_CGI_XML_CT_V9: {len(added_in_z)} nodes")
print(f"Only in CGI_CT_V9: {len(only_in_cgi)} nodes")

# Check EX_STATUS values in Z
ex_statuses = set(n["ex_status"] for n in z_nodes_000.values())
print(f"Distinct EX_STATUS in Z_PT_CGI_XML_CT_V9: {ex_statuses}")

# For added nodes, show details
print("\nAdded nodes in Z_PT_CGI_XML_CT_V9 (000):")
for nid in sorted(added_in_z):
    n = z_nodes_000[nid]
    print(f"  {nid}: tech={n['tech_name']} type={n['node_type']} parent={n['parent_id']} ex_status={n['ex_status']} map={n['tab']}.{n['fld']} const={n['const']} exit={n['exit']}")

# Check redefined nodes: compare mapping with CGI
print("\nSample Redefined nodes (comparing Z vs CGI):")
redef_count = 0
diff_mapping = 0
for nid, r in z_redef_000.items():
    if r["redefined"]:
        redef_count += 1
        zn = z_nodes_000.get(nid, {})
        cn = cgi_nodes.get(nid, {})
        z_map = f"{zn.get('tab')}.{zn.get('fld')}|const:{zn.get('const')}|exit:{zn.get('exit')}"
        c_map = f"{cn.get('tab')}.{cn.get('fld')}|const:{cn.get('const')}|exit:{cn.get('exit')}"
        if z_map != c_map:
            diff_mapping += 1
            if diff_mapping <= 15:
                print(f"  {nid} ({zn.get('tech_name')}):")
                print(f"     CGI: {c_map}")
                print(f"     Z_V9: {z_map}")

print(f"\nTotal redefined in DMEE_TREE_NODE_R: {redef_count}")
print(f"Redefined with different mapping vs CGI: {diff_mapping}")

conn.close()


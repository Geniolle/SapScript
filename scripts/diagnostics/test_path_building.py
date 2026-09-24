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

def read_tree_complete(tree_id, ver="000"):
    # Read nodes
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE", DELIMITER="|",
                    FIELDS=[
                        {"FIELDNAME": "NODE_ID"},
                        {"FIELDNAME": "TECH_NAME"},
                        {"FIELDNAME": "PARENT_ID"},
                        {"FIELDNAME": "FIRSTCHILD_ID"},
                        {"FIELDNAME": "BROTHER_ID"},
                        {"FIELDNAME": "NODE_TYPE"},
                        {"FIELDNAME": "LEV"},
                        {"FIELDNAME": "LENGTH"},
                        {"FIELDNAME": "DATA_TYPE"},
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
    nodes = {}
    for r in res.get("DATA", []):
        parts = [p.strip() for p in r["WA"].split("|")]
        nodes[parts[0]] = {
            "node_id": parts[0],
            "tech_name": parts[1],
            "parent_id": parts[2],
            "firstchild_id": parts[3],
            "brother_id": parts[4],
            "node_type": parts[5],
            "lev": parts[6],
            "length": parts[7],
            "data_type": parts[8],
            "tab": parts[9],
            "fld": parts[10],
            "const": parts[11],
            "exit": parts[12],
            "atom": parts[13],
        }
    return nodes

def build_path(node_id, nodes):
    path_parts = []
    curr_id = node_id
    visited = set()
    while curr_id and curr_id in nodes and curr_id not in visited:
        visited.add(curr_id)
        node = nodes[curr_id]
        tech = node["tech_name"]
        ntype = node["node_type"]
        if ntype == "ATOM":
            seg = f":atom:{tech}"
        elif ntype == "XMAT":
            seg = f"@{tech}"
        elif ntype == "TECH":
            seg = f"(TECH:{tech})"
        else:
            seg = tech
        path_parts.append(seg)
        curr_id = node["parent_id"]
    return "/" + "/".join(reversed(path_parts))

z_sepa = read_tree_complete("Z_SEPA_CT", "000")
z_v9 = read_tree_complete("Z_PT_CGI_XML_CT_V9", "000")

print("Z_SEPA_CT paths matching InitgPty:")
for nid, n in z_sepa.items():
    p = build_path(nid, z_sepa)
    if "InitgPty" in p:
        print(f"  {p} | map={n['tab']}.{n['fld']} const={n['const']} exit={n['exit']}")

print("\nZ_PT_CGI_XML_CT_V9 paths matching InitgPty:")
for nid, n in z_v9.items():
    p = build_path(nid, z_v9)
    if "InitgPty" in p:
        print(f"  {p} | map={n['tab']}.{n['fld']} const={n['const']} exit={n['exit']}")

conn.close()


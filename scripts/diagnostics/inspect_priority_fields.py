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

def read_full_tree(tree_id, ver="000"):
    # Read nodes
    # We query DMEE_TREE_NODE fields
    fields = [
        "NODE_ID", "TECH_NAME", "REF_NAME", "PARENT_ID", "FIRSTCHILD_ID", "BROTHER_ID",
        "NODE_TYPE", "LENGTH", "DATA_TYPE", "LEV", "ATOM_HANDL", "MP_OFFSET",
        "MP_SC_TAB", "MP_SC_FLD", "MP_SC_OFFSET", "MP_IF_TP", "MP_SC_NODE",
        "MP_CONST", "CV_RULE", "MP_EXIT_FUNC", "CK_EXIT_FUNC"
    ]
    opts = [
        {"TEXT": "TREE_TYPE = 'PAYM' AND "},
        {"TEXT": f"TREE_ID = '{tree_id}' AND "},
        {"TEXT": f"VERSION = '{ver}'"},
    ]
    
    nodes = {}
    skip = 0
    batch = 500
    while True:
        res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE", DELIMITER="|",
                        FIELDS=[{"FIELDNAME": f} for f in fields],
                        OPTIONS=opts,
                        ROWCOUNT=batch,
                        ROWSKIPS=skip,
                        GET_SORTED="X")
        data = res.get("DATA", [])
        if not data:
            break
        for r in data:
            vals = [v.strip() for v in r["WA"].split("|")]
            row = dict(zip(fields, vals))
            nodes[row["NODE_ID"]] = row
        if len(data) < batch:
            break
        skip += len(data)

    # Read DMEE_TREE_NODE_R
    redef_map = {}
    try:
        res_r = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE_R", DELIMITER="|",
                          FIELDS=[{"FIELDNAME": "NODE_ID"}, {"FIELDNAME": "NODE_REDEFINED"}, {"FIELDNAME": "NODE_DEACTIVATED"}],
                          OPTIONS=opts,
                          ROWCOUNT=1000)
        for r in res_r.get("DATA", []):
            vals = [v.strip() for v in r["WA"].split("|")]
            redef_map[vals[0]] = {"redefined": vals[1] == "X", "deactivated": vals[2] == "X"}
    except Exception as e:
        print(f"Error reading NODE_R for {tree_id} {ver}: {e}")

    # Read DMEE_TREE_COND
    conds_map = {}
    try:
        cond_fields = [
            "NODE_ID", "COND_NUMBER", "PAR_OPEN", "ARG1_FLD", "ARG1_TAB", "ARG1_CONST",
            "ARG1_TYPE", "OPERATOR", "ARG2_FLD", "ARG2_TAB", "ARG2_CONST", "ARG2_TYPE",
            "PAR_CLOSE", "LINK_OPERATOR", "CD_EXIT_FUNC"
        ]
        skip_c = 0
        while True:
            res_c = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_COND", DELIMITER="|",
                              FIELDS=[{"FIELDNAME": f} for f in cond_fields],
                              OPTIONS=opts,
                              ROWCOUNT=500,
                              ROWSKIPS=skip_c,
                              GET_SORTED="X")
            cdata = res_c.get("DATA", [])
            if not cdata:
                break
            for r in cdata:
                vals = [v.strip() for v in r["WA"].split("|")]
                crow = dict(zip(cond_fields, vals))
                nid = crow["NODE_ID"]
                if nid not in conds_map:
                    conds_map[nid] = []
                conds_map[nid].append(crow)
            if len(cdata) < 500:
                break
            skip_c += len(cdata)
    except Exception as e:
        print(f"Error reading COND for {tree_id} {ver}: {e}")

    # Enrich nodes
    for nid, node in nodes.items():
        node["redefined"] = redef_map.get(nid, {}).get("redefined", False)
        node["deactivated"] = redef_map.get(nid, {}).get("deactivated", False)
        node["conditions"] = conds_map.get(nid, [])

    return nodes

def build_path(node_id, nodes):
    path_parts = []
    curr_id = node_id
    visited = set()
    while curr_id and curr_id in nodes and curr_id not in visited:
        visited.add(curr_id)
        node = nodes[curr_id]
        tech = node["TECH_NAME"]
        ntype = node["NODE_TYPE"]
        if ntype == "ATOM":
            seg = f":atom:{tech}"
        elif ntype == "XMAT":
            seg = f"@{tech}"
        elif ntype == "TECH":
            seg = f"(TECH:{tech})"
        else:
            seg = tech
        path_parts.append(seg)
        curr_id = node["PARENT_ID"]
    return "/" + "/".join(reversed(path_parts))

print("Loading trees...")
trees = {
    "Z_SEPA_CT": read_full_tree("Z_SEPA_CT", "000"),
    "Z_PT_CGI_XML_CT_V9": read_full_tree("Z_PT_CGI_XML_CT_V9", "000"),
    "CGI_CT_V9": read_full_tree("CGI_CT_V9", "000"),
    "PT_CGI_XML_CT_V9": read_full_tree("PT_CGI_XML_CT_V9", "000"),
}

priority_targets = [
    ("A: InitgPty/Nm", "InitgPty/Nm"),
    ("B: InitgPty/Id", "InitgPty/Id"),
    ("C: Dbtr/Id", "Dbtr/Id"),
    ("D: DbtrAcct/Ccy", "DbtrAcct/Ccy"),
    ("E: ChrgBr", "ChrgBr"),
    ("F: InstrId", "InstrId"),
    ("G: Authstn/Prtry", "Authstn"),
    ("H: RmtInf/Ustrd", "Ustrd"),
]

for label, p_filter in priority_targets:
    print(f"\n=======================================================")
    print(f"PRIORITY FIELD {label} (Filter: {p_filter})")
    print(f"=======================================================")
    for tname in ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9", "PT_CGI_XML_CT_V9"]:
        tnodes = trees[tname]
        matches = []
        for nid, n in tnodes.items():
            p = build_path(nid, tnodes)
            if p_filter in p:
                matches.append((p, nid, n))
        print(f"\n--- {tname} (matches: {len(matches)}) ---")
        for p, nid, n in sorted(matches, key=lambda x: (x[0], x[1])):
            cond_str = ""
            if n["conditions"]:
                cond_strs = []
                for c in n["conditions"]:
                    c1 = f"{c['ARG1_TAB']}-{c['ARG1_FLD']}" if c['ARG1_TAB'] or c['ARG1_FLD'] else c['ARG1_CONST']
                    c2 = f"{c['ARG2_TAB']}-{c['ARG2_FLD']}" if c['ARG2_TAB'] or c['ARG2_FLD'] else c['ARG2_CONST']
                    cond_strs.append(f"{c['PAR_OPEN']}{c1} {c['OPERATOR']} {c2}{c['PAR_CLOSE']} {c['LINK_OPERATOR']}".strip())
                cond_str = " | COND: " + " ".join(cond_strs)
            flags = []
            if n["redefined"]: flags.append("REDEF")
            if n["deactivated"]: flags.append("DEACT")
            flag_str = f" [{','.join(flags)}]" if flags else ""
            print(f"  [{nid}]{flag_str} {p}")
            print(f"      Type: {n['NODE_TYPE']} | Tab.Fld: {n['MP_SC_TAB']}.{n['MP_SC_FLD']} | Const: '{n['MP_CONST']}' | Exit: {n['MP_EXIT_FUNC']}{cond_str}")

conn.close()

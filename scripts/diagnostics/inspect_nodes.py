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

def read_table(table, fields, options, rowcount=0):
    opt_payload = [{"TEXT": opt} for opt in options]
    out = []
    skip = 0
    batch = 500
    while True:
        res = conn.call("RFC_READ_TABLE", QUERY_TABLE=table, DELIMITER="|",
                        FIELDS=[{"FIELDNAME": f} for f in fields],
                        OPTIONS=opt_payload,
                        ROWCOUNT=batch,
                        ROWSKIPS=skip,
                        GET_SORTED="X")
        data = res.get("DATA", [])
        if not data:
            break
        for r in data:
            vals = [v.strip() for v in r["WA"].split("|")]
            out.append(dict(zip(fields, vals)))
        if len(data) < batch or (rowcount > 0 and len(out) >= rowcount):
            break
        skip += len(data)
    return out

trees = ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9", "PT_CGI_XML_CT_V9"]

print("=== CONTAGEM REAL DE NOS ===")
for t in trees:
    for ver in ["000", "999"]:
        opts = [
            f"TREE_TYPE = 'PAYM' AND",
            f"TREE_ID = '{t}' AND",
            f"VERSION = '{ver}'"
        ]
        nodes = read_table("DMEE_TREE_NODE", ["NODE_ID", "PARENT_ID", "TECH_NAME", "NODE_TYPE"], opts)
        print(f"Tree: {t} | Version: {ver} | Nodes: {len(nodes)}")
        if len(nodes) > 0 and len(nodes) <= 10:
            print("  Nodes sample:")
            for n in nodes:
                print(f"    id={n['NODE_ID']} parent={n['PARENT_ID']} tech={n['TECH_NAME']} type={n['NODE_TYPE']}")

conn.close()


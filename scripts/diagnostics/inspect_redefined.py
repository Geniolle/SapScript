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

for ver in ["000", "999"]:
    opts = [
        "TREE_TYPE = 'PAYM' AND",
        "TREE_ID = 'Z_PT_CGI_XML_CT_V9' AND",
        f"VERSION = '{ver}'"
    ]
    rows = read_table("DMEE_TREE_NODE_R",
                      ["NODE_ID", "NODE_REDEFINED", "NODE_DEACTIVATED", "VAR_NAME"],
                      opts)
    print(f"\n=== DMEE_TREE_NODE_R for Z_PT_CGI_XML_CT_V9 (ver {ver}): {len(rows)} rows ===")
    redef = [r for r in rows if r.get("NODE_REDEFINED") == "X"]
    deact = [r for r in rows if r.get("NODE_DEACTIVATED") == "X"]
    print(f"  Redefined nodes count: {len(redef)}")
    for r in redef:
        print(f"    redefined: {r['NODE_ID']} var={r.get('VAR_NAME')}")
    print(f"  Deactivated nodes count: {len(deact)}")
    for r in deact:
        print(f"    deactivated: {r['NODE_ID']} var={r.get('VAR_NAME')}")

conn.close()


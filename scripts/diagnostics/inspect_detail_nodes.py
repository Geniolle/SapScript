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

# Let's search for "Nm" or "InitgPty" nodes in Z_SEPA_CT and Z_PT_CGI_XML_CT_V9
for t in ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9"]:
    opts = [
        "TREE_TYPE = 'PAYM' AND",
        f"TREE_ID = '{t}' AND",
        "VERSION = '000' AND",
        "( TECH_NAME = 'Nm' OR TECH_NAME = 'InitgPty' OR TECH_NAME = 'AUST1' OR TECH_NAME = 'NAMEZ' )"
    ]
    nodes = read_table("DMEE_TREE_NODE",
                       ["NODE_ID", "PARENT_ID", "TECH_NAME", "NODE_TYPE", "MP_SC_TAB", "MP_SC_FLD", "MP_CONST", "MP_EXIT_FUNC", "ATOM_HANDL"],
                       opts)
    print(f"\nNodes matching Nm/InitgPty/AUST1/NAMEZ in {t} (ver 000):")
    for n in nodes:
        print("  ", n)

conn.close()


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

trees_versions = [
    ("Z_PT_CGI_XML_CT_V9", "000"),
    ("Z_PT_CGI_XML_CT_V9", "999"),
    ("PT_CGI_XML_CT_V9", "000"),
    ("PT_CGI_XML_CT_V9", "999"),
]

for t, v in trees_versions:
    opts = [
        {"TEXT": "TREE_TYPE = 'PAYM' AND "},
        {"TEXT": f"TREE_ID = '{t}' AND "},
        {"TEXT": f"VERSION = '{v}'"},
    ]
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE_R", DELIMITER="|",
                    FIELDS=[{"FIELDNAME": "NODE_ID"}, {"FIELDNAME": "NODE_REDEFINED"}, {"FIELDNAME": "NODE_DEACTIVATED"}, {"FIELDNAME": "VAR_NAME"}],
                    OPTIONS=opts,
                    ROWCOUNT=1000)
    data = res.get("DATA", [])
    redef = [r["WA"].split("|")[0].strip() for r in data if r["WA"].split("|")[1].strip() == "X"]
    deact = [r["WA"].split("|")[0].strip() for r in data if r["WA"].split("|")[2].strip() == "X"]
    print(f"{t} v{v}: total_r={len(data)}, redefined={len(redef)}, deactivated={len(deact)}")

conn.close()

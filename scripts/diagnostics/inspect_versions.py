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

trees = ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9", "PT_CGI_XML_CT_V9"]

print("=== DMEE_TREE_HEAD ===")
for t in trees:
    res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_HEAD", DELIMITER="|",
                    FIELDS=[{"FIELDNAME": "TREE_TYPE"}, {"FIELDNAME": "TREE_ID"}, {"FIELDNAME": "VERSION"}, {"FIELDNAME": "FIRSTNODE_ID"}, {"FIELDNAME": "VERSION_DESCRIPTION"}],
                    OPTIONS=[{"TEXT": f"TREE_TYPE = 'PAYM' AND TREE_ID = '{t}'"}])
    for r in res.get("DATA", []):
        print("  ", r["WA"])

print("\n=== CONTAGEM DE NOS POR VERSAO EM DMEE_TREE_NODE ===")
for t in trees:
    for ver in ["000", "999"]:
        res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_NODE", DELIMITER="|",
                        FIELDS=[{"FIELDNAME": "NODE_ID"}],
                        OPTIONS=[{"TEXT": f"TREE_TYPE = 'PAYM' AND TREE_ID = '{t}' AND VERSION = '{ver}'"}],
                        ROWCOUNT=5000)
        cnt = len(res.get("DATA", []))
        print(f"Tree '{t}' ver '{ver}': {cnt} nos")

conn.close()


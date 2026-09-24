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

fields_needed = ["TREE_TYPE", "TREE_ID", "VERSION", "FIRSTNODE_ID", "VERS_USER", "VERS_DATE", "VERS_TIME", "VERSION_TYPE", "REQUEST_ID"]
res = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE_HEAD", DELIMITER="|",
                FIELDS=[{"FIELDNAME": f} for f in fields_needed],
                OPTIONS=[{"TEXT": "TREE_ID = 'Z_PT_CGI_XML_CT_V9'"}])
for r in res.get("DATA", []):
    print("DMEE_TREE_HEAD Z_PT_CGI_XML_CT_V9:", r["WA"])

res_tree = conn.call("RFC_READ_TABLE", QUERY_TABLE="DMEE_TREE", DELIMITER="|",
                     FIELDS=[{"FIELDNAME": f} for f in ["TREE_ID", "CREA_USER", "CREA_DATE", "CREA_TIME", "CHNG_USER", "CHNG_DATE", "CHNG_TIME", "PARENT_ID", "DMEEX"]],
                     OPTIONS=[{"TEXT": "TREE_ID = 'Z_PT_CGI_XML_CT_V9'"}])
for r in res_tree.get("DATA", []):
    print("DMEE_TREE Z_PT_CGI_XML_CT_V9:", r["WA"])

conn.close()


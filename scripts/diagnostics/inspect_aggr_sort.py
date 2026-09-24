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

for tab in ["DMEE_TREE_AGGR", "DMEE_TREE_SORT", "DMEE_TREE_COND"]:
    res = conn.call("DDIF_FIELDINFO_GET", TABNAME=tab)
    print(f"Campos {tab}:", [f["FIELDNAME"] for f in res.get("DFIES_TAB", [])])

conn.close()


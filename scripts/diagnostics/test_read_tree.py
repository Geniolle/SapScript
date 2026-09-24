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

# Test reading all nodes of Z_SEPA_CT
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
                    {"TEXT": "AND TREE_ID = 'Z_SEPA_CT' "},
                    {"TEXT": "AND VERSION = '000' "},
                ],
                ROWCOUNT=500,
                GET_SORTED="X")

data = res.get("DATA", [])
print(f"Lidos com sucesso {len(data)} nos de Z_SEPA_CT!")
for r in data[:5]:
    print("  ", r["WA"])

conn.close()


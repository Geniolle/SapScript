"""Teste simples de LFB1"""

import os, sys
from pathlib import Path
from pyrfc import Connection

project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)

env_prefix = "QAD"
params = {
    "user": os.getenv(f"SAP_{env_prefix}_USER", "").strip(),
    "passwd": os.getenv(f"SAP_{env_prefix}_PASSWD", "").strip(),
    "ashost": os.getenv(f"SAP_{env_prefix}_ASHOST", "").strip(),
    "sysnr": os.getenv(f"SAP_{env_prefix}_SYSNR", "").strip(),
    "client": os.getenv(f"SAP_{env_prefix}_CLIENT", "").strip(),
}

conn = Connection(**params)

print("Testando LFB1 com Fornecedor 0010007092 e Empresa 2010")
print("-" * 80)

try:
    result = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE="LFB1",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "LIFNR"}, {"FIELDNAME": "BUKRS"}, {"FIELDNAME": "ZTERM"}],
        OPTIONS=[{"TEXT": "LIFNR = '0010007092' AND BUKRS = '2010'"}],
        ROWCOUNT=0
    )

    rows = result.get("DATA", [])
    print(f"✓ LFB1 encontrada: {len(rows)} linha(s)\n")

    if rows:
        parts = str(rows[0].get("WA", "")).split("|")
        print(f"  LIFNR: {parts[0].strip() if len(parts) > 0 else 'N/A'}")
        print(f"  BUKRS: {parts[1].strip() if len(parts) > 1 else 'N/A'}")
        print(f"  ZTERM: {parts[2].strip() if len(parts) > 2 else 'N/A'}")
    else:
        print("✗ Nenhuma linha encontrada")
        print("\nTentando apenas com LIFNR:")

        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="LFB1",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "LIFNR"}, {"FIELDNAME": "BUKRS"}],
            OPTIONS=[{"TEXT": "LIFNR = '0010007092'"}],
            ROWCOUNT=5
        )

        rows = result.get("DATA", [])
        print(f"Resultados com LIFNR: {len(rows)} linha(s)")
        for row in rows:
            parts = str(row.get("WA", "")).split("|")
            print(f"  LIFNR: {parts[0].strip()}, BUKRS: {parts[1].strip() if len(parts) > 1 else 'N/A'}")

except Exception as e:
    print(f"✗ ERRO: {e}")

conn.close()

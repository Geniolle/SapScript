"""Continuar teste dos JOINs restantes"""

import os, sys
from pathlib import Path
from pyrfc import Connection

project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)

def connect():
    env_prefix = "QAD"
    params = {
        "user": os.getenv(f"SAP_{env_prefix}_USER", "").strip(),
        "passwd": os.getenv(f"SAP_{env_prefix}_PASSWD", "").strip(),
        "ashost": os.getenv(f"SAP_{env_prefix}_ASHOST", "").strip(),
        "sysnr": os.getenv(f"SAP_{env_prefix}_SYSNR", "").strip(),
        "client": os.getenv(f"SAP_{env_prefix}_CLIENT", "").strip(),
    }
    return Connection(**params)

conn = connect()

print("="*100)
print("CONTINUANDO TESTES DOS JOINs")
print("="*100)

# RESUMO até agora
print("\n[RESUMO ATÉ AGORA]")
print("EKKO:      PASSA ✓")
print("EKET:      PASSA ✓")
print("LFB1:      PASSA ✓ (ZTERM=1029)")

# T052
print("\n[TESTE] T052 (INNER JOIN T052)")
print("-" * 100)
print("Condição ON: F~ZTERM EQ D~ZTERM")
print("Valor: ZTERM = '1029'")

try:
    result = conn.call("RFC_READ_TABLE",
        QUERY_TABLE="T052",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "ZTERM"}],
        OPTIONS=[{"TEXT": "ZTERM = '1029'"}],
        ROWCOUNT=0
    )
    rows = len(result.get("DATA", []))
    print(f"Resultado: {rows} linha(s) - {'PASSA' if rows > 0 else 'FALHA'} ✓" if rows > 0 else f"Resultado: {rows} linhas - FALHA ✗")
except Exception as e:
    print(f"✗ ERRO: {e}")

# ZFI_PAY_DATE_T
print("\n[TESTE] ZFI_PAY_DATE_T (INNER JOIN ZFI_PAY_DATE_T)")
print("-" * 100)

# Primeiro obter ZZCOORI e ZZEXPVZ de EKKO
try:
    result_ekko = conn.call("RFC_READ_TABLE",
        QUERY_TABLE="EKKO",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "ZZCOORI"}, {"FIELDNAME": "ZZEXPVZ"}],
        OPTIONS=[{"TEXT": "EBELN = '4300000002'"}],
        ROWCOUNT=0
    )

    ekko_data = result_ekko.get("DATA", [])
    if ekko_data:
        parts = str(ekko_data[0].get("WA", "")).split("|")
        zzcoori = parts[0].strip()
        zzexpvz = parts[1].strip() if len(parts) > 1 else ""

        print(f"Condição ON: E~ZZCOORI EQ A~ZZCOORI AND E~ZZEXPVZ EQ A~ZZEXPVZ")
        print(f"Valores: ZZCOORI='{zzcoori}', ZZEXPVZ='{zzexpvz}'")

        try:
            result = conn.call("RFC_READ_TABLE",
                QUERY_TABLE="ZFI_PAY_DATE_T",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": "ZZCOORI"}, {"FIELDNAME": "ZZEXPVZ"}, {"FIELDNAME": "NDAYS"}],
                OPTIONS=[{"TEXT": f"ZZCOORI = '{zzcoori}' AND ZZEXPVZ = '{zzexpvz}'"}],
                ROWCOUNT=0
            )
            rows = len(result.get("DATA", []))
            if rows > 0:
                parts = str(result["DATA"][0].get("WA", "")).split("|")
                ndays = parts[2].strip() if len(parts) > 2 else "N/A"
                print(f"Resultado: {rows} linha(s) - PASSA ✓")
                print(f"  NDAYS: {ndays}")
            else:
                print(f"Resultado: 0 linhas - FALHA ✗")
        except Exception as e:
            print(f"⚠️  ERRO ao consultar ZFI_PAY_DATE_T: {e}")
    else:
        print("EKKO não retornou ZZCOORI/ZZEXPVZ")

except Exception as e:
    print(f"✗ ERRO ao obter dados de EKKO: {e}")

# EKPO
print("\n[TESTE] EKPO (INNER JOIN EKPO)")
print("-" * 100)

# Obter EBELP de EKET
try:
    result_eket = conn.call("RFC_READ_TABLE",
        QUERY_TABLE="EKET",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "EBELP"}],
        OPTIONS=[{"TEXT": "EBELN = '4300000002'"}],
        ROWCOUNT=0
    )

    eket_data = result_eket.get("DATA", [])
    if eket_data:
        ebelp = str(eket_data[0].get("WA", "")).split("|")[0].strip()

        print(f"Condição ON: G~EBELN = A~EBELN AND G~EBELP = B~EBELP")
        print(f"Valores: EBELN='4300000002', EBELP='{ebelp}'")

        try:
            result = conn.call("RFC_READ_TABLE",
                QUERY_TABLE="EKPO",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": "EBELN"}, {"FIELDNAME": "EBELP"}, {"FIELDNAME": "NETWR"}],
                OPTIONS=[{"TEXT": f"EBELN = '4300000002' AND EBELP = '{ebelp}'"}],
                ROWCOUNT=0
            )
            rows = len(result.get("DATA", []))
            if rows > 0:
                parts = str(result["DATA"][0].get("WA", "")).split("|")
                netwr = parts[2].strip() if len(parts) > 2 else "N/A"
                print(f"Resultado: {rows} linha(s) - PASSA ✓")
                print(f"  NETWR: {netwr}")
            else:
                print(f"Resultado: 0 linhas - FALHA ✗")
        except Exception as e:
            print(f"✗ ERRO ao consultar EKPO: {e}")
    else:
        print("EKET não retornou dados")

except Exception as e:
    print(f"✗ ERRO ao obter dados de EKET: {e}")

conn.close()

print("\n" + "="*100)
print("FIM DO TESTE")
print("="*100)

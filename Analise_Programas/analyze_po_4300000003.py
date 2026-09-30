"""
Análise de PO 4300000003

Baseado nos critérios descobertos:
- Precisa de ZZCOORI e ZZEXPVZ preenchidos em EKKO
- Precisa de posições (EKPO)
- Precisa de correspondência em ZFI_PAY_DATE_T
"""

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

def read_table(conn, table, fields, options):
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": f} for f in fields],
            OPTIONS=[{"TEXT": o} for o in options],
            ROWCOUNT=0
        )
        rows = result.get("DATA", [])
        data = []
        for row in rows:
            parts = [p.strip() for p in str(row.get("WA", "")).split("|")]
            data.append({field: (parts[i] if i < len(parts) else "") for i, field in enumerate(fields)})
        return data
    except Exception as e:
        return []

conn = connect()
po = "4300000003"

print("="*100)
print(f"ANÁLISE: PO {po}")
print("="*100)

# Verificar critérios
print("\n[CRITÉRIO 1] EKKO deve existir")
ekko = read_table(conn, "EKKO", ["EBELN", "BUKRS", "ZZCOORI", "ZZEXPVZ"], [f"EBELN = '{po}'"])

if not ekko:
    print(f"✗ FALHA: PO {po} não existe")
    conn.close()
    sys.exit(1)

print(f"✓ PASSA: PO encontrada")
ekko_data = ekko[0]
print(f"  BUKRS: {ekko_data.get('BUKRS')}")
print(f"  ZZCOORI: '{ekko_data.get('ZZCOORI', '')}' {'(VAZIO)' if not ekko_data.get('ZZCOORI') else ''}")
print(f"  ZZEXPVZ: '{ekko_data.get('ZZEXPVZ', '')}' {'(VAZIO)' if not ekko_data.get('ZZEXPVZ') else ''}")

# Verificar ZZCOORI e ZZEXPVZ
print("\n[CRITÉRIO 2] ZZCOORI e ZZEXPVZ devem estar preenchidos")
zzcoori = ekko_data.get('ZZCOORI', '').strip()
zzexpvz = ekko_data.get('ZZEXPVZ', '').strip()

if not zzcoori or not zzexpvz:
    print(f"✗ FALHA: ZZCOORI ou ZZEXPVZ vazios")
    print(f"\n❌ CONCLUSÃO: PO {po} NÃO DEVERIA APARECER")
    print(f"\n   Razão: Faltam valores ZZCOORI/ZZEXPVZ em EKKO")
    print(f"   Resultado: INNER JOIN com ZFI_PAY_DATE_T falhará")
    conn.close()
    sys.exit(1)

print(f"✓ PASSA: ZZCOORI='{zzcoori}', ZZEXPVZ='{zzexpvz}'")

# Verificar EKPO
print("\n[CRITÉRIO 3] EKPO deve ter posições")
ekpo = read_table(conn, "EKPO", ["EBELP", "PSTYP"], [f"EBELN = '{po}'"])

if not ekpo:
    print(f"✗ FALHA: Nenhuma posição em EKPO")
    print(f"\n❌ CONCLUSÃO: PO {po} NÃO DEVERIA APARECER")
    print(f"\n   Razão: Sem posições em EKPO")
    conn.close()
    sys.exit(1)

print(f"✓ PASSA: {len(ekpo)} posição(ões)")
for row in ekpo:
    pstyp = row.get('PSTYP', '0')
    tipo = "SERVIÇO" if pstyp == "5" else "MATERIAL"
    print(f"  EBELP={row.get('EBELP')}, PSTYP={pstyp} ({tipo})")

# Verificar ZFI_PAY_DATE_T
print("\n[CRITÉRIO 4] ZFI_PAY_DATE_T deve ter correspondência")
zfi = read_table(conn, "ZFI_PAY_DATE_T", ["ZZCOORI", "ZZEXPVZ", "NDAYS"],
                [f"ZZCOORI = '{zzcoori}' AND ZZEXPVZ = '{zzexpvz}'"])

if not zfi:
    print(f"✗ FALHA: Nenhuma correspondência em ZFI_PAY_DATE_T")
    print(f"\n❌ CONCLUSÃO: PO {po} NÃO DEVERIA APARECER")
    print(f"\n   Razão: Sem correspondência em ZFI_PAY_DATE_T para ZZCOORI='{zzcoori}', ZZEXPVZ='{zzexpvz}'")
    conn.close()
    sys.exit(1)

print(f"✓ PASSA: 1 linha encontrada")
print(f"  NDAYS: {zfi[0].get('NDAYS')}")

# Se chegou aqui, tudo passa
print("\n" + "="*100)
print(f"✅ CONCLUSÃO: PO {po} DEVERIA APARECER")
print("="*100)
print(f"\nMOTIVO:")
print(f"  1. ✓ EKKO existe")
print(f"  2. ✓ ZZCOORI preenchido ('{zzcoori}')")
print(f"  3. ✓ ZZEXPVZ preenchido ('{zzexpvz}')")
print(f"  4. ✓ EKPO tem {len(ekpo)} posição(ões)")
print(f"  5. ✓ ZFI_PAY_DATE_T com correspondência")
print(f"\nA PO {po} passa em TODOS os critérios de seleção do AUTO_INST_ASSIGN.")

conn.close()

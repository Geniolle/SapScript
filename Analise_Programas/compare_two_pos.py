"""
Comparação Técnica: PO 4300000002 vs PO 4000055997

Objetivo: Descobrir por que uma aparece e a outra não aparece em AUTO_INST_ASSIGN
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
        print(f"  ERRO: {e}")
        return []

conn = connect()
print("✓ Conectado a S4Q/100\n")

po_problema = "4300000002"
po_controlo = "4000055997"

print("="*120)
print("COMPARAÇÃO TÉCNICA: PO 4300000002 vs PO 4000055997")
print("="*120)

# 1. EKKO
print("\n[1] EKKO - Cabeçalho das POs")
print("-"*120)

ekko_fields = ["EBELN", "BUKRS", "BSART", "LIFNR", "WAERS", "WKURS", "KUFIX", "LOEKZ", "BEDAT", "AEDAT", "STATU", "ZZCOORI", "ZZEXPVZ"]

ekko_4300 = read_table(conn, "EKKO", ekko_fields, [f"EBELN = '{po_problema}'"])
ekko_4000 = read_table(conn, "EKKO", ekko_fields, [f"EBELN = '{po_controlo}'"])

if ekko_4300 and ekko_4000:
    print(f"\n{'Campo':<20} | {po_problema:<20} | {po_controlo:<20} | {'Diferença':<15}")
    print("-"*120)

    for field in ekko_fields:
        val1 = ekko_4300[0].get(field, "N/A")
        val2 = ekko_4000[0].get(field, "N/A")
        diff = "DIFERENTE ⚠️" if val1 != val2 else "igual"
        print(f"{field:<20} | {val1:<20} | {val2:<20} | {diff:<15}")

# 2. EKPO
print("\n\n[2] EKPO - Posições das POs")
print("-"*120)

ekpo_fields = ["EBELN", "EBELP", "PSTYP", "LOEKZ", "MENGE", "MEINS", "NETWR", "BRTWR"]

ekpo_4300 = read_table(conn, "EKPO", ekpo_fields, [f"EBELN = '{po_problema}'"])
ekpo_4000 = read_table(conn, "EKPO", ekpo_fields, [f"EBELN = '{po_controlo}'"])

print(f"\nPO {po_problema}: {len(ekpo_4300)} posição(ões)")
for row in ekpo_4300:
    print(f"  EBELP={row.get('EBELP')}, PSTYP={row.get('PSTYP')}, MENGE={row.get('MENGE')}, NETWR={row.get('NETWR')}")

print(f"\nPO {po_controlo}: {len(ekpo_4000)} posição(ões)")
for row in ekpo_4000:
    print(f"  EBELP={row.get('EBELP')}, PSTYP={row.get('PSTYP')}, MENGE={row.get('MENGE')}, NETWR={row.get('NETWR')}")

# 3. EKET
print("\n\n[3] EKET - Prazos de Entrega")
print("-"*120)

eket_fields = ["EBELN", "EBELP", "ETENR", "EINDT", "MENGE", "WEMNG"]

eket_4300 = read_table(conn, "EKET", eket_fields, [f"EBELN = '{po_problema}'"])
eket_4000 = read_table(conn, "EKET", eket_fields, [f"EBELN = '{po_controlo}'"])

print(f"\nPO {po_problema}: {len(eket_4300)} linha(s)")
for row in eket_4300:
    print(f"  EBELP={row.get('EBELP')}, EINDT={row.get('EINDT')}, MENGE={row.get('MENGE')}, WEMNG={row.get('WEMNG')}")

print(f"\nPO {po_controlo}: {len(eket_4000)} linha(s)")
for row in eket_4000:
    print(f"  EBELP={row.get('EBELP')}, EINDT={row.get('EINDT')}, MENGE={row.get('MENGE')}, WEMNG={row.get('WEMNG')}")

# 4. LFB1 + T052
print("\n\n[4] LFB1 + T052 - Dados de Fornecedor e Termos")
print("-"*120)

if ekko_4300 and ekko_4000:
    for po_val, ekko_val in [(po_problema, ekko_4300[0]), (po_controlo, ekko_4000[0])]:
        lifnr = ekko_val.get("LIFNR")
        bukrs = ekko_val.get("BUKRS")

        lfb1_data = read_table(conn, "LFB1", ["LIFNR", "BUKRS", "ZTERM"],
                               [f"LIFNR = '{lifnr}' AND BUKRS = '{bukrs}'"])

        if lfb1_data:
            zterm = lfb1_data[0].get("ZTERM")
            t052_data = read_table(conn, "T052", ["ZTERM"], [f"ZTERM = '{zterm}'"])

            print(f"\nPO {po_val}:")
            print(f"  LIFNR={lifnr}, BUKRS={bukrs} → LFB1-ZTERM={zterm}")
            print(f"  T052 com ZTERM={zterm}: {len(t052_data)} linha(s)")

# 5. ZFI_PAY_DATE_T - PONTO PRINCIPAL
print("\n\n[5] ZFI_PAY_DATE_T - PONTO CRÍTICO")
print("-"*120)

if ekko_4300 and ekko_4000:
    print(f"\n{'Campo':<20} | {po_problema:<25} | {po_controlo:<25} | {'Status':<15}")
    print("-"*120)

    for po_val, ekko_val in [(po_problema, ekko_4300[0]), (po_controlo, ekko_4000[0])]:
        zzcoori = ekko_val.get("ZZCOORI")
        zzexpvz = ekko_val.get("ZZEXPVZ")

        print(f"\nEKKO-ZZCOORI     | '{zzcoori}'{'(vazio)' if not zzcoori else '':<12} | ", end="")

        zfi_data = read_table(conn, "ZFI_PAY_DATE_T",
                             ["ZZCOORI", "ZZEXPVZ", "NDAYS"],
                             [f"ZZCOORI = '{zzcoori}' AND ZZEXPVZ = '{zzexpvz}'"])

        if zfi_data:
            ndays = zfi_data[0].get("NDAYS")
            print(f"'{zfi_data[0].get('ZZCOORI')}'")
            print(f"EKKO-ZZEXPVZ     | '{zzexpvz}' | '{zfi_data[0].get('ZZEXPVZ')}'")
            print(f"ZFI_PAY_DATE_T   | {len(zfi_data)} linha(s) - NDAYS={ndays}")
        else:
            print(f"(vazio) | ZFI_PAY_DATE_T = 0 linhas")

# 6. ZFI_DOC_EX_LOG_T
print("\n\n[6] ZFI_DOC_EX_LOG_T - Log de Processamento")
print("-"*120)

for po_val in [po_problema, po_controlo]:
    log_data = read_table(conn, "ZFI_DOC_EX_LOG_T", ["EBELN"], [f"EBELN = '{po_val}'"])
    print(f"PO {po_val}: {len(log_data)} registo(s) no log")

# 7. Reproduzir JOINs
print("\n\n[7] REPRODUZIR INNER JOINs")
print("-"*120)

print(f"\n{'JOIN/Tabela':<20} | {po_problema:<15} | {po_controlo:<15}")
print("-"*120)

for po_val in [po_problema, po_controlo]:
    joins_status = {}

    # EKKO
    ekko = read_table(conn, "EKKO", ["EBELN"], [f"EBELN = '{po_val}'"])
    joins_status["EKKO"] = "PASSA" if ekko else "FALHA"

    # EKET
    eket = read_table(conn, "EKET", ["EBELN", "EBELP"], [f"EBELN = '{po_val}'"])
    joins_status["EKET"] = "PASSA" if eket else "FALHA"

    if ekko:
        # LFB1
        lifnr = ekko[0].get("LIFNR")
        bukrs = ekko[0].get("BUKRS")
        lfb1 = read_table(conn, "LFB1", ["LIFNR", "BUKRS", "ZTERM"],
                         [f"LIFNR = '{lifnr}' AND BUKRS = '{bukrs}'"])
        joins_status["LFB1"] = "PASSA" if lfb1 else "FALHA"

        # T052
        if lfb1:
            zterm = lfb1[0].get("ZTERM")
            t052 = read_table(conn, "T052", ["ZTERM"], [f"ZTERM = '{zterm}'"])
            joins_status["T052"] = "PASSA" if t052 else "FALHA"
        else:
            joins_status["T052"] = "FALHA"

        # ZFI_PAY_DATE_T
        zzcoori = ekko[0].get("ZZCOORI")
        zzexpvz = ekko[0].get("ZZEXPVZ")
        zfi = read_table(conn, "ZFI_PAY_DATE_T", ["ZZCOORI", "ZZEXPVZ"],
                        [f"ZZCOORI = '{zzcoori}' AND ZZEXPVZ = '{zzexpvz}'"])
        joins_status["ZFI_PAY_DATE_T"] = "PASSA" if zfi else "FALHA"

        # EKPO
        if eket:
            ebelp = eket[0].get("EBELP")
            ekpo = read_table(conn, "EKPO", ["EBELN", "EBELP"],
                             [f"EBELN = '{po_val}' AND EBELP = '{ebelp}'"])
            joins_status["EKPO"] = "PASSA" if ekpo else "FALHA"
        else:
            joins_status["EKPO"] = "FALHA"

# Print results for both POs
if po_val == po_problema:
    results_4300 = joins_status
elif po_val == po_controlo:
    results_4000 = joins_status

print(f"\n{'JOIN/Tabela':<20} | {po_problema:<15} | {po_controlo:<15}")
print("-"*120)

for join in ["EKKO", "EKET", "LFB1", "T052", "ZFI_PAY_DATE_T", "EKPO"]:
    status_4300 = results_4300.get(join, "?") if 'results_4300' in locals() else "?"
    status_4000 = results_4000.get(join, "?") if 'results_4000' in locals() else "?"
    print(f"{join:<20} | {status_4300:<15} | {status_4000:<15}")

conn.close()

print("\n" + "="*120)
print("ANÁLISE CONCLUÍDA")
print("="*120)

#!/usr/bin/env python3
"""
Extração completa dos objetos reais da OT S4DK953607:
1. R3TR ENHO ZFI_ENH_CHANGE_BANK_ID_3
2. R3TR ENHO ZFI_ENH_CHANGE_ACC_TYPE
3. R3TR FUGR ZFI_CHANGE_BANKID_FG
4. LIMU REPS LZFI_CHANGE_BANKID_FGUXX
"""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

project_root = Path(__file__).resolve().parent.parent
load_dotenv(project_root / ".env", override=True)

from pyrfc import Connection

def get_conn(env: str):
    prefix = f"SAP_{env.upper()}"
    return Connection(
        user=os.environ[f"{prefix}_USER"],
        passwd=os.environ[f"{prefix}_PASSWD"],
        ashost=os.environ[f"{prefix}_ASHOST"],
        sysnr=os.environ[f"{prefix}_SYSNR"],
        client=os.environ[f"{prefix}_CLIENT"],
        lang="PT",
    )

def inspect_enhancement(conn, env: str):
    print(f"\n=================== INSPECIONANDO EM {env} ===================")
    out_dir = project_root / "docs" / "analises" / "FIAPPT_CBR_COPE_SAP75" / env
    out_dir.mkdir(parents=True, exist_ok=True)

    # 1. Procurar código dos ENHO na REPOSRC ou ENHOBJ
    # Enhancements guardam código em includes gerados pelo SAP tipo ZFI_ENH_CHANGE_ACC_TYPE=======IP
    # Vamos pesquisar na REPOSRC por PROGNAME LIKE 'ZFI%'
    try:
        reposrc = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="REPOSRC",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "PROGNAME"}, {"FIELDNAME": "R3STATE"}, {"FIELDNAME": "UNAM"}, {"FIELDNAME": "UDAT"}],
            OPTIONS=[{"TEXT": "PROGNAME LIKE 'ZFI_ENH%' OR PROGNAME LIKE '%CHANGE_BANKID%'"}],
        )
        print(f"Programas/Includes encontrados em {env}:")
        for r in reposrc.get("DATA", []):
            print("  ", r["WA"])
            prog = r["WA"].split("|")[0].strip()
            try:
                res = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME=prog)
                lines = [l.get("LINE", "") for l in res.get("SOURCE_EXTENDED", [])]
                fpath = out_dir / f"{prog}.abap"
                fpath.write_text("\n".join(lines), encoding="utf-8")
                print(f"   -> Salvo {prog} ({len(lines)} linhas)")
            except Exception as ex:
                print(f"   -> Erro lendo {prog}: {ex}")
    except Exception as e:
        print("Erro REPOSRC:", e)

    # 2. Procurar tabelas de customising para COPE / CBR
    print("\n--- Tabelas de Configuração COPE / CBR ---")
    try:
        tab_list = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="DD02L",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "TABNAME"}, {"FIELDNAME": "TABCLASS"}],
            OPTIONS=[{"TEXT": "TABNAME LIKE 'FIAPPT_CBR%' OR TABNAME LIKE 'ZFI%CBR%'"}],
        )
        print("Tabelas CBR encontradas:", tab_list.get("DATA"))
        for t in tab_list.get("DATA", []):
            tname = t["WA"].split("|")[0].strip()
            try:
                tdata = conn.call("RFC_READ_TABLE", QUERY_TABLE=tname, DELIMITER="|", ROWCOUNT=50)
                print(f"\nDados de {tname} ({len(tdata.get('DATA', []))} linhas):")
                for row in tdata.get("DATA", [])[:10]:
                    print("  ", row["WA"])
            except Exception as ex:
                print(f"Erro lendo {tname}: {ex}")
    except Exception as e:
        print("Erro DD02L:", e)

if __name__ == "__main__":
    conn_dev = get_conn("DEV")
    inspect_enhancement(conn_dev, "DEV")
    conn_dev.close()

    conn_qad = get_conn("QAD")
    inspect_enhancement(conn_qad, "QAD")
    conn_qad.close()

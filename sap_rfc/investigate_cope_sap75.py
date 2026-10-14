#!/usr/bin/env python3
"""
Investigação e extração de código-fonte e mapeamentos do processo COPE / Balança Comercial.
Transação: FIAPPT_CBR
Classe: CL_FIAPPT_CBR
OT: S4DK953607
Ambientes: QAD vs PRD
"""
import os
import sys
import json
from pathlib import Path
from dotenv import load_dotenv

project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

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
        lang=os.getenv(f"{prefix}_LANG", "PT"),
    )

def read_source_via_rfc(conn, name: str) -> list[str]:
    # Tenta RPY_PROGRAM_READ primeiro
    try:
        res = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME=name, LANGUAGE="PT")
        src_tab = res.get("SOURCE_TAB") or res.get("SOURCE") or []
        lines = []
        for line_obj in src_tab:
            if isinstance(line_obj, dict):
                lines.append(line_obj.get("LINE", ""))
            else:
                lines.append(str(line_obj))
        if lines:
            return lines
    except Exception as e:
        print(f"RPY_PROGRAM_READ falhou para {name}: {e}")

    # Fallback: RFC_READ_TABLE na REPOSRC ou DOKTL / TRDIR
    try:
        res = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="REPOSRC",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "PROGNAME"}, {"FIELDNAME": "R3STATE"}],
            OPTIONS=[{"TEXT": f"PROGNAME LIKE '{name[:15]}%'"}],
            ROWCOUNT=10,
        )
        print(f"REPOSRC search {name}: {res.get('DATA')}")
    except Exception as e:
        print(f"RFC_READ_TABLE REPOSRC falhou: {e}")

    return []

def read_class_includes(conn, class_name: str) -> list[str]:
    # Em SAP, classes OO guardam os métodos em includes CL_...===CCAU, CL_...===CCIMP, CL_...===CM001...
    # Pesquisar includes na REPOSRC ou SEO*
    try:
        pattern = f"{class_name.upper()}%"
        res = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="REPOSRC",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "PROGNAME"}],
            OPTIONS=[{"TEXT": f"PROGNAME LIKE '{pattern[:15]}%'"}],
            ROWCOUNT=50,
        )
        progs = [r["WA"].strip() for r in res.get("DATA", [])]
        print(f"Includes encontrados para {class_name}: {progs}")
        return progs
    except Exception as e:
        print(f"Erro ao listar includes: {e}")
        return []

def main():
    print("=== Conectando a QAD ===")
    conn_qad = get_conn("QAD")
    print("Conectado a QAD.")

    # 1. Verificar transação FIAPPT_CBR na TSTC
    tstc = conn_qad.call(
        "RFC_READ_TABLE",
        QUERY_TABLE="TSTC",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "TCODE"}, {"FIELDNAME": "PGMNA"}],
        OPTIONS=[{"TEXT": "TCODE = 'FIAPPT_CBR'"}],
    )
    print("TSTC FIAPPT_CBR:", tstc.get("DATA"))
    prog_name = ""
    if tstc.get("DATA"):
        prog_name = tstc["DATA"][0]["WA"].split("|")[1].strip()

    # 2. Verificar detalhes da OT S4DK953607 na E070 / E071
    print("\n=== Verificando OT S4DK953607 ===")
    try:
        e070 = conn_qad.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E070",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "TRKORR"}, {"FIELDNAME": "AS4USER"}, {"FIELDNAME": "AS4DATE"}, {"FIELDNAME": "AS4TEXT"}],
            OPTIONS=[{"TEXT": "TRKORR = 'S4DK953607'"}],
        )
        print("E070:", e070.get("DATA"))
        
        e071 = conn_qad.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E071",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "PGMID"}, {"FIELDNAME": "OBJECT"}, {"FIELDNAME": "OBJ_NAME"}],
            OPTIONS=[{"TEXT": "TRKORR = 'S4DK953607'"}],
        )
        print("E071 (Objetos da OT):", e071.get("DATA"))
    except Exception as e:
        print("Erro ao ler E070/E071:", e)

    # 3. Ler o código do programa principal
    output_dir = project_root / "docs" / "analises" / "FIAPPT_CBR_COPE_SAP75"
    output_dir.mkdir(parents=True, exist_ok=True)

    if prog_name:
        src = read_source_via_rfc(conn_qad, prog_name)
        if src:
            fpath = output_dir / f"{prog_name}_QAD.abap"
            fpath.write_text("\n".join(src), encoding="utf-8")
            print(f"Salvo código do programa principal em: {fpath} ({len(src)} linhas)")

    # 4. Ler includes da classe CL_FIAPPT_CBR
    includes = read_class_includes(conn_qad, "CL_FIAPPT_CBR")
    for inc in includes:
        inc_clean = inc.split("|")[0].strip()
        if not inc_clean:
            continue
        inc_src = read_source_via_rfc(conn_qad, inc_clean)
        if inc_src:
            fpath = output_dir / f"{inc_clean.replace('/', '_')}_QAD.abap"
            fpath.write_text("\n".join(inc_src), encoding="utf-8")
            print(f"Salvo include {inc_clean} em: {fpath} ({len(inc_src)} linhas)")

    # 5. Investigar tabelas de mapeamento de COPE / Conta 27112000
    print("\n=== Verificando Mapeamento da Conta 27112000 em QAD ===")
    # Procurar tabelas com 'CBR' ou 'COPE' ou 'FIAPPT' na DD02T
    try:
        dd02t = conn_qad.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="DD02T",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "TABNAME"}, {"FIELDNAME": "DDTEXT"}],
            OPTIONS=[{"TEXT": "SPRAS = 'P' AND ( TABNAME LIKE 'FIAPPT%' OR TABNAME LIKE '%CBR%' )"}],
            ROWCOUNT=30,
        )
        print("Tabelas CBR encontradas:", dd02t.get("DATA"))
    except Exception as e:
        print("Erro DD02T:", e)

    conn_qad.close()

if __name__ == "__main__":
    main()

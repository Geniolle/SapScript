#!/usr/bin/env python3
"""
Extração completa dos métodos da classe CL_FIAPPT_CBR e do programa RFFOPT_CBR em QAD e PRD.
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

def dump_sources(env: str):
    print(f"\n=================== DUMPING {env} ===================")
    conn = get_conn(env)
    output_dir = project_root / "docs" / "analises" / "FIAPPT_CBR_COPE_SAP75" / env
    output_dir.mkdir(parents=True, exist_ok=True)

    # 1. Obter mapeamento de métodos via TMDIR se disponível
    method_map = {}
    try:
        methods_res = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="TMDIR",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "METHODNAME"}, {"FIELDNAME": "INCCOMP"}],
            OPTIONS=[{"TEXT": "CLASSNAME = 'CL_FIAPPT_CBR'"}],
        )
        for r in methods_res.get("DATA", []):
            parts = r["WA"].split("|")
            m_name = parts[0].strip()
            inc = parts[1].strip()
            method_map[inc] = m_name
    except Exception as e:
        print(f"Aviso TMDIR: {e}")

    print(f"Total métodos mapeados em TMDIR ({env}): {len(method_map)}")

    # 2. Ler todos os includes da classe
    pool_name = "CL_FIAPPT_CBR".ljust(30, "=") + "CP"
    pool_res = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME=pool_name)
    includes = [i.get("INCLNAME") for i in pool_res.get("INCLUDE_TAB", []) if i.get("INCLNAME")]

    found_fix_lines = []

    for inc in includes:
        res = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME=inc)
        lines = [l.get("LINE", "") for l in res.get("SOURCE_EXTENDED", [])]
        
        # Nome amigável com método se existir
        inc_suffix = inc.split("=")[-1]
        m_name = method_map.get(inc_suffix, inc_suffix)
        filename = f"{inc_suffix}_{m_name}.abap"
        
        (output_dir / filename).write_text("\n".join(lines), encoding="utf-8")
        
        # Procurar menções de 'X' ou 'E' ou 'ACC' ou '27112000' ou 'S4DK953607'
        text_full = "\n".join(lines)
        if "S4DK953607" in text_full or "SAP75" in text_full or "MPEREIRA" in text_full:
            print(f"!!! ENCONTRADA MENÇÃO DE OT/AUTOR EM {filename} ({env}) !!!")
            found_fix_lines.append((filename, "OT mention"))
        
        for idx, line in enumerate(lines, 1):
            if "27112000" in line or "E2010" in line or "E2020" in line:
                print(f"[{env}] [{filename}:{idx}] {line.strip()}")

    # 3. Ler o report RFFOPT_CBR e seus includes
    rffopt_res = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME="RFFOPT_CBR")
    rffopt_includes = ["RFFOPT_CBR"]
    for l in rffopt_res.get("SOURCE_EXTENDED", []):
        line = l.get("LINE", "").strip()
        if line.startswith("INCLUDE "):
            inc_name = line.split()[1].rstrip(".")
            rffopt_includes.append(inc_name)

    for r_inc in rffopt_includes:
        res = conn.call("RPY_PROGRAM_READ", PROGRAM_NAME=r_inc)
        lines = [l.get("LINE", "") for l in res.get("SOURCE_EXTENDED", [])]
        (output_dir / f"{r_inc}.abap").write_text("\n".join(lines), encoding="utf-8")

    print(f"Todos os ficheiros de {env} foram guardados em: {output_dir}")
    conn.close()

if __name__ == "__main__":
    dump_sources("QAD")
    dump_sources("PRD")

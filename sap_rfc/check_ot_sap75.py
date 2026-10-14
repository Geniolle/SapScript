#!/usr/bin/env python3
"""
Verificar a OT S4DK953607 em DEV, QAD e PRD e listar exatamente os objetos transportados.
"""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

project_root = Path(__file__).resolve().parent.parent
load_dotenv(project_root / ".env", override=True)

from pyrfc import Connection

def inspect_ot(env: str):
    print(f"\n=== OT S4DK953607 em {env} ===")
    prefix = f"SAP_{env.upper()}"
    try:
        conn = Connection(
            user=os.environ[f"{prefix}_USER"],
            passwd=os.environ[f"{prefix}_PASSWD"],
            ashost=os.environ[f"{prefix}_ASHOST"],
            sysnr=os.environ[f"{prefix}_SYSNR"],
            client=os.environ[f"{prefix}_CLIENT"],
            lang="PT",
        )
    except Exception as e:
        print(f"Erro ao conectar a {env}: {e}")
        return

    # E070
    try:
        e070 = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E070",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "TRKORR"}, {"FIELDNAME": "STRKORR"}, {"FIELDNAME": "TRSTATUS"}, {"FIELDNAME": "AS4USER"}, {"FIELDNAME": "AS4DATE"}],
            OPTIONS=[{"TEXT": "TRKORR = 'S4DK953607'"}],
        )
        print("E070:", e070.get("DATA"))
    except Exception as e:
        print("Erro E070:", e)

    # E070T (Texto)
    try:
        e070t = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E070T",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "TRKORR"}, {"FIELDNAME": "AS4TEXT"}],
            OPTIONS=[{"TEXT": "TRKORR = 'S4DK953607'"}],
        )
        print("E070T (Texto da OT):", e070t.get("DATA"))
    except Exception as e:
        print("Erro E070T:", e)

    # E071 (Objetos)
    try:
        e071 = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E071",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "PGMID"}, {"FIELDNAME": "OBJECT"}, {"FIELDNAME": "OBJ_NAME"}],
            OPTIONS=[{"TEXT": "TRKORR = 'S4DK953607'"}],
        )
        print("E071 (Objetos):", e071.get("DATA"))
    except Exception as e:
        print("Erro E071:", e)

    # Sub-tarefas se for request customising/workbench
    try:
        subtasks = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E070",
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": "TRKORR"}, {"FIELDNAME": "TRSTATUS"}, {"FIELDNAME": "AS4USER"}, {"FIELDNAME": "AS4DATE"}],
            OPTIONS=[{"TEXT": "STRKORR = 'S4DK953607'"}],
        )
        print("Tarefas filhas da OT:", subtasks.get("DATA"))
        for st in subtasks.get("DATA", []):
            task_korr = st["WA"].split("|")[0].strip()
            st_objs = conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="E071",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": "PGMID"}, {"FIELDNAME": "OBJECT"}, {"FIELDNAME": "OBJ_NAME"}],
                OPTIONS=[{"TEXT": f"TRKORR = '{task_korr}'"}],
            )
            print(f"Objetos da sub-tarefa {task_korr}:", st_objs.get("DATA"))
    except Exception as e:
        print("Erro sub-tarefas:", e)

    conn.close()

if __name__ == "__main__":
    inspect_ot("DEV")
    inspect_ot("QAD")

#!/usr/bin/env python3
"""
Verifica status da Request S4DK953657 e objetos com erro.
"""

import os
import sys
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

load_dotenv(project_root / ".env")


def get_request_info(conn, request_id: str) -> dict:
    """Obtém informações gerais da Request."""
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E070",  # Tabela de Requests
            OPTIONS=[
                {"TEXT": f"TRKORR = '{request_id}'"},
            ],
            FIELDS=[
                {"FIELDNAME": "TRKORR"},
                {"FIELDNAME": "STRUSER"},
                {"FIELDNAME": "TRSTATUS"},
                {"FIELDNAME": "TRSYSTEM"},
                {"FIELDNAME": "LASTCHANGE"},
                {"FIELDNAME": "COMMENT"},
            ]
        )

        if result.get("DATA"):
            row = result["DATA"][0]
            return {
                "TRKORR": row.get("WA", "").split("|")[0].strip(),
                "STRUSER": row.get("WA", "").split("|")[1].strip() if len(row.get("WA", "").split("|")) > 1 else "",
                "TRSTATUS": row.get("WA", "").split("|")[2].strip() if len(row.get("WA", "").split("|")) > 2 else "",
                "TRSYSTEM": row.get("WA", "").split("|")[3].strip() if len(row.get("WA", "").split("|")) > 3 else "",
            }
        return {}
    except Exception as e:
        print(f"    Erro ao ler E070: {str(e)[:80]}")
        return {}


def get_request_objects(conn, request_id: str) -> list:
    """Obtém lista de objetos na Request."""
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E071",  # Objetos na Request
            OPTIONS=[
                {"TEXT": f"TRKORR = '{request_id}'"},
            ],
            FIELDS=[
                {"FIELDNAME": "TRKORR"},
                {"FIELDNAME": "STROBJ"},
                {"FIELDNAME": "OBJECT"},
                {"FIELDNAME": "OBJNAME"},
                {"FIELDNAME": "MASTERLANG"},
                {"FIELDNAME": "DEVCLASS"},
            ]
        )

        objects = []
        if result.get("DATA"):
            for row in result["DATA"]:
                parts = row.get("WA", "").split("|")
                if len(parts) >= 6:
                    objects.append({
                        "OBJECT": parts[2].strip(),
                        "OBJNAME": parts[3].strip(),
                        "MASTERLANG": parts[4].strip(),
                        "DEVCLASS": parts[5].strip(),
                    })
        return objects
    except Exception as e:
        print(f"    Erro ao ler E071: {str(e)[:80]}")
        return []


def get_request_log(conn, request_id: str) -> list:
    """Obtém log de execução/erros da Request."""
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E07I",  # Log de objetos transportados
            OPTIONS=[
                {"TEXT": f"TRKORR = '{request_id}'"},
            ],
            FIELDS=[
                {"FIELDNAME": "TRKORR"},
                {"FIELDNAME": "OBJECT"},
                {"FIELDNAME": "OBJNAME"},
                {"FIELDNAME": "STEP"},
                {"FIELDNAME": "TCODE"},
                {"FIELDNAME": "DEVCLASS"},
                {"FIELDNAME": "AUTHOR"},
            ]
        )

        logs = []
        if result.get("DATA"):
            for row in result["DATA"]:
                parts = row.get("WA", "").split("|")
                if len(parts) >= 7:
                    logs.append({
                        "OBJECT": parts[1].strip(),
                        "OBJNAME": parts[2].strip(),
                        "STEP": parts[3].strip(),
                        "TCODE": parts[4].strip(),
                    })
        return logs
    except Exception as e:
        print(f"    Erro ao ler E07I: {str(e)[:80]}")
        return []


def main():
    print("=" * 90)
    print("🔍 ANÁLISE DE REQUEST: S4DK953657")
    print("=" * 90)
    print()

    prd_config = {
        "user": os.getenv("SAP_PRD_USER"),
        "passwd": os.getenv("SAP_PRD_PASSWD"),
        "ashost": os.getenv("SAP_PRD_ASHOST"),
        "sysnr": os.getenv("SAP_PRD_SYSNR"),
        "client": os.getenv("SAP_PRD_CLIENT"),
    }

    request_id = "S4DK953657"

    try:
        from pyrfc import Connection

        print(f"📡 Conectando a PRD (172.19.34.22:00)...")
        prd_conn = Connection(**prd_config)
        print("   ✅ Conectado\n")

        # ========== INFO GERAL ==========
        print(f"📋 Informações Gerais da Request {request_id}:")
        info = get_request_info(prd_conn, request_id)

        if info:
            print(f"   Request: {info.get('TRKORR', 'N/A')}")
            print(f"   Utilizador: {info.get('STRUSER', 'N/A')}")
            print(f"   Status: {info.get('TRSTATUS', 'N/A')}")
            print(f"   Sistema: {info.get('TRSYSTEM', 'N/A')}")
        else:
            print(f"   ⚠️  Request não encontrada em PRD")
        print()

        # ========== OBJETOS ==========
        print(f"📦 Objetos na Request:")
        objects = get_request_objects(prd_conn, request_id)

        if objects:
            print(f"   Total: {len(objects)} objetos\n")

            # Agrupa por tipo
            by_type = {}
            for obj in objects:
                obj_type = obj["OBJECT"]
                if obj_type not in by_type:
                    by_type[obj_type] = []
                by_type[obj_type].append(obj)

            for obj_type in sorted(by_type.keys()):
                items = by_type[obj_type]
                print(f"   🔹 {obj_type} ({len(items)}):")
                for item in items:
                    print(f"      - {item['OBJNAME']}")
        else:
            print(f"   ⚠️  Nenhum objeto encontrado")
        print()

        # ========== LOG DE EXECUÇÃO ==========
        print(f"📊 Log de Execução:")
        logs = get_request_log(prd_conn, request_id)

        if logs:
            print(f"   Total: {len(logs)} registos\n")
            for log in logs[:20]:  # Primeiros 20
                print(f"   - {log['OBJECT']:20} {log['OBJNAME']:30} STEP={log['STEP']:2} {log['TCODE']}")
            if len(logs) > 20:
                print(f"   ... e mais {len(logs) - 20} registos")
        else:
            print(f"   ⚠️  Nenhum log encontrado (Request pode não ter sido executada)")
        print()

        prd_conn.close()

        # ========== DIAGNÓSTICO ==========
        print("=" * 90)
        print("📌 DIAGNÓSTICO")
        print("=" * 90)
        print()

        if not info:
            print("❌ Request não encontrada em PRD")
            print()
            print("Possíveis causas:")
            print("  1. Request pode estar em DEV ou QAD, não em PRD")
            print("  2. Número da Request pode estar errado")
            print("  3. Request pode ter sido deletada")
            print()
            print("Ações:")
            print("  1. Execute SE09 em DEV/QAD")
            print("  2. Procure por 'S4DK953657'")
            print("  3. Confirme o número exacto")
        elif info.get("TRSTATUS") in ["C", "R"]:
            print(f"✅ Request está LIBERADA/RELEASED")
            print()
            print("Status: C = Released, R = Running/Completed")
            print()
            if objects:
                # Procura por Z_PT_CGI_XML_CT_V9
                estrutura_encontrada = False
                for obj in objects:
                    if "Z_PT_CGI_XML_CT_V9" in obj["OBJNAME"]:
                        estrutura_encontrada = True
                        print(f"🎯 Estrutura Z_PT_CGI_XML_CT_V9 encontrada:")
                        print(f"   Tipo: {obj['OBJECT']}")
                        print(f"   Nome: {obj['OBJNAME']}")
                        break

                if not estrutura_encontrada:
                    print("⚠️  Z_PT_CGI_XML_CT_V9 NÃO ESTÁ NESTA REQUEST")
                    print()
                    print("Objetos Z_* transportados:")
                    for obj in objects:
                        if obj["OBJNAME"].startswith("Z_"):
                            print(f"   - {obj['OBJECT']:10} {obj['OBJNAME']}")
        else:
            print(f"ℹ️  Status: {info.get('TRSTATUS', 'DESCONHECIDO')}")

        print()

    except ModuleNotFoundError:
        print("❌ PyRFC não instalado")
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

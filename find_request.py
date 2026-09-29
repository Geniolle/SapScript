#!/usr/bin/env python3
"""
Procura Request S4DK953657 em todos os ambientes (DEV, QAD, PRD).
"""

import os
import sys
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

load_dotenv(project_root / ".env")


def search_request_in_env(conn, request_id: str, env_name: str) -> dict:
    """Procura uma Request num ambiente."""
    result = {
        "found": False,
        "info": {},
        "objects": [],
        "error": None
    }

    try:
        # Tenta ler E070 sem filtro de tabela (método alternativo)
        response = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="E070",
            ROWCOUNT=100
        )

        # Se conseguir ler, processa
        if response.get("DATA"):
            for row in response["DATA"]:
                wa = row.get("WA", "")
                if request_id in wa:
                    result["found"] = True
                    result["info"] = {"raw": wa}
                    break

        if not result["found"]:
            result["error"] = "Request não encontrada"

        return result

    except Exception as e:
        result["error"] = str(e)[:100]
        return result


def get_request_via_function(conn, request_id: str) -> dict:
    """Tenta obter informações via função ABAP."""
    try:
        result = conn.call(
            "TR_READ_REQUEST",
            REQUEST=request_id,
        )

        if result:
            return {
                "found": True,
                "status": result.get("TRSTATUS", ""),
                "user": result.get("STRUSER", ""),
                "type": result.get("TRTYPE", ""),
            }
        return {"found": False}
    except Exception as e:
        return {"found": False, "error": str(e)[:80]}


def main():
    print("=" * 90)
    print("🔍 PROCURANDO REQUEST: S4DK953657")
    print("=" * 90)
    print()

    configs = {
        "DEV": {
            "user": os.getenv("SAP_DEV_USER"),
            "passwd": os.getenv("SAP_DEV_PASSWD"),
            "ashost": os.getenv("SAP_DEV_ASHOST"),
            "sysnr": os.getenv("SAP_DEV_SYSNR"),
            "client": os.getenv("SAP_DEV_CLIENT"),
            "host_ip": "172.19.66.4"
        },
        "QAD": {
            "user": os.getenv("SAP_QAD_USER"),
            "passwd": os.getenv("SAP_QAD_PASSWD"),
            "ashost": os.getenv("SAP_QAD_ASHOST"),
            "sysnr": os.getenv("SAP_QAD_SYSNR"),
            "client": os.getenv("SAP_QAD_CLIENT"),
            "host_ip": "172.19.66.22"
        },
        "PRD": {
            "user": os.getenv("SAP_PRD_USER"),
            "passwd": os.getenv("SAP_PRD_PASSWD"),
            "ashost": os.getenv("SAP_PRD_ASHOST"),
            "sysnr": os.getenv("SAP_PRD_SYSNR"),
            "client": os.getenv("SAP_PRD_CLIENT"),
            "host_ip": "172.19.34.22"
        }
    }

    request_id = "S4DK953657"
    found_in = []

    try:
        from pyrfc import Connection

        for env_name, config in configs.items():
            print(f"📡 Procurando em {env_name} ({config['host_ip']})...", end=" ")

            try:
                conn = Connection(
                    user=config["user"],
                    passwd=config["passwd"],
                    ashost=config["ashost"],
                    sysnr=config["sysnr"],
                    client=config["client"],
                )

                # Tenta via TR_READ_REQUEST (mais confiável)
                result = get_request_via_function(conn, request_id)

                conn.close()

                if result.get("found"):
                    print("✅ ENCONTRADA!")
                    found_in.append({
                        "env": env_name,
                        "status": result.get("status", "?"),
                        "user": result.get("user", "?"),
                        "type": result.get("type", "?"),
                    })
                else:
                    print("❌")

            except Exception as e:
                print(f"⚠️  Erro: {str(e)[:40]}")

        print()
        print("=" * 90)
        print("📊 RESULTADO")
        print("=" * 90)
        print()

        if not found_in:
            print(f"❌ Request {request_id} não foi encontrada em nenhum ambiente!")
            print()
            print("Possibilidades:")
            print("  1. ✓ Número da Request pode estar errado")
            print("  2. ✓ Request foi deletada")
            print("  3. ✓ Request foi completada e arquivada")
            print("  4. ✓ Sistema de transportes foi limpo")
            print()
            print("Próximas ações:")
            print("  1. Abra SE09 em DEV e procure por transportes recentes")
            print("  2. Verifique SE10 para transportes liberados")
            print("  3. Procure por transportes que contenham 'Z_PT_'")
            print("  4. Verifique o histórico de transportes: TMS → Monitoring")
        else:
            print(f"✅ Request encontrada em {len(found_in)} ambiente(s):\n")
            for item in found_in:
                print(f"  📍 {item['env']}")
                print(f"     Status: {item['status']}")
                print(f"     Utilizador: {item['user']}")
                print(f"     Tipo: {item['type']}")
                print()

    except ModuleNotFoundError:
        print("❌ PyRFC não instalado")
    except Exception as e:
        print(f"❌ Erro: {str(e)}")


if __name__ == "__main__":
    main()

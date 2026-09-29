#!/usr/bin/env python3
"""
Verifica diferenças de Z_PT_CGI_XML_CT_V9 entre QAD e PRD.
Abordagem simplificada usando RFC_READ_TABLE.
"""

import os
import sys
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

load_dotenv(project_root / ".env")


def check_struct_in_env(conn, struct_name: str, env_name: str) -> dict:
    """Verifica estrutura numa conexão SAP."""
    result = {
        "exists": False,
        "fields": [],
        "error": None
    }

    try:
        # Tenta ler a tabela DD03L (Definition of Fields in Tables)
        response = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="DD03L",
            OPTIONS=[
                {"TEXT": f"TABNAME = '{struct_name}'"},
                {"TEXT": f"AS4LOCAL = 'A'"},  # Objeto ativo
            ],
            FIELDS=[
                {"FIELDNAME": "TABNAME"},
                {"FIELDNAME": "FIELDNAME"},
                {"FIELDNAME": "POSITION"},
                {"FIELDNAME": "INTLEN"},
                {"FIELDNAME": "DATATYPE"},
                {"FIELDNAME": "LENG"},
                {"FIELDNAME": "DECIMALS"},
            ]
        )

        if response.get("DATA"):
            result["exists"] = True
            data = response["DATA"]

            for row in data:
                # Parse dos campos
                values = row.get("WA", "").split("|")
                if len(values) >= 5:
                    result["fields"].append({
                        "FIELDNAME": values[1].strip(),
                        "POSITION": int(values[2].strip()) if values[2].strip() else 0,
                        "INTLEN": int(values[3].strip()) if values[3].strip() else 0,
                        "DATATYPE": values[4].strip(),
                        "LENG": int(values[5].strip()) if len(values) > 5 and values[5].strip() else 0,
                    })

            result["count"] = len(result["fields"])
        else:
            result["error"] = "Nenhuma linha encontrada - estrutura pode não existir"

        return result

    except Exception as e:
        result["error"] = f"{str(e)[:100]}"
        return result


def compare(qad_result: dict, prd_result: dict) -> None:
    """Compara resultados entre QAD e PRD."""
    print("=" * 85)
    print("📊 RESULTADOS")
    print("=" * 85)
    print()

    if not qad_result["exists"] and not prd_result["exists"]:
        print("⚠️  Estrutura não encontrada em nenhum ambiente")
        print()
        print(f"QAD erro: {qad_result['error']}")
        print(f"PRD erro: {prd_result['error']}")
        print()
        print("Verificações:")
        print("  1. Transporte pode não ter sido liberado")
        print("  2. Nome da estrutura pode estar errado")
        print("  3. Estrutura pode estar numa Request aberta (não em produção)")
        return

    if not qad_result["exists"]:
        print("🔴 Estrutura NÃO EXISTE em QAD")
        print(f"   Erro: {qad_result['error']}")
        print()
    else:
        print(f"✅ QAD: {len(qad_result['fields'])} campos")

    if not prd_result["exists"]:
        print("🔴 Estrutura NÃO EXISTE em PRD")
        print(f"   Erro: {prd_result['error']}")
        print()
    else:
        print(f"✅ PRD: {len(prd_result['fields'])} campos")
    print()

    if qad_result["exists"] and prd_result["exists"]:
        qad_fields = {f["FIELDNAME"]: f for f in qad_result["fields"]}
        prd_fields = {f["FIELDNAME"]: f for f in prd_result["fields"]}

        only_qad = set(qad_fields.keys()) - set(prd_fields.keys())
        only_prd = set(prd_fields.keys()) - set(qad_fields.keys())

        if only_qad:
            print(f"🔴 Campos APENAS em QAD ({len(only_qad)}):")
            for f in sorted(list(only_qad))[:10]:
                print(f"   - {f}")
            if len(only_qad) > 10:
                print(f"   ... e mais {len(only_qad) - 10}")
            print()

        if only_prd:
            print(f"🟠 Campos APENAS em PRD ({len(only_prd)}):")
            for f in sorted(list(only_prd))[:10]:
                print(f"   - {f}")
            if len(only_prd) > 10:
                print(f"   ... e mais {len(only_prd) - 10}")
            print()

        if not only_qad and not only_prd:
            print("✅ ESTRUTURAS SINCRONIZADAS!")
            print("   Nenhuma diferença de campos detectada.")
            print()
            print("📌 Conclusão:")
            print("   A estrutura está idêntica em QAD e PRD.")
            print("   O erro de DMEEX pode ser de outro factor:")
            print("   - Função RFC chamada incorrectamente")
            print("   - Dados de entrada inválidos")
            print("   - Versão ABAP incompatível")
        else:
            print("🔧 Ações:")
            print("   1. Transporte a estrutura novamente (STRUC type)")
            print("   2. Sincronize as Request abertas")
            print("   3. Execute SE09 para ver transportes pendentes")


def main():
    print("=" * 85)
    print("🔍 VERIFICADOR DE ESTRUTURA Z_PT_CGI_XML_CT_V9")
    print("=" * 85)
    print()

    qad_config = {
        "user": os.getenv("SAP_QAD_USER"),
        "passwd": os.getenv("SAP_QAD_PASSWD"),
        "ashost": os.getenv("SAP_QAD_ASHOST"),
        "sysnr": os.getenv("SAP_QAD_SYSNR"),
        "client": os.getenv("SAP_QAD_CLIENT"),
    }

    prd_config = {
        "user": os.getenv("SAP_PRD_USER"),
        "passwd": os.getenv("SAP_PRD_PASSWD"),
        "ashost": os.getenv("SAP_PRD_ASHOST"),
        "sysnr": os.getenv("SAP_PRD_SYSNR"),
        "client": os.getenv("SAP_PRD_CLIENT"),
    }

    struct_name = "Z_PT_CGI_XML_CT_V9"

    try:
        from pyrfc import Connection

        print(f"📡 QAD (172.19.66.22:00)...", end=" ")
        qad_conn = Connection(**qad_config)
        print("✅")
        qad_result = check_struct_in_env(qad_conn, struct_name, "QAD")
        qad_conn.close()

        print(f"📡 PRD (172.19.34.22:00)...", end=" ")
        prd_conn = Connection(**prd_config)
        print("✅")
        prd_result = check_struct_in_env(prd_conn, struct_name, "PRD")
        prd_conn.close()

        print()
        compare(qad_result, prd_result)

    except ModuleNotFoundError:
        print("❌ PyRFC não instalado. Use: .venv-rfc\\Scripts\\python.exe")
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

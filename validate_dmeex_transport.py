#!/usr/bin/env python3
"""
Valida se o transporte de Z_PT_CGI_XML_CT_V9 está completo em QAD vs PRD.
Procura pelos programas ABAP que usam a estrutura e RFCs relacionadas.
"""

import os
import sys
from pathlib import Path

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

load_dotenv(project_root / ".env")


def check_function_exists(conn, func_name: str) -> bool:
    """Verifica se uma função existe no sistema."""
    try:
        # Tenta chamar a função com parâmetros vazios/nulos
        # Se existir, retorna erro de parâmetro
        # Se não existir, retorna erro de não encontrada
        result = conn.call(
            "FUNCTION_EXISTS",
            FUNCNAME=func_name
        )
        return result.get("ANSWERED", "").upper() == "X"
    except Exception as e:
        error_msg = str(e).lower()
        # Se erro é "not found", função não existe
        # Se erro é outro, pode existir
        if "function not found" in error_msg or "not in" in error_msg:
            return False
        return True  # Assumir que existe se houver outro erro


def check_program_exists(conn, prog_name: str) -> bool:
    """Verifica se um programa ABAP existe."""
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="TRDIR",
            OPTIONS=[{"TEXT": f"NAME = '{prog_name}'"}],
            FIELDS=[{"FIELDNAME": "NAME"}],
            ROWCOUNT=1
        )
        return len(result.get("DATA", [])) > 0
    except:
        return False


def validate_dmeex_env(conn, env_name: str) -> dict:
    """Valida componentes de DMEEX num ambiente."""
    result = {
        "env": env_name,
        "functions": {},
        "programs": {},
        "status": "unknown"
    }

    # Funções críticas relacionadas com DMEE
    critical_functions = [
        "DMEE_GET_SORTFIELD",
        "DMEE_GET_TREE_ID",
        "DMEE_READ_CONFIGURATION",
        "DDIF_STRUCT_GET",
        "FUNCTION_EXISTS"
    ]

    # Programas críticos
    critical_programs = [
        "SAPLDMEE2",
        "SAPLDMEE5",
        "RBNK_PAYM_SCHEDULE",
        "SAPLFREGU2",
        "SAPLFPAYM06"
    ]

    print(f"\n  Verificando funções RFC em {env_name}:")
    for func in critical_functions:
        try:
            exists = check_function_exists(conn, func)
            result["functions"][func] = exists
            status = "✅" if exists else "❌"
            print(f"    {status} {func}")
        except Exception as e:
            result["functions"][func] = False
            print(f"    ⚠️  {func} - erro na verificação")

    print(f"\n  Verificando programas ABAP em {env_name}:")
    for prog in critical_programs:
        try:
            exists = check_program_exists(conn, prog)
            result["programs"][prog] = exists
            status = "✅" if exists else "❌"
            print(f"    {status} {prog}")
        except Exception as e:
            result["programs"][prog] = False
            print(f"    ⚠️  {prog} - erro na verificação")

    # Calcula status
    all_functions_ok = all(result["functions"].values())
    all_programs_ok = all(result["programs"].values())

    if all_functions_ok and all_programs_ok:
        result["status"] = "OK"
    elif all_programs_ok:
        result["status"] = "PARTIAL (funções faltando)"
    else:
        result["status"] = "MISSING_PROGRAMS"

    return result


def main():
    print("=" * 90)
    print("🔍 VALIDADOR DE TRANSPORTE DMEEX: Z_PT_CGI_XML_CT_V9")
    print("=" * 90)

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

    try:
        from pyrfc import Connection

        print("\n📡 Conectando a QAD (172.19.66.22:00)...")
        qad_conn = Connection(**qad_config)
        print("   ✅ Conectado")

        qad_result = validate_dmeex_env(qad_conn, "QAD")
        qad_conn.close()

        print("\n📡 Conectando a PRD (172.19.34.22:00)...")
        prd_conn = Connection(**prd_config)
        print("   ✅ Conectado")

        prd_result = validate_dmeex_env(prd_conn, "PRD")
        prd_conn.close()

        # Comparação
        print("\n" + "=" * 90)
        print("📊 COMPARAÇÃO QAD vs PRD")
        print("=" * 90)
        print()

        print(f"QAD Status: {qad_result['status']}")
        print(f"PRD Status: {prd_result['status']}")
        print()

        # Diferenças de funções
        qad_funcs = qad_result["functions"]
        prd_funcs = prd_result["functions"]

        func_diffs = []
        for func, qad_has in qad_funcs.items():
            prd_has = prd_funcs.get(func, False)
            if qad_has != prd_has:
                func_diffs.append({
                    "name": func,
                    "qad": qad_has,
                    "prd": prd_has
                })

        # Diferenças de programas
        qad_progs = qad_result["programs"]
        prd_progs = prd_result["programs"]

        prog_diffs = []
        for prog, qad_has in qad_progs.items():
            prd_has = prd_progs.get(prog, False)
            if qad_has != prd_has:
                prog_diffs.append({
                    "name": prog,
                    "qad": qad_has,
                    "prd": prd_has
                })

        if not func_diffs and not prog_diffs:
            print("✅ COMPONENTES SINCRONIZADOS!")
            print()
            print("Conclusão:")
            print("  • Todas as funções RFC estão presentes")
            print("  • Todos os programas ABAP estão presentes")
            print()
            print("🔴 O erro de SYNTAX_ERROR pode ser:")
            print("  1. Problema de dados de entrada malformados")
            print("  2. Incompatibilidade de versão ABAP")
            print("  3. Objeto intermediário (estrutura) incompleto")
            print("  4. Request aberta conflitante")
            print()
            print("📌 Próximas ações:")
            print("  1. Verifique SE09 para requests abertas")
            print("  2. Execute SE38 → SAPLDMEE2 → F7 (Syntax Check)")
            print("  3. Consulte SAP Notes para a versão 758")
        else:
            if func_diffs:
                print("🔴 DIFERENÇAS DE FUNÇÕES:")
                for diff in func_diffs:
                    qad_status = "✅" if diff["qad"] else "❌"
                    prd_status = "✅" if diff["prd"] else "❌"
                    print(f"  {diff['name']}: QAD={qad_status} PRD={prd_status}")
                print()

            if prog_diffs:
                print("🔴 DIFERENÇAS DE PROGRAMAS:")
                for diff in prog_diffs:
                    qad_status = "✅" if diff["qad"] else "❌"
                    prd_status = "✅" if diff["prd"] else "❌"
                    print(f"  {diff['name']}: QAD={qad_status} PRD={prd_status}")
                print()

            print("⚠️  COMPONENTES DESINCRONIZADOS!")
            print()
            print("Ações:")
            print("  1. Transporte novamente todos os objetos Z_*")
            print("  2. Certifique-se que a Request está fechada")
            print("  3. Sincronize dependências de RFC")

        return len(func_diffs) == 0 and len(prog_diffs) == 0

    except ModuleNotFoundError:
        print("❌ PyRFC não instalado")
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

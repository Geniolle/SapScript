#!/usr/bin/env python3
"""
Compara estrutura Z_PT_CGI_XML_CT_V9 entre QAD e PRD via RFC (alternativo).
Usa DDIF_TABL_GET e DDIF_DTEL_GET.
"""

import os
import sys
from pathlib import Path
from typing import Dict, List

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

load_dotenv(project_root / ".env")


def get_table_definition(conn, table_name: str) -> Dict:
    """Recolhe definição de tabela/estrutura via DDIF_TABL_GET."""
    try:
        result = conn.call(
            "DDIF_TABL_GET",
            LANGU="E",
            PNAME=table_name,
            MODE="S",  # Sintaxe check
        )
        return result
    except Exception as e:
        print(f"    Erro DDIF_TABL_GET: {str(e)[:80]}")
        return {}


def get_dtel_definition(conn, dtel_name: str) -> Dict:
    """Recolhe definição de elemento de dados via DDIF_DTEL_GET."""
    try:
        result = conn.call(
            "DDIF_DTEL_GET",
            LANGU="E",
            PNAME=dtel_name,
            MODE="S",
        )
        return result
    except Exception as e:
        print(f"    Erro DDIF_DTEL_GET: {str(e)[:80]}")
        return {}


def get_struct_fields_alt(conn, struct_name: str) -> List[Dict]:
    """Recolhe campos de estrutura usando método alternativo."""
    try:
        # Tenta buscar como tabela (estruturas em SAP são tratadas como tabelas)
        result = get_table_definition(conn, struct_name)

        if result and result.get("DF") and result["DF"]:
            fields = []
            for field in result["DF"]:
                fields.append({
                    "FIELDNAME": field.get("FIELDNAME", "").strip(),
                    "POSITION": field.get("POSITION", 0),
                    "INTLEN": field.get("INTLEN", 0),
                    "DATATYPE": field.get("DATATYPE", "").strip(),
                    "LENG": field.get("LENG", 0),
                    "DECIMALS": field.get("DECIMALS", 0),
                })
            return sorted(fields, key=lambda x: x["POSITION"])

        return []
    except Exception as e:
        print(f"    Erro ao recolher campos: {str(e)[:80]}")
        return []


def compare_structures(qad_fields: List[Dict], prd_fields: List[Dict]) -> tuple:
    """Compara estruturas entre QAD e PRD."""
    qad_dict = {f["FIELDNAME"]: f for f in qad_fields}
    prd_dict = {f["FIELDNAME"]: f for f in prd_fields}

    only_qad = [f for f in qad_dict.keys() if f not in prd_dict]
    only_prd = [f for f in prd_dict.keys() if f not in qad_dict]

    differences = []
    for field in set(qad_dict.keys()) & set(prd_dict.keys()):
        q = qad_dict[field]
        p = prd_dict[field]
        if q["DATATYPE"] != p["DATATYPE"] or q["LENG"] != p["LENG"]:
            differences.append({
                "FIELD": field,
                "QAD": f"{q['DATATYPE']}({q['LENG']})",
                "PRD": f"{p['DATATYPE']}({p['LENG']})"
            })

    return only_qad, only_prd, differences


def main():
    print("=" * 85)
    print("🔍 COMPARADOR DE ESTRUTURAS Z_PT_CGI_XML_CT_V9 (QAD vs PRD)")
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

        # ========== QAD ==========
        print(f"📡 Conectando a QAD (172.19.66.22:00)...")
        qad_conn = Connection(**qad_config)
        print("   ✅ Conectado com sucesso")
        print()

        print(f"📋 Buscando estrutura '{struct_name}' em QAD...")
        qad_fields = get_struct_fields_alt(qad_conn, struct_name)
        if qad_fields:
            print(f"   ✅ Encontrados {len(qad_fields)} campos")
            print(f"      Campos: {', '.join([f['FIELDNAME'] for f in qad_fields[:5]])}")
            if len(qad_fields) > 5:
                print(f"      ... e mais {len(qad_fields) - 5}")
        else:
            print(f"   ⚠️  Nenhum campo encontrado em QAD")
        print()

        qad_conn.close()

        # ========== PRD ==========
        print(f"📡 Conectando a PRD (172.19.34.22:00)...")
        prd_conn = Connection(**prd_config)
        print("   ✅ Conectado com sucesso")
        print()

        print(f"📋 Buscando estrutura '{struct_name}' em PRD...")
        prd_fields = get_struct_fields_alt(prd_conn, struct_name)
        if prd_fields:
            print(f"   ✅ Encontrados {len(prd_fields)} campos")
            print(f"      Campos: {', '.join([f['FIELDNAME'] for f in prd_fields[:5]])}")
            if len(prd_fields) > 5:
                print(f"      ... e mais {len(prd_fields) - 5}")
        else:
            print(f"   ⚠️  Nenhum campo encontrado em PRD")
        print()

        prd_conn.close()

        # ========== COMPARAÇÃO ==========
        print("=" * 85)
        print("📊 RESULTADOS DA COMPARAÇÃO")
        print("=" * 85)
        print()

        if not qad_fields or not prd_fields:
            print("⚠️  Não foi possível recolher dados de ambos os ambientes.")
            print()
            print("Próximas ações:")
            print("  1. Verifique em SE11 (Data Dictionary) se a estrutura existe")
            print("  2. Confirme se o nome exacto é 'Z_PT_CGI_XML_CT_V9'")
            print("  3. Verifique se há autorização para ler a estrutura (RFC)")
            return False

        only_qad, only_prd, differences = compare_structures(qad_fields, prd_fields)

        print(f"📊 Resumo:")
        print(f"   QAD: {len(qad_fields)} campos")
        print(f"   PRD: {len(prd_fields)} campos")
        print()

        if not only_qad and not only_prd and not differences:
            print("✅ ESTRUTURAS IDÊNTICAS!")
            print()
            print("Conclusão: A estrutura Z_PT_CGI_XML_CT_V9 está sincronizada entre QAD e PRD.")
            print("O erro que viu pode ser de outro componente ou versão ABAP.")
            return True

        if only_qad:
            print(f"🔴 Campos APENAS em QAD ({len(only_qad)}):")
            for f in sorted(only_qad)[:10]:
                print(f"   - {f}")
            if len(only_qad) > 10:
                print(f"   ... e mais {len(only_qad) - 10}")
            print()

        if only_prd:
            print(f"🟠 Campos APENAS em PRD ({len(only_prd)}):")
            for f in sorted(only_prd)[:10]:
                print(f"   - {f}")
            if len(only_prd) > 10:
                print(f"   ... e mais {len(only_prd) - 10}")
            print()

        if differences:
            print(f"🟡 Campos com TIPOS DIFERENTES ({len(differences)}):")
            for d in differences[:10]:
                print(f"   - {d['FIELD']}: QAD={d['QAD']} vs PRD={d['PRD']}")
            if len(differences) > 10:
                print(f"   ... e mais {len(differences) - 10}")
            print()

        print("=" * 85)
        print("⚠️  DIFERENÇAS DETECTADAS")
        print("=" * 85)
        print()
        print("Ações recomendadas:")
        print("  1. Transporte de novo a estrutura (STRUC) de DEV para PRD")
        print("  2. Valide o transporte com: SE10 (Transportes)")
        print("  3. Sincronize todas as dependências (RFC, programas ABAP)")
        print()

        return False

    except ModuleNotFoundError:
        print("❌ PyRFC não instalado")
        print("Use: .venv-rfc\\Scripts\\python.exe")
        return False
    except Exception as e:
        print(f"❌ Erro: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

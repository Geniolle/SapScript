#!/usr/bin/env python3
"""
Compara estrutura Z_PT_CGI_XML_CT_V9 entre QAD e PRD via RFC.
Modo leitura apenas.
"""

import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple

project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

load_dotenv(project_root / ".env")


def get_struct_fields(conn, struct_name: str) -> Dict:
    """Recolhe informações sobre uma estrutura ABAP via RFC_DDIC_TABLE_GET."""
    try:
        result = conn.call(
            "DDIF_DTEL_GET",
            LANGU="E",
            PNAME=struct_name
        )
        return result
    except Exception as e:
        print(f"  ❌ Erro ao buscar estructura: {str(e)[:100]}")
        return {}


def get_struct_components(conn, struct_name: str) -> List[Dict]:
    """Recolhe componentes de uma estrutura via RFC."""
    try:
        # Tenta via DDIF_STRUCT_GET
        result = conn.call(
            "DDIF_STRUCT_GET",
            LANGU="E",
            PNAME=struct_name
        )

        if result.get("DFIES"):
            components = []
            for field in result["DFIES"]:
                components.append({
                    "FIELDNAME": field.get("FIELDNAME", "").strip(),
                    "POSITION": field.get("POSITION", 0),
                    "INTLEN": field.get("INTLEN", 0),
                    "DATATYPE": field.get("DATATYPE", "").strip(),
                    "LENG": field.get("LENG", 0),
                    "DECIMALS": field.get("DECIMALS", 0),
                })
            return sorted(components, key=lambda x: x["POSITION"])
        return []
    except Exception as e:
        print(f"  ⚠️  RFC DDIF_STRUCT_GET falhou: {str(e)[:60]}")
        return []


def get_struct_info_via_rfc(conn, struct_name: str) -> Dict:
    """Recolhe informações gerais sobre uma estrutura."""
    try:
        result = conn.call(
            "DDIF_DTEL_GET",
            LANGU="E",
            PNAME=struct_name
        )
        if result.get("DD01"):
            dd01 = result["DD01"]
            return {
                "NAME": dd01.get("TABNAME", "").strip(),
                "VERSION": dd01.get("VERSNO", 0),
                "LASTMOD": dd01.get("LASTEDITBY", "").strip(),
                "TIMESTAMP": dd01.get("TIMESTAMP", "").strip(),
            }
    except:
        pass
    return {}


def compare_envs(qad_components: List[Dict], prd_components: List[Dict]) -> Tuple[List, List, List]:
    """Compara componentes entre QAD e PRD."""
    qad_fields = {c["FIELDNAME"]: c for c in qad_components}
    prd_fields = {c["FIELDNAME"]: c for c in prd_components}

    # Campos apenas em QAD
    only_qad = [f for f in qad_fields.keys() if f not in prd_fields]

    # Campos apenas em PRD
    only_prd = [f for f in prd_fields.keys() if f not in qad_fields]

    # Campos com diferenças
    differences = []
    for field in set(qad_fields.keys()) & set(prd_fields.keys()):
        qad = qad_fields[field]
        prd = prd_fields[field]
        if qad.get("DATATYPE") != prd.get("DATATYPE") or qad.get("LENG") != prd.get("LENG"):
            differences.append({
                "FIELD": field,
                "QAD_TYPE": f"{qad.get('DATATYPE')}({qad.get('LENG')})",
                "PRD_TYPE": f"{prd.get('DATATYPE')}({prd.get('LENG')})"
            })

    return only_qad, only_prd, differences


def main():
    print("=" * 80)
    print("🔍 COMPARADOR DE ESTRUTURAS ABAP: Z_PT_CGI_XML_CT_V9")
    print("=" * 80)
    print()

    # Carrega credenciais
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
        print(f"📡 Conectando a QAD (172.19.66.22)...")
        qad_conn = Connection(**qad_config)
        print("  ✅ Conectado a QAD")
        print()

        print(f"  📋 Buscando estrutura '{struct_name}' em QAD...")
        qad_components = get_struct_components(qad_conn, struct_name)
        if qad_components:
            print(f"  ✅ Encontrados {len(qad_components)} campos em QAD")
        else:
            print(f"  ❌ Estrutura não encontrada em QAD ou sem componentes")

        qad_conn.close()
        print()

        # ========== PRD ==========
        print(f"📡 Conectando a PRD (172.19.34.22)...")
        prd_conn = Connection(**prd_config)
        print("  ✅ Conectado a PRD")
        print()

        print(f"  📋 Buscando estrutura '{struct_name}' em PRD...")
        prd_components = get_struct_components(prd_conn, struct_name)
        if prd_components:
            print(f"  ✅ Encontrados {len(prd_components)} campos em PRD")
        else:
            print(f"  ❌ Estrutura não encontrada em PRD ou sem componentes")

        prd_conn.close()
        print()

        # ========== COMPARAÇÃO ==========
        print("=" * 80)
        print("📊 RESULTADOS DA COMPARAÇÃO")
        print("=" * 80)
        print()

        if not qad_components or not prd_components:
            print("❌ Não foi possível comparar - faltam dados de um ou ambos os ambientes")
            return False

        only_qad, only_prd, differences = compare_envs(qad_components, prd_components)

        # Resumo
        print(f"Total de campos em QAD: {len(qad_components)}")
        print(f"Total de campos em PRD: {len(prd_components)}")
        print()

        # Campos apenas em QAD
        if only_qad:
            print(f"🔴 CAMPOS APENAS EM QAD ({len(only_qad)}):")
            for field in sorted(only_qad):
                print(f"   - {field}")
            print()

        # Campos apenas em PRD
        if only_prd:
            print(f"🟠 CAMPOS APENAS EM PRD ({len(only_prd)}):")
            for field in sorted(only_prd):
                print(f"   - {field}")
            print()

        # Diferenças de tipo
        if differences:
            print(f"🟡 CAMPOS COM TIPOS DIFERENTES ({len(differences)}):")
            print()
            for diff in differences:
                print(f"   Campo: {diff['FIELD']}")
                print(f"     QAD: {diff['QAD_TYPE']}")
                print(f"     PRD: {diff['PRD_TYPE']}")
                print()

        # Status final
        print("=" * 80)
        if not only_qad and not only_prd and not differences:
            print("✅ ESTRUTURAS IDÊNTICAS ENTRE QAD E PRD!")
        else:
            print("⚠️  DIFERENÇAS DETECTADAS ENTRE QAD E PRD")
            if only_qad or only_prd:
                print("   → Sincronize a estrutura completa (STRUC)")
            if differences:
                print("   → Verifique compatibilidade de tipos de dados")
        print()

        return len(only_qad) == 0 and len(only_prd) == 0 and len(differences) == 0

    except ModuleNotFoundError:
        print("❌ PyRFC não instalado. Use: .venv-rfc\\Scripts\\python.exe")
        return False
    except Exception as e:
        print(f"❌ Erro geral: {str(e)}")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

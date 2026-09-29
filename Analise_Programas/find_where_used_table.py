"""
Investigação: Todos os objetos SAP que utilizam ZFI_DOC_EX_LOG_T

Objetivo: Obter a Where-Used List (lista de utilizações) da tabela
no repositório ABAP do SAP QAD via RFC.

Modo: SOMENTE LEITURA
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

try:
    from pyrfc import Connection
except Exception as exc:
    print(f"ERRO: PyRFC não disponível. {exc}")
    sys.exit(1)

# Setup
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)


class WhereUsedFinder:
    """Encontra utilizações de tabelas/objetos no SAP."""

    def __init__(self):
        self.conn = None

    def connect(self):
        """Conecta ao SAP QAD."""
        env_prefix = "QAD"
        params = {
            "user": os.getenv(f"SAP_{env_prefix}_USER", "").strip(),
            "passwd": os.getenv(f"SAP_{env_prefix}_PASSWD", "").strip(),
            "ashost": os.getenv(f"SAP_{env_prefix}_ASHOST", "").strip(),
            "sysnr": os.getenv(f"SAP_{env_prefix}_SYSNR", "").strip(),
            "client": os.getenv(f"SAP_{env_prefix}_CLIENT", "").strip(),
        }

        self.conn = Connection(**params)
        self.conn.call("RFC_PING")
        print("✓ Conectado a S4Q/100")

    def read_table(self, table: str, fields: list[str], options: list[str] = None) -> list[dict[str, Any]]:
        """Lê tabela via RFC."""
        try:
            result = self.conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE=table,
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": f} for f in fields],
                OPTIONS=[{"TEXT": o} for o in (options or [])],
                ROWCOUNT=0,
            )

            rows = []
            for item in result.get("DATA", []) or []:
                parts = [p.strip() for p in str(item.get("WA", "")).split("|")]
                rows.append({field: (parts[i] if i < len(parts) else "") for i, field in enumerate(fields)})

            return rows
        except Exception as e:
            if "TABLE_WITHOUT_DATA" in str(e).upper():
                return []
            raise RuntimeError(f"Erro em {table}: {e}") from e

    def find_where_used(self):
        """Procura onde ZFI_DOC_EX_LOG_T é utilizada."""

        print("\n" + "="*100)
        print("INVESTIGAÇÃO: ONDE-UTILIZADA (ZFI_DOC_EX_LOG_T)")
        print("="*100)

        # Tentativa 1: Tabela TADIR (Dynamic Repository)
        print("\n[TENTATIVA 1] Consultar TADIR (Dynamic Repository)")
        print("-" * 100)

        try:
            tadir_fields = ["PGMID", "OBJECT", "OBJ_NAME", "AUTHOR", "CREATED_ON", "CHANGED_ON"]
            tadir_rows = self.read_table(
                "TADIR",
                tadir_fields,
                [
                    "OBJECT IN ('PROG', 'REPT', 'FUNC', 'CLAS', 'INTF')",
                    "OBJ_NAME LIKE 'Z%'"  # Apenas objetos Z
                ]
            )

            print(f"Encontrados {len(tadir_rows)} objetos Z* no repositório")

            # Agora pesquisar por referências a ZFI_DOC_EX_LOG_T
            # Isto requeria ler o código-fonte, o que RFC_READ_REPORT não permitiu
            print("⚠️  Nota: Verificação de referências requeria leitura de código-fonte")
            print("    (RFC_READ_REPORT não está disponível neste ambiente)")

        except Exception as e:
            print(f"✗ Erro: {e}")

        # Tentativa 2: Tabela DRAD (Dynamic Repository Table Access)
        print("\n[TENTATIVA 2] Tabela DRAD (Table References)")
        print("-" * 100)

        try:
            drad_fields = ["PROG", "TNAME", "ATYP", "ANAME"]
            drad_rows = self.read_table(
                "DRAD",
                drad_fields,
                ["TNAME = 'ZFI_DOC_EX_LOG_T'"]
            )

            if drad_rows:
                print(f"✓ Encontradas {len(drad_rows)} referências em DRAD:")
                for row in drad_rows:
                    print(f"  Programa: {row.get('PROG')}")
                    print(f"  Tipo: {row.get('ATYP')}")
                    print(f"  Nome: {row.get('ANAME')}")
                    print()
            else:
                print("✗ Nenhuma referência encontrada em DRAD")

        except Exception as e:
            print(f"⚠️  Erro ao consultar DRAD: {e}")
            print("    (Tabela pode não estar disponível)")

        # Tentativa 3: Tabela D010S (Source Code Repository)
        print("\n[TENTATIVA 3] Tabela D010S (Source Code)")
        print("-" * 100)

        try:
            d010s_fields = ["PROG", "CNAM", "CCNT"]
            d010s_rows = self.read_table(
                "D010S",
                d010s_fields,
                []  # Sem filtro, para ver o padrão
            )

            if d010s_rows:
                print(f"✓ Tabela D010S acessível (contém {len(d010s_rows)} blocos)")
                print("  ⚠️  Mas busca por 'ZFI_DOC_EX_LOG_T' requeria grep no conteúdo")
            else:
                print("⚠️  D010S existe mas está vazia ou inacessível")

        except Exception as e:
            print(f"✗ Erro ao consultar D010S: {e}")

        # Tentativa 4: Procurar programas Z que possam usar a tabela
        print("\n[TENTATIVA 4] Programas Z* que começam com 'ZFI'")
        print("-" * 100)

        try:
            tadir_fields = ["PGMID", "OBJECT", "OBJ_NAME"]
            tadir_rows = self.read_table(
                "TADIR",
                tadir_fields,
                [
                    "OBJECT = 'PROG'",
                    "OBJ_NAME LIKE 'ZFI%'"
                ]
            )

            if tadir_rows:
                print(f"\n✓ Encontrados {len(tadir_rows)} programas ZFI*:")
                for row in tadir_rows[:20]:  # Primeiros 20
                    print(f"  - {row.get('OBJ_NAME')}")

                if len(tadir_rows) > 20:
                    print(f"  ... e mais {len(tadir_rows) - 20}")
            else:
                print("✗ Nenhum programa ZFI* encontrado")

        except Exception as e:
            print(f"✗ Erro: {e}")

        # Tentativa 5: Classes que usem ZFI_DOC_EX_LOG_T
        print("\n[TENTATIVA 5] Classes Z* que possam usar ZFI_DOC_EX_LOG_T")
        print("-" * 100)

        try:
            tadir_fields = ["PGMID", "OBJECT", "OBJ_NAME"]
            tadir_rows = self.read_table(
                "TADIR",
                tadir_fields,
                [
                    "OBJECT = 'CLAS'",
                    "OBJ_NAME LIKE 'Z%'"
                ]
            )

            if tadir_rows:
                print(f"\n✓ Encontradas {len(tadir_rows)} classes Z*:")
                for row in tadir_rows[:10]:  # Primeiros 10
                    print(f"  - {row.get('OBJ_NAME')}")

                if len(tadir_rows) > 10:
                    print(f"  ... e mais {len(tadir_rows) - 10}")
            else:
                print("✗ Nenhuma classe Z* encontrada")

        except Exception as e:
            print(f"✗ Erro: {e}")

        # Conclusão
        print("\n" + "="*100)
        print("CONCLUSÃO")
        print("="*100)

        print("\n⚠️  LIMITAÇÃO RFC")
        print("  As funções RFC standard (RFC_READ_REPORT, READ_REPOSITORY) não estão")
        print("  disponíveis neste ambiente SAP QAD para leitura de código-fonte.")
        print("\n  Para obter a Where-Used List completa, é necessário:")
        print("  1. Usar SE11 → ZFI_DOC_EX_LOG_T → Utilizações")
        print("  2. Ou usar ABAP Query / Dynamic Repository")
        print("  3. Ou usar ferramentas como SE80 (Object Navigator)")

    def close(self):
        """Fecha conexão."""
        if self.conn:
            try:
                self.conn.close()
            except:
                pass


def main():
    finder = None
    try:
        finder = WhereUsedFinder()
        finder.connect()
        finder.find_where_used()
    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if finder:
            finder.close()


if __name__ == "__main__":
    main()

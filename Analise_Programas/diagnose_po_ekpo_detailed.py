"""
Investigação Detalhada: Por que EKPO não retorna posições para PO 4300000002?

Achado crítico: A PO existe em EKKO mas não tem posições em EKPO.
Isto pode ser a CAUSA RAIZ de não aparecer em AUTO_INST_ASSIGN!
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


class EKPOInvestigation:
    """Investiga EKPO em detalhes."""

    def __init__(self):
        self.conn = None
        self.po_number = "4300000002"

    def connect(self):
        """Conecta ao SAP."""
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

    def investigate(self):
        """Executa investigação."""

        print("\n" + "="*80)
        print("INVESTIGAÇÃO DETALHADA: EKPO VAZIO?")
        print("="*80)

        # Teste 1: Verificar EKPO com apenas EBELN
        print("\n--- TESTE 1: Consultar EKPO com EBELN = '4300000002' ---")
        print(f"Condição WHERE: EBELN = '{self.po_number}'")

        ekpo_fields_basic = ["EBELN", "EBELP"]
        result = self.read_table("EKPO", ekpo_fields_basic, [f"EBELN = '{self.po_number}'"])

        print(f"Resultado: {len(result)} linhas encontradas")
        if result:
            for row in result:
                print(f"  EBELN: {row.get('EBELN')}, EBELP: {row.get('EBELP')}")

        # Teste 2: Consultar EKPO com campos expandidos
        print("\n--- TESTE 2: Consultando EKPO com campos extendidos ---")

        ekpo_fields_extended = [
            "EBELN",    # Número do documento
            "EBELP",    # Número da posição
            "LOEKZ",    # Indicador de exclusão
            "PSTYP",    # Categoria de posição
            "EVERS",    # Status de entrega
            "MENGE",    # Quantidade
            "MEINS",    # Unidade de medida
        ]

        result_extended = self.read_table("EKPO", ekpo_fields_extended, [f"EBELN = '{self.po_number}'"])

        print(f"Resultado: {len(result_extended)} linhas encontradas")
        for row in result_extended:
            print(f"  Posição {row.get('EBELP')}: LOEKZ={row.get('LOEKZ')}, PSTYP={row.get('PSTYP')}, Qtd={row.get('MENGE')}")

        # Teste 3: Busca de POs similares (que começam com 43)
        print("\n--- TESTE 3: Verificar outras POs com padrão 43* ---")
        print("Buscando POs que começam com '43'...")

        try:
            po_sample = self.read_table("EKPO", ["EBELN", "EBELP"], [], rowcount=100)
            pos_with_43 = [r for r in po_sample if r.get('EBELN', '').startswith('43')]

            if pos_with_43:
                print(f"✓ Encontradas {len(pos_with_43)} posições em POs que começam com '43':")
                for row in pos_with_43[:5]:
                    print(f"  {row.get('EBELN')}: posição {row.get('EBELP')}")
            else:
                print("✗ Nenhuma posição encontrada em POs que começam com '43'")
        except Exception as e:
            print(f"Erro: {e}")

        # Teste 4: Verificar se a PO pode estar em uma tabela de histórico
        print("\n--- TESTE 4: Verificar dados adicionais de EKKO ---")

        ekko_fields_extra = [
            "EBELN",
            "BUKRS",
            "BSTYP",    # Tipo de pedido
            "BSART",    # Tipo de documento
            "LOEKZ",    # Indicador de exclusão
            "STATU",    # Status geral
            "AEDAT",    # Data de alteração
        ]

        ekko_result = self.read_table("EKKO", ekko_fields_extra, [f"EBELN = '{self.po_number}'"])

        if ekko_result:
            ekko = ekko_result[0]
            print(f"Status da PO em EKKO:")
            print(f"  BSTYP (tipo pedido): {ekko.get('BSTYP')}")
            print(f"  BSART (tipo doc): {ekko.get('BSART')}")
            print(f"  LOEKZ (exclusão): '{ekko.get('LOEKZ', 'VAZIO')}'")
            print(f"  STATU (status): {ekko.get('STATU')}")

        # Teste 5: Contar total de registos em EKPO para esta PO
        print("\n--- TESTE 5: Contar registos em EKPO ---")

        try:
            # Tenta ler com COUNT
            raw_result = self.conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="EKPO",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": "EBELN"}, {"FIELDNAME": "EBELP"}],
                OPTIONS=[{"TEXT": f"EBELN = '{self.po_number}'"}],
            )

            raw_data = raw_result.get("DATA", [])
            print(f"RFC retornou: {len(raw_data)} linhas (raw)")

            if len(raw_data) > 0:
                print("Primeiras 3 linhas raw:")
                for i, item in enumerate(raw_data[:3]):
                    print(f"  {i+1}: {item}")

        except Exception as e:
            print(f"Erro na leitura raw: {e}")

        # Conclusão
        print("\n" + "="*80)
        print("CONCLUSÃO DA INVESTIGAÇÃO EKPO")
        print("="*80)

        if not result_extended:
            print("\n✗ ACHADO CRÍTICO: A PO 4300000002 NÃO TEM POSIÇÕES EM EKPO!")
            print("\nIsso significa:")
            print("  1. A PO foi criada mas não tem linhas de material")
            print("  2. AUTO_INST_ASSIGN procura por POs com posições")
            print("  3. Uma PO sem posições é AUTOMATICAMENTE EXCLUÍDA")
            print("\nPOR ISSO A PO NÃO APARECE NO AUTO_INST_ASSIGN:")
            print("  ► A PO 4300000002 não tem registos em EKPO")
        else:
            print(f"\n✓ PO tem {len(result_extended)} posição(ões)")
            print("  Investigar outras condições de filtro")

    def close(self):
        """Fecha conexão."""
        if self.conn:
            try:
                self.conn.close()
            except:
                pass


def main():
    diag = None
    try:
        diag = EKPOInvestigation()
        diag.connect()
        print("✓ Conectado a S4Q/100")
        diag.investigate()
    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if diag:
            diag.close()


if __name__ == "__main__":
    main()

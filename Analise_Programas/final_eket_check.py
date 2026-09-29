"""
Verificação Final: EKET para PO 4300000002

Objetivo: Confirmar se a PO tem registos em EKET
(para validar se INNER JOIN EKET causa a exclusão)

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


class EKETCheck:
    """Verifica EKET para PO 4300000002."""

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
                ROWCOUNT=0,  # Sem limite
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

    def check_eket(self):
        """Consulta EKET para a PO 4300000002."""

        print("\n" + "="*100)
        print("VERIFICAÇÃO FINAL: EKET PARA PO 4300000002")
        print("="*100)

        print("\nConsultando EKET...")
        print("Tabela: EKET")
        print("Condição: EBELN = '4300000002'")
        print("Sem limite de resultados")
        print("-" * 100)

        # Campos solicitados
        eket_fields = [
            "EBELN",    # Número do documento de compras
            "EBELP",    # Número da posição
            "ETENR",    # Número sequencial do termo
            "EINDT",    # Data de entrega prevista
            "MENGE",    # Quantidade
            "WEMNG",    # Quantidade recebida/faturada
        ]

        try:
            eket_rows = self.read_table(
                "EKET",
                eket_fields,
                ["EBELN = '4300000002'"]
            )

            # Resultado
            print(f"\n{'RESULTADO:':<50}")
            print(f"{'='*100}")

            if eket_rows:
                print(f"\n✓ EKET encontrados: {len(eket_rows)} linha(s)\n")

                # Cabeçalho da tabela
                print(f"{'EBELN':<15} | {'EBELP':<8} | {'ETENR':<8} | {'EINDT':<12} | {'MENGE':<15} | {'WEMNG':<15}")
                print("-" * 100)

                # Dados
                for row in eket_rows:
                    print(f"{row.get('EBELN', 'N/A'):<15} | {row.get('EBELP', 'N/A'):<8} | {row.get('ETENR', 'N/A'):<8} | {row.get('EINDT', 'N/A'):<12} | {row.get('MENGE', 'N/A'):<15} | {row.get('WEMNG', 'N/A'):<15}")

                print("\nDetalhes completos de cada linha:")
                for i, row in enumerate(eket_rows, 1):
                    print(f"\nLinha {i}:")
                    for field in eket_fields:
                        print(f"  {field}: {row.get(field, 'N/A')}")

            else:
                print(f"\n✗ EKET encontrados: 0 linhas\n")

            # Conclusão
            print("\n" + "="*100)
            print("CONCLUSÃO FINAL")
            print("="*100)

            if not eket_rows:
                print("\n🎯 CAUSA CONFIRMADA!")
                print("\n━" * 50)
                print("A PO 4300000002 é ELIMINADA pelo INNER JOIN com EKET")
                print("no método AUTO_INST_ASSIGN porque NÃO possui correspondência em EKET.")
                print("━" * 50)

                print("\nPROVA ABAP REAL:")
                print("""
Ficheiro: scratch\\ZFI_PURCH_DOC_EX_RATE_LCL.abap
Método: AUTO_INST_ASSIGN (linhas 304-325)
Trecho:

    SELECT A~BUKRS,
           A~WAERS,
           A~EBELN,
           B~EINDT,
           ...
    FROM EKKO AS A
    INNER JOIN EKET AS B ON B~EBELN EQ A~EBELN  ← INNER JOIN crítico
    ...
    INTO TABLE @DATA(LT_AUTO)
    WHERE A~BUKRS EQ @IV_BUKRS
    AND A~WAERS EQ @IV_WAERS
    ...

Explicação:
- INNER JOIN EKET retorna ZERO linhas se não há registos em EKET
- Logo, a PO 4300000002 é automaticamente excluída do resultado
- Isto acontece ANTES de qualquer outro filtro WHERE
- Não é um filtro explícito, mas uma exclusão implícita pela lógica do JOIN
                """)

                print("\n✅ INVESTIGAÇÃO CONCLUÍDA COM SUCESSO")
                print("   Causa raiz: INNER JOIN EKET (linha 314 do código ABAP)")

                return 0

            else:
                print(f"\n⚠️  EKET NÃO é a causa")
                print(f"\n{len(eket_rows)} registo(s) encontrado(s) em EKET")
                print("\nContinuando investigação para identificar outro filtro...")
                print("\nPróximos passos:")
                print("1. Testar JOIN com LFB1 (dados de fornecedor)")
                print("2. Testar JOIN com ZFI_PAY_DATE_T")
                print("3. Testar JOIN com T052 (condições de pagamento)")
                print("4. Testar JOIN com EKPO (posições)")

                return 1

        except Exception as e:
            print(f"\n✗ ERRO ao consultar EKET: {e}")
            import traceback
            traceback.print_exc()
            return -1

    def close(self):
        """Fecha conexão."""
        if self.conn:
            try:
                self.conn.close()
            except:
                pass


def main():
    checker = None
    try:
        checker = EKETCheck()
        checker.connect()
        result = checker.check_eket()
        sys.exit(result)
    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(-1)
    finally:
        if checker:
            checker.close()


if __name__ == "__main__":
    main()

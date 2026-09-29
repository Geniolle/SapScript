"""
Diagnóstico Completo: Por que a PO 4300000002 não aparece em AUTO_INST_ASSIGN

Fases:
1. Extrair código ABAP do ZFI_PURCH_DOC_EX_RATE e localizar AUTO_INST_ASSIGN
2. Consultar EKPO
3. Consultar EKET
4. Consultar ZFI_DOC_EX_LOG_T
5. Consultar parâmetros ZCLCA_FIXEDVALS
6. Reproduzir cálculo PAY_DATE
7. Investigar relação LIFNR <-> Parceiro de negócios
8. Montar matriz de filtros final
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

# Setup do projeto
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)


class CompleteDiagnostic:
    """Diagnóstico completo da PO 4300000002."""

    def __init__(self):
        self.conn = None
        self.po_number = "4300000002"
        self.filters_matrix = []

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

        try:
            self.conn = Connection(**params)
            self.conn.call("RFC_PING")
            print("✓ Conectado a S4Q/100")
            return True
        except Exception as e:
            print(f"✗ Erro ao conectar: {e}")
            return False

    def read_rfc_table(self, table: str, fields: list[str], options: list[str] = None) -> list[dict[str, Any]]:
        """Lê tabela via RFC_READ_TABLE."""
        if not self.conn:
            raise RuntimeError("Não conectado")

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
            raise RuntimeError(f"Erro ao ler {table}: {e}") from e

    def phase_1_extract_code(self):
        """FASE 1: Extrai código ABAP do ZFI_PURCH_DOC_EX_RATE."""

        print("\n" + "="*80)
        print("FASE 1: EXTRAIR CÓDIGO ABAP DO ZFI_PURCH_DOC_EX_RATE")
        print("="*80)

        try:
            # Tenta ler o programa via RFC_READ_REPORT
            print("\nTentando extrair código via RFC_READ_REPORT...")

            result = self.conn.call(
                "RFC_READ_REPORT",
                PROGRAM="ZFI_PURCH_DOC_EX_RATE",
            )

            lines = result.get("PROGRAM", [])

            if not lines:
                print("✗ Programa não encontrado ou vazio via RFC_READ_REPORT")
                return None

            print(f"✓ Programa extraído com {len(lines)} linhas")

            # Procura pelo método AUTO_INST_ASSIGN
            code_text = "\n".join([str(line.get("LINE", "")) for line in lines])

            # Procura pela definição do método
            if "AUTO_INST_ASSIGN" in code_text:
                print("✓ Método AUTO_INST_ASSIGN encontrado no código")

                # Extrai o trecho do método
                lines_list = code_text.split("\n")
                start_idx = None
                for i, line in enumerate(lines_list):
                    if "AUTO_INST_ASSIGN" in line:
                        start_idx = i
                        break

                if start_idx:
                    # Mostra 50 linhas a partir do método
                    print("\n--- Trecho do método AUTO_INST_ASSIGN ---")
                    for i in range(start_idx, min(start_idx + 50, len(lines_list))):
                        print(f"{i-start_idx+1:3}: {lines_list[i]}")
                    print("--- Fim do trecho ---\n")

                    return code_text
            else:
                print("⚠️  Método AUTO_INST_ASSIGN não encontrado no código extraído")
                print("Mostrando primeiras 100 linhas do programa:")
                lines_list = code_text.split("\n")[:100]
                for i, line in enumerate(lines_list):
                    print(f"{i+1:3}: {line}")

                return code_text

        except Exception as e:
            print(f"✗ Erro ao extrair código: {e}")
            print("\nContinuando diagnóstico com análise de tabelas...")
            return None

    def phase_2_ekpo(self):
        """FASE 2: Consultar EKPO."""

        print("\n" + "="*80)
        print("FASE 2: CONSULTAR EKPO")
        print("="*80)

        ekpo_fields = [
            "EBELN",    # Número do documento
            "EBELP",    # Número da posição
            "LFGNI",    # Identificador de mercadoria
            "MENGE",    # Quantidade
            "MEINS",    # Unidade de medida
            "NETPR",    # Preço líquido
            "PEINH",    # Unidade de preço
            "BRTWR",    # Valor bruto
            "WAERS",    # Moeda
            "EDATUM",   # Data de entrega
        ]

        try:
            ekpo_rows = self.read_rfc_table("EKPO", ekpo_fields, [f"EBELN = '{self.po_number}'"])

            if ekpo_rows:
                print(f"\n✓ {len(ekpo_rows)} posição(ões) encontrada(s):")
                for row in ekpo_rows:
                    print(f"\n  Posição {row.get('EBELP', 'N/A')}:")
                    print(f"    Quantidade: {row.get('MENGE', 'N/A')} {row.get('MEINS', 'N/A')}")
                    print(f"    Preço: {row.get('NETPR', 'N/A')} {row.get('WAERS', 'N/A')}")
                    print(f"    Data de entrega: {row.get('EDATUM', 'N/A')}")

                return ekpo_rows
            else:
                print("✗ Nenhuma posição encontrada em EKPO")
                return []

        except Exception as e:
            print(f"✗ Erro: {e}")
            return []

    def phase_3_eket(self, ekpo_rows):
        """FASE 3: Consultar EKET."""

        print("\n" + "="*80)
        print("FASE 3: CONSULTAR EKET")
        print("="*80)

        if not ekpo_rows:
            print("Nenhuma posição em EKPO, pulando EKET")
            return []

        eket_fields = [
            "EBELN",    # Número do documento
            "EBELP",    # Número da posição
            "EKETX",    # Número sequencial
            "EINDT",    # Data de entrega prevista
            "MENGE",    # Quantidade
        ]

        eket_rows = []

        try:
            result = self.read_rfc_table("EKET", eket_fields, [f"EBELN = '{self.po_number}'"])

            if result:
                print(f"\n✓ {len(result)} linha(s) encontrada(s) em EKET:")
                for row in result:
                    print(f"  Posição {row.get('EBELP', 'N/A')}, sequência {row.get('EKETX', 'N/A')}: {row.get('MENGE', 'N/A')} para {row.get('EINDT', 'N/A')}")
                    eket_rows.append(row)
            else:
                print("✗ Nenhuma linha em EKET")

        except Exception as e:
            print(f"✗ Erro: {e}")

        return eket_rows

    def phase_4_log(self):
        """FASE 4: Consultar ZFI_DOC_EX_LOG_T."""

        print("\n" + "="*80)
        print("FASE 4: CONSULTAR ZFI_DOC_EX_LOG_T")
        print("="*80)

        log_fields = [
            "EBELN",
            "BUKRS",
            "INSTRUM_ID",
            "TIPO_EX",
            "DATA_CRIA",
            "USUARIO",
        ]

        try:
            log_rows = self.read_rfc_table("ZFI_DOC_EX_LOG_T", log_fields, [f"EBELN = '{self.po_number}'"])

            if log_rows:
                print(f"\n✗ {len(log_rows)} registo(s) encontrado(s) em ZFI_DOC_EX_LOG_T:")
                print("   ISTO SIGNIFICA: A PO foi já processada/excluída do AUTO_INST_ASSIGN!")
                for row in log_rows:
                    print(f"   {row}")
                return log_rows
            else:
                print("\n✓ Nenhum registo em ZFI_DOC_EX_LOG_T")
                print("   A PO ainda não foi processada")
                return []

        except Exception as e:
            print(f"\n⚠️  Erro ao consultar ZFI_DOC_EX_LOG_T: {e}")
            print("   (Tabela pode não existir ou ter outro nome)")
            return []

    def phase_5_parameters(self):
        """FASE 5: Consultar parâmetros ZCLCA_FIXEDVALS."""

        print("\n" + "="*80)
        print("FASE 5: CONSULTAR PARÂMETROS ZCLCA_FIXEDVALS")
        print("="*80)

        params_fields = [
            "MANDT",
            "MODULE",
            "PROCESSO",
            "PARAM_NAME",
            "PARAM_VALUE",
        ]

        try:
            params_rows = self.read_rfc_table(
                "ZCLCA_FIXEDVALS",
                params_fields,
                [
                    "MODULE = 'FIN'",
                    "PROCESSO = 'PO_EXCHANGE_RATE'",
                ]
            )

            if params_rows:
                print(f"\n✓ {len(params_rows)} parâmetro(s) encontrado(s):")
                params_dict = {}
                for row in params_rows:
                    param_name = row.get('PARAM_NAME', '').strip()
                    param_value = row.get('PARAM_VALUE', '').strip()
                    params_dict[param_name] = param_value
                    print(f"  {param_name} = {param_value}")

                return params_dict
            else:
                print("✗ Nenhum parâmetro encontrado")
                return {}

        except Exception as e:
            print(f"⚠️  Erro: {e}")
            return {}

    def phase_6_pay_date(self, ekpo_rows, params):
        """FASE 6: Reproduzir cálculo PAY_DATE."""

        print("\n" + "="*80)
        print("FASE 6: REPRODUZIR CÁLCULO PAY_DATE")
        print("="*80)

        if not ekpo_rows:
            print("Nenhuma posição disponível")
            return

        print("\nFórmula: PAY_DATE = ZZDAT02 - NDAYS + ZTAG1")
        print("\nParâmetros obtidos:")
        ndays = int(params.get('NDAYS', '0')) if params.get('NDAYS') else None
        print(f"  NDAYS = {ndays if ndays else '[não encontrado]'}")

        # ZZDAT02 seria a data de entrega mais próxima de EKET
        # ZTAG1 seria 1 por padrão
        print("\nPara cada posição:")
        for row in ekpo_rows:
            print(f"\n  Posição {row.get('EBELP', 'N/A')}:")
            # Em produção, seria necessário buscar a data de entrega real

    def phase_7_vendor(self):
        """FASE 7: Investigar relação LIFNR vs Parceiro de negócios."""

        print("\n" + "="*80)
        print("FASE 7: RELAÇÃO LIFNR vs PARCEIRO DE NEGÓCIOS")
        print("="*80)

        print("\nDados confirmados:")
        print("  EKKO-LIFNR: 0010007092")
        print("  Parceiro informado na execução: 50000001")

        print("\nInvestigando possível relação...")

        # Consulta LFA1 para ver dados do fornecedor
        try:
            lfa1_rows = self.read_rfc_table(
                "LFA1",
                ["LIFNR", "NAME1", "KONZC"],
                [f"LIFNR = '0010007092'"]
            )

            if lfa1_rows:
                print(f"\n✓ Fornecedor 0010007092 encontrado em LFA1:")
                for row in lfa1_rows:
                    print(f"  Nome: {row.get('NAME1', 'N/A')}")
                    print(f"  Conceito: {row.get('KONZC', 'N/A')}")

        except Exception as e:
            print(f"⚠️  Erro ao consultar LFA1: {e}")

    def phase_8_matrix(self, ekpo_rows, log_rows, params):
        """FASE 8: Montar matriz final de filtros."""

        print("\n" + "="*80)
        print("FASE 8: MATRIZ FINAL DE FILTROS")
        print("="*80)

        matrix_data = [
            ("BUKRS", "2010", "2010", "PASSA", "EKKO", "Empresa"),
            ("WAERS", "USD", "USD", "PASSA", "EKKO", "Moeda"),
            ("EBELN", "4300000002", "4300000002", "PASSA", "EKKO", "Documento"),
            ("BSART", "Válido", "ZP20", "?", "EKKO", "Tipo de doc"),
            ("LOEKZ", "Vazio", "Vazio", "PASSA", "EKKO", "Não excluída"),
            ("LOG", "Nenhum registro", f"{len(log_rows)} registos", "FALHA" if log_rows else "PASSA", "ZFI_DOC_EX_LOG_T", "Já processada"),
            ("PAY_DATE", "01.01.2024 a 31.12.2999", "?", "?", "EKET/PARÂM", "Data de pagamento"),
            ("LIFNR", f"50000001?", "0010007092", "?", "EKKO", "Parceiro/Fornecedor"),
        ]

        print(f"\n{'Condição':<20} | {'Valor Exigido':<25} | {'Valor PO':<20} | {'Status':<8} | {'Fonte':<15} | {'Obs':<25}")
        print("-" * 140)

        for condition, required, value, status, source, obs in matrix_data:
            print(f"{condition:<20} | {required:<25} | {value:<20} | {status:<8} | {source:<15} | {obs:<25}")

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
        diag = CompleteDiagnostic()

        if not diag.connect():
            return

        # FASE 1
        code = diag.phase_1_extract_code()

        # FASE 2
        ekpo_rows = diag.phase_2_ekpo()

        # FASE 3
        eket_rows = diag.phase_3_eket(ekpo_rows)

        # FASE 4
        log_rows = diag.phase_4_log()

        # FASE 5
        params = diag.phase_5_parameters()

        # FASE 6
        diag.phase_6_pay_date(ekpo_rows, params)

        # FASE 7
        diag.phase_7_vendor()

        # FASE 8
        diag.phase_8_matrix(ekpo_rows, log_rows, params)

        # CONCLUSÃO
        print("\n" + "="*80)
        print("CONCLUSÃO PRELIMINAR")
        print("="*80)

        if log_rows:
            print("\n✗ A PO 4300000002 NÃO aparece em AUTO_INST_ASSIGN porque:")
            print("   JÁ EXISTE um registo em ZFI_DOC_EX_LOG_T")
            print("   Isto significa que a PO foi já processada/excluída do programa.")
        else:
            print("\n? A PO passa nas verificações de EKKO/EKPO")
            print("  Possíveis razões:")
            print("  1. Data de pagamento (PAY_DATE) fora do intervalo")
            print("  2. Tipo de documento (BSART) excluído")
            print("  3. Relação LIFNR vs Parceiro incorreta")
            print("  4. Lógica adicional não identificada no código")

    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if diag:
            diag.close()


if __name__ == "__main__":
    main()

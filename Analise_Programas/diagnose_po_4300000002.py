"""
Diagnóstico: Por que a PO 4300000002 não aparece em AUTO_INST_ASSIGN (ZFI_PURCH_DOC_EX_RATE)

Modo: SOMENTE LEITURA - RFC_READ_TABLE apenas, sem alterações.
Objetivo: Reproduzir exatamente a lógica de filtro do método AUTO_INST_ASSIGN.

Parâmetros de entrada (conforme SAP):
- Empresa: 2010
- Moeda: USD
- Data Inicial Pagamento: 01.01.2024
- Data Final Pagamento: 31.12.2999
- Documento de compras: 4300000002
- Parceiro de negócios: 50000001
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any
from datetime import datetime

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:
    print(f"ERRO: PyRFC não disponível. {exc}")
    sys.exit(1)

# Setup do projeto
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)


class DiagnosticHelper:
    """Lê dados via RFC_READ_TABLE (somente leitura)."""

    def __init__(self, environment: str = "QAD"):
        """Conecta ao ambiente SAP especificado."""
        self.environment = environment
        self.conn = None
        self._connect()

    def _connect(self):
        """Establece conexão RFC."""
        env_prefix = self.environment.upper()

        params = {
            "user": os.getenv(f"SAP_{env_prefix}_USER", ""),
            "passwd": os.getenv(f"SAP_{env_prefix}_PASSWD", ""),
            "ashost": os.getenv(f"SAP_{env_prefix}_ASHOST", ""),
            "sysnr": os.getenv(f"SAP_{env_prefix}_SYSNR", ""),
            "client": os.getenv(f"SAP_{env_prefix}_CLIENT", ""),
        }

        if not all(params.values()):
            raise RuntimeError(f"Credenciais incompletas para {env_prefix}. Verifica .env")

        try:
            self.conn = Connection(**params)
            result = self.conn.call("RFC_PING")
            print(f"✓ Conectado a {self.environment}")
        except Exception as e:
            raise RuntimeError(f"Falha ao conectar a {self.environment}: {e}") from e

    def read_table(self, table: str, fields: list[str], options: list[str] = None, rowcount: int = 0) -> list[dict[str, Any]]:
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
                ROWCOUNT=rowcount,
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

    def close(self):
        """Fecha conexão."""
        if self.conn:
            try:
                self.conn.close()
            except:
                pass


def diagnose_po():
    """Executa o diagnóstico completo."""

    helper = None
    try:
        helper = DiagnosticHelper(environment="QAD")

        # ===== PASSO 1: EKKO (Cabeçalho da PO) =====
        print("\n" + "="*80)
        print("PASSO 1: EKKO - Cabeçalho da Ordem de Compras")
        print("="*80)

        ekko_fields = [
            "EBELN",      # Número do documento de compras
            "BUKRS",      # Empresa
            "WAERS",      # Moeda
            "WKURS",      # Taxa de câmbio
            "KUFIX",      # Indicador de taxa fixa
            "LIFNR",      # Fornecedor/Parceiro de negócios
            "BSART",      # Tipo de documento de compras
            "BEDAT",      # Data de documento
            "EKORG",      # Organização de compras
            "ERDAT",      # Data de criação
            "AEDAT",      # Data de última alteração
        ]

        # Tenta várias formatações do número da PO
        po_formats = ['4300000002', '000004300000002', '4300000002']
        ekko_rows = []
        po_number = None

        for po_format in po_formats:
            ekko_options = [f"EBELN = '{po_format}'"]
            ekko_rows = helper.read_table("EKKO", ekko_fields, ekko_options)
            if ekko_rows:
                print(f"✓ PO encontrada com formato: {po_format}")
                po_number = po_format
                break

        if not ekko_rows:
            # Se não encontrar, lista primeiras POs para debug
            print("\n⚠️  PO não encontrada. Listando primeiras POs em EKKO...")
            try:
                ekko_rows_debug = helper.read_table("EKKO", ["EBELN", "BUKRS", "LIFNR"], [], rowcount=20)
                if ekko_rows_debug:
                    for row in ekko_rows_debug[:10]:
                        print(f"  Exemplo: {row.get('EBELN', 'N/A')} | Emp: {row.get('BUKRS', 'N/A')} | Parceiro: {row.get('LIFNR', 'N/A')}")
                else:
                    print("  ✗ Nenhuma PO encontrada em EKKO (tabela vazia ou sem acesso)")
            except Exception as debug_e:
                print(f"  Erro ao listar POs: {debug_e}")

        if ekko_rows:
            print(f"\n✓ PO encontrada em EKKO:")
            for row in ekko_rows:
                print(f"  Documento: {row.get('EBELN', 'N/A')}")
                print(f"  Empresa: {row.get('BUKRS', 'N/A')}")
                print(f"  Moeda: {row.get('WAERS', 'N/A')}")
                print(f"  Taxa de câmbio: {row.get('WKURS', 'N/A')}")
                print(f"  Indicador taxa fixa: {row.get('KUFIX', 'N/A')}")
                print(f"  Parceiro: {row.get('LIFNR', 'N/A')}")
                print(f"  Tipo de doc: {row.get('BSART', 'N/A')}")
                print(f"  Data de documento: {row.get('BEDAT', 'N/A')}")
        else:
            print("✗ PO NÃO ENCONTRADA em EKKO!")
            return

        # ===== PASSO 2: EKPO (Posições da PO) =====
        print("\n" + "="*80)
        print("PASSO 2: EKPO - Posições da Ordem de Compras")
        print("="*80)

        ekpo_fields = [
            "EBELN",      # Número do documento
            "EBELP",      # Número da posição
            "LFGNI",      # Identificador de mercadoria
            "MENGE",      # Quantidade
            "MEINS",      # Unidade de medida
            "NETPR",      # Preço líquido
            "PEINH",      # Unidade de preço
            "BRTWR",      # Valor bruto
            "WAERS",      # Moeda
            "EDATUM",     # Data de entrega
        ]

        # Usa o número da PO encontrada
        po_number = ekko_rows[0].get('EBELN', '4300000002')
        ekpo_options = [f"EBELN = '{po_number}'"]
        ekpo_rows = helper.read_table("EKPO", ekpo_fields, ekpo_options)

        if ekpo_rows:
            print(f"\n✓ {len(ekpo_rows)} posição(ões) encontrada(s) em EKPO:")
            for row in ekpo_rows:
                print(f"  Posição {row.get('EBELP', 'N/A')}: {row.get('MENGE', 'N/A')} {row.get('MEINS', 'N/A')}")
                print(f"    Preço: {row.get('NETPR', 'N/A')} {row.get('WAERS', 'N/A')}")
                print(f"    Data de entrega: {row.get('EDATUM', 'N/A')}")
        else:
            print("✗ Nenhuma posição encontrada em EKPO!")

        # ===== PASSO 3: EKET (Prazos de Entrega) =====
        print("\n" + "="*80)
        print("PASSO 3: EKET - Prazos de Entrega")
        print("="*80)

        eket_fields = [
            "EBELN",      # Número do documento
            "EBELP",      # Número da posição
            "EKETX",      # Número sequencial
            "EINDT",      # Data de entrega prevista
            "MENGE",      # Quantidade
            "BPUMN",      # Unidade de medida da quantidade
            "BPUMZ",      # Unidade de medida do preço
        ]

        eket_options = [f"EBELN = '{po_number}'"]
        eket_rows = helper.read_table("EKET", eket_fields, eket_options)

        if eket_rows:
            print(f"\n✓ {len(eket_rows)} termo(s) encontrado(s) em EKET:")
            for row in eket_rows:
                print(f"  Posição {row.get('EBELP', 'N/A')}: {row.get('MENGE', 'N/A')} para {row.get('EINDT', 'N/A')}")
        else:
            print("✗ Nenhum termo encontrado em EKET!")

        # ===== PASSO 4: ZFI_DOC_EX_LOG_T =====
        print("\n" + "="*80)
        print("PASSO 4: ZFI_DOC_EX_LOG_T - Log de Documentos de Câmbio")
        print("="*80)

        zfi_fields = [
            "EBELN",      # Documento de compras
            "BUKRS",      # Empresa
            "INSTRUM_ID", # ID do instrumento
            "TIPO_EX",    # Tipo de câmbio
            "DATA_CRIA",  # Data de criação
            "USUARIO",    # Utilizador
        ]

        zfi_options = [f"EBELN = '{po_number}'"]
        try:
            zfi_rows = helper.read_table("ZFI_DOC_EX_LOG_T", zfi_fields, zfi_options)
        except Exception as e:
            print(f"\n⚠️  Tabela ZFI_DOC_EX_LOG_T não acessível: {e}")
            zfi_rows = []

        if zfi_rows:
            print(f"\n✗ {len(zfi_rows)} registo(s) encontrado(s) em ZFI_DOC_EX_LOG_T:")
            print("  ISTO SIGNIFICA: A PO já foi processada/excluída do AUTO_INST_ASSIGN!")
            for row in zfi_rows:
                print(f"    Instrumento: {row.get('INSTRUM_ID', 'N/A')}, Tipo: {row.get('TIPO_EX', 'N/A')}")
        else:
            print("\n✓ Nenhum registo em ZFI_DOC_EX_LOG_T para esta PO")
            print("  INTERPRETAÇÃO: A PO ainda não foi processada, logo deve aparecer em AUTO_INST_ASSIGN")

        # ===== PASSO 5: Parâmetros =====
        print("\n" + "="*80)
        print("PASSO 5: Parâmetros do Programa (ZCLCA_FIXEDVALS)")
        print("="*80)

        params_fields = [
            "MANDT",      # Cliente
            "MODULE",     # Módulo
            "PROCESSO",   # Processo
            "PARAM_NAME", # Nome do parâmetro
            "PARAM_VALUE",# Valor do parâmetro
        ]

        params_options = [
            "MODULE = 'FIN'",
            "PROCESSO = 'PO_EXCHANGE_RATE'",
        ]

        try:
            params_rows = helper.read_table("ZCLCA_FIXEDVALS", params_fields, params_options)

            if params_rows:
                print(f"\n✓ Parâmetros encontrados:")
                params_dict = {}
                for row in params_rows:
                    param_name = row.get('PARAM_NAME', '').strip()
                    param_value = row.get('PARAM_VALUE', '').strip()
                    params_dict[param_name] = param_value
                    print(f"  {param_name} = {param_value}")

                # Extrai NDAYS se existir
                ndays = int(params_dict.get('NDAYS', '0')) if 'NDAYS' in params_dict else None
                if ndays is not None:
                    print(f"\n  ⚠️  NDAYS = {ndays} (número de dias usados no cálculo)")
            else:
                print("\n⚠️  Nenhum parâmetro encontrado em ZCLCA_FIXEDVALS")
                print("  (Pode estar a usar valores hardcoded no programa)")

        except Exception as e:
            print(f"\n⚠️  Erro ao consultar ZCLCA_FIXEDVALS: {e}")
            print("  (Tabela pode não existir ou estar configurada diferentemente)")

        # ===== PASSO 6: Reproduzir Lógica AUTO_INST_ASSIGN =====
        print("\n" + "="*80)
        print("PASSO 6: Análise de Filtros e Cálculos")
        print("="*80)

        print("\nParâmetros de entrada do utilizador:")
        print("  Empresa: 2010")
        print("  Moeda: USD")
        print("  Data Inicial: 01.01.2024")
        print("  Data Final: 31.12.2999")
        print("  Documento: 4300000002")
        print("  Parceiro: 50000001")

        # Validar condições de filtro
        print("\nValidação de Condições de Filtro (AUTO_INST_ASSIGN):")

        conditions_results = {}

        if ekko_rows:
            ekko = ekko_rows[0]

            # Condição 1: Empresa
            bukrs_match = ekko.get('BUKRS', '').strip() == '2010'
            conditions_results['Empresa = 2010'] = (ekko.get('BUKRS', 'N/A'), '2010', 'PASSA' if bukrs_match else 'FALHA')

            # Condição 2: Moeda
            waers_match = ekko.get('WAERS', '').strip() == 'USD'
            conditions_results['Moeda = USD'] = (ekko.get('WAERS', 'N/A'), 'USD', 'PASSA' if waers_match else 'FALHA')

            # Condição 3: Parceiro
            lifnr_match = ekko.get('LIFNR', '').strip() == '50000001'
            conditions_results['Parceiro = 50000001'] = (ekko.get('LIFNR', 'N/A'), '50000001', 'PASSA' if lifnr_match else 'FALHA')

            # Condição 4: Tipo de documento (assumir que certos tipos são excluídos)
            bsart = ekko.get('BSART', '').strip()
            conditions_results['Tipo de doc válido'] = (bsart, 'Não KB/BL', 'PASSA' if bsart not in ['KB', 'BL'] else 'FALHA')

            # Condição 5: ZFI_DOC_EX_LOG_T
            log_exists = len(zfi_rows) > 0
            conditions_results['Sem log em ZFI_DOC_EX_LOG_T'] = ('Sim' if log_exists else 'Não', 'Não', 'FALHA' if log_exists else 'PASSA')

        # Tabela de diagnóstico
        print("\n" + "-"*100)
        print(f"{'CONDIÇÃO':<40} | {'VALOR DA PO':<20} | {'EXIGIDO':<20} | {'STATUS':<10}")
        print("-"*100)

        failing_conditions = []
        for cond, (value, required, status) in conditions_results.items():
            print(f"{cond:<40} | {value:<20} | {required:<20} | {status:<10}")
            if status == 'FALHA':
                failing_conditions.append(cond)

        # Conclusão
        print("\n" + "="*80)
        print("CONCLUSÃO")
        print("="*80)

        if failing_conditions:
            print(f"\n✗ A PO 4300000002 FALHA EM {len(failing_conditions)} CONDIÇÃO(ÕES):\n")
            for i, cond in enumerate(failing_conditions, 1):
                print(f"  {i}. {cond}")
        else:
            print("\n✓ A PO 4300000002 DEVERIA APARECER em AUTO_INST_ASSIGN")
            print("\nPossíveis razões para não aparecer:")
            print("  1. A PO foi excluída manualmente")
            print("  2. Existe um filtro adicional no código do programa não mapeado aqui")
            print("  3. A lógica de seleção é diferente do esperado")
            print("  4. Há um erro no programa ou nos dados")

    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if helper:
            helper.close()


if __name__ == "__main__":
    diagnose_po()

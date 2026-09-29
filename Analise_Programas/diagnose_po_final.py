"""
Diagnóstico Final Consolidado: Por que PO 4300000002 não aparece em AUTO_INST_ASSIGN

Dados confirmados:
- EKKO existe: ✓
- EKPO tem 1 posição: ✓
- Agora: investigar por que não aparece em AUTO_INST_ASSIGN
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any
from datetime import datetime

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


class FinalDiagnostic:
    """Diagnóstico final consolidado."""

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

    def run_all_phases(self):
        """Executa todas as fases do diagnóstico."""

        print("\n" + "="*100)
        print("DIAGNÓSTICO FINAL: PO 4300000002 em ZFI_PURCH_DOC_EX_RATE")
        print("="*100)

        # FASE 1: EKKO Completo
        print("\n[FASE 1] EKKO - Cabeçalho da PO")
        print("-" * 100)

        ekko_fields = [
            "EBELN", "BUKRS", "WAERS", "WKURS", "KUFIX", "LIFNR", "BSART", "BSTYP",
            "LOEKZ", "STATU", "BEDAT", "AEDAT", "EKORG",
        ]

        ekko_rows = self.read_table("EKKO", ekko_fields, [f"EBELN = '{self.po_number}'"])

        if ekko_rows:
            ekko = ekko_rows[0]
            print(f"✓ EKKO encontrada")
            print(f"  EBELN: {ekko.get('EBELN')}")
            print(f"  BUKRS (Empresa): {ekko.get('BUKRS')}")
            print(f"  WAERS (Moeda): {ekko.get('WAERS')}")
            print(f"  WKURS (Taxa): {ekko.get('WKURS')}")
            print(f"  KUFIX (Taxa fixa): '{ekko.get('KUFIX', 'VAZIO')}'")
            print(f"  LIFNR (Fornecedor): {ekko.get('LIFNR')}")
            print(f"  BSART (Tipo doc): {ekko.get('BSART')}")
            print(f"  BSTYP (Tipo pedido): {ekko.get('BSTYP')}")
            print(f"  LOEKZ (Exclusão): '{ekko.get('LOEKZ', 'VAZIO')}'")
            print(f"  STATU (Status): {ekko.get('STATU')}")
            print(f"  BEDAT (Data doc): {ekko.get('BEDAT')}")
            print(f"  AEDAT (Data alt): {ekko.get('AEDAT')}")

        # FASE 2: EKPO Completo
        print("\n[FASE 2] EKPO - Posições da PO")
        print("-" * 100)

        ekpo_fields = [
            "EBELN", "EBELP", "PSTYP", "LOEKZ", "MENGE", "MEINS", "NETPR", "PEINH",
            "WAERS", "EDATUM", "EVERS", "KNUMV",
        ]

        ekpo_rows = self.read_table("EKPO", ekpo_fields, [f"EBELN = '{self.po_number}'"])

        print(f"✓ {len(ekpo_rows)} posição(ões) encontrada(s)")
        for ekpo in ekpo_rows:
            print(f"\n  Posição {ekpo.get('EBELP')}:")
            print(f"    PSTYP (Categoria): {ekpo.get('PSTYP')}")
            print(f"    LOEKZ (Exclusão): '{ekpo.get('LOEKZ', 'VAZIO')}'")
            print(f"    MENGE (Quantidade): {ekpo.get('MENGE')}")
            print(f"    MEINS (Unidade): {ekpo.get('MEINS')}")
            print(f"    NETPR (Preço): {ekpo.get('NETPR')}")
            print(f"    WAERS (Moeda): {ekpo.get('WAERS')}")
            print(f"    EDATUM (Data entrega): {ekpo.get('EDATUM')}")
            print(f"    EVERS (Status entrega): {ekpo.get('EVERS')}")

        # FASE 3: EKET - Prazos
        print("\n[FASE 3] EKET - Prazos de Entrega")
        print("-" * 100)

        eket_fields = [
            "EBELN", "EBELP", "EKETX", "EINDT", "MENGE", "BPUMN",
        ]

        eket_rows = self.read_table("EKET", eket_fields, [f"EBELN = '{self.po_number}'"])

        if eket_rows:
            print(f"✓ {len(eket_rows)} termo(s) encontrado(s)")
            for eket in eket_rows:
                print(f"  Pos {eket.get('EBELP')}, seq {eket.get('EKETX')}: {eket.get('MENGE')} para {eket.get('EINDT')}")
        else:
            print("✗ Nenhum termo em EKET")

        # FASE 4: ZFI_DOC_EX_LOG_T - Log
        print("\n[FASE 4] ZFI_DOC_EX_LOG_T - Log de Processamento")
        print("-" * 100)

        log_fields = ["EBELN", "BUKRS", "INSTRUM_ID", "TIPO_EX", "DATA_CRIA", "USUARIO"]

        try:
            log_rows = self.read_table("ZFI_DOC_EX_LOG_T", log_fields, [f"EBELN = '{self.po_number}'"])

            if log_rows:
                print(f"✗ {len(log_rows)} registo(s) em log - PO JÁ FOI PROCESSADA!")
                for log in log_rows:
                    print(f"  {log}")
            else:
                print("✓ Nenhum log - PO ainda não foi processada")
        except Exception as e:
            print(f"⚠️  Erro ao consultar log: {e}")

        # FASE 5: Investigar PSTYP
        print("\n[FASE 5] Análise de PSTYP (Categoria de Posição)")
        print("-" * 100)

        if ekpo_rows:
            pstyp = ekpo_rows[0].get('PSTYP', '')
            print(f"PSTYP encontrado: {pstyp}")

            # Consultar T161 para descrição
            try:
                t161_rows = self.read_table("T161", ["PSTYP", "PSTXT"], [f"PSTYP = '{pstyp}'"])
                if t161_rows:
                    print(f"  Descrição: {t161_rows[0].get('PSTXT')}")
            except:
                pass

            # Análise de categorias conhecidas
            pstyp_analysis = {
                "1": "Item padrão",
                "5": "Serviço",
                "9": "Item subcontratado",
                "": "Não informado",
            }

            analysis = pstyp_analysis.get(pstyp, f"Tipo desconhecido: {pstyp}")
            print(f"  Tipo: {analysis}")

            # AUTO_INST_ASSIGN pode excluir serviços (PSTYP=5)
            if pstyp == "5":
                print(f"\n  ⚠️  ACHADO POTENCIAL: PSTYP=5 (Serviço)")
                print(f"      AUTO_INST_ASSIGN pode excluir posições de serviço!")

        # FASE 6: Dados de Câmbio
        print("\n[FASE 6] Dados de Câmbio")
        print("-" * 100)

        if ekko_rows:
            ekko = ekko_rows[0]
            print(f"WKURS (Taxa SAP): {ekko.get('WKURS')}")
            print(f"KUFIX (Taxa fixa): '{ekko.get('KUFIX', 'VAZIO')}'")

            # Taxa é relevante para AUTO_INST_ASSIGN?
            if not ekko.get('WKURS') or ekko.get('WKURS').strip() == '0':
                print("  ⚠️  Taxa de câmbio é zero ou vazio - pode ser excluída")
            else:
                print("  ✓ Taxa de câmbio presente")

        # FASE 7: Relação LIFNR vs Parceiro
        print("\n[FASE 7] Relação LIFNR vs Parceiro")
        print("-" * 100)

        if ekko_rows:
            lifnr = ekko_rows[0].get('LIFNR', '')
            print(f"EKKO-LIFNR: {lifnr}")
            print(f"Parceiro mencionado na execução: 50000001")

            # Investigar se há relação
            try:
                lfa1_rows = self.read_table("LFA1", ["LIFNR", "NAME1"], [f"LIFNR = '{lifnr}'"])
                if lfa1_rows:
                    print(f"✓ Fornecedor {lifnr}: {lfa1_rows[0].get('NAME1')}")
            except:
                pass

            # Se são diferentes, pode haver filtro por parceiro
            if lifnr != "50000001":
                print(f"\n  ⚠️  ACHADO POTENCIAL: LIFNR ({lifnr}) ≠ Parceiro ({50000001})")
                print(f"      AUTO_INST_ASSIGN pode filtrar por parceiro específico!")

        # FASE 8: Status e Exclusão
        print("\n[FASE 8] Status e Exclusão")
        print("-" * 100)

        if ekko_rows:
            ekko = ekko_rows[0]
            loekz = ekko.get('LOEKZ', '')
            statu = ekko.get('STATU', '')

            print(f"LOEKZ (Marcada exclusão): '{loekz}' (vazio = NÃO)")
            print(f"STATU (Status): {statu}")

            if loekz and loekz.strip():
                print("  ✗ ACHADO: PO marcada para exclusão")
            else:
                print("  ✓ PO não marcada para exclusão")

            if statu == "9":
                print("  STATU=9 normalmente significa: Em preparação/Não liberada")
            elif statu == "4":
                print("  STATU=4 normalmente significa: Completa")

        # MATRIZ FINAL
        print("\n" + "="*100)
        print("MATRIZ FINAL DE FILTROS")
        print("="*100)

        matrix_data = []

        if ekko_rows:
            ekko = ekko_rows[0]

            # Montar matriz
            matrix_data = [
                ("BUKRS", "2010", ekko.get('BUKRS', 'N/A'), "✓" if ekko.get('BUKRS') == "2010" else "✗"),
                ("WAERS", "USD", ekko.get('WAERS', 'N/A'), "✓" if ekko.get('WAERS') == "USD" else "✗"),
                ("LOEKZ", "Vazio", ekko.get('LOEKZ', 'VAZIO'), "✓" if not ekko.get('LOEKZ', '').strip() else "✗"),
                ("EBELN", "4300000002", ekko.get('EBELN', 'N/A'), "✓"),
                ("LIFNR", "?", ekko.get('LIFNR', 'N/A'), "?" if ekko.get('LIFNR') != "50000001" else "✓"),
                ("PSTYP", "Não 5?", ekpo_rows[0].get('PSTYP', 'N/A') if ekpo_rows else "N/A", "✗" if (ekpo_rows and ekpo_rows[0].get('PSTYP') == "5") else "?"),
                ("WKURS", ">0", ekko.get('WKURS', 'N/A'), "✓" if ekko.get('WKURS', '0') != '0' else "✗"),
                ("LOG", "Vazio", f"{len(log_rows) if 'log_rows' in locals() else 0} registos", "✓" if (len(log_rows) if 'log_rows' in locals() else 0) == 0 else "✗"),
            ]

        print(f"\n{'Filtro':<15} | {'Esperado':<20} | {'Valor':<20} | {'Status':<10}")
        print("-" * 70)
        for filtro, esperado, valor, status in matrix_data:
            print(f"{filtro:<15} | {esperado:<20} | {valor:<20} | {status:<10}")

        # CONCLUSÃO
        print("\n" + "="*100)
        print("CONCLUSÃO")
        print("="*100)

        falhas = [f for f, _, _, s in matrix_data if s == "✗"]
        incertos = [f for f, _, _, s in matrix_data if s == "?"]

        if falhas:
            print(f"\n✗ A PO NÃO PASSA em {len(falhas)} condição(ões):")
            for f in falhas:
                print(f"  • {f}")

        if incertos:
            print(f"\n? INCERTO em {len(incertos)} condição(ões):")
            for f in incertos:
                print(f"  • {f}")

        if not falhas and not incertos:
            print("\n✓ A PO passa em todas as verificações básicas")
            print("  Pode haver lógica adicional não identificada no código")

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
        diag = FinalDiagnostic()
        diag.connect()
        print("✓ Conectado a S4Q/100")
        diag.run_all_phases()
    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if diag:
            diag.close()


if __name__ == "__main__":
    main()

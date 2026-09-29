"""
Extração do Código-Fonte ABAP do ZFI_PURCH_DOC_EX_RATE

Objetivo: Localizar o método AUTO_INST_ASSIGN e extrair
a lógica real de seleção para validar a conclusão sobre PSTYP=5.

SOMENTE LEITURA - nenhuma alteração.
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


class ABAPSourceExtractor:
    """Extrai código-fonte ABAP do SAP."""

    def __init__(self):
        self.conn = None

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

    def extract_program_source(self, program_name: str) -> list[str]:
        """Tenta extrair código-fonte do programa."""

        print(f"\n[TENTATIVA 1] RFC_READ_REPORT para {program_name}")
        print("-" * 80)

        try:
            result = self.conn.call(
                "RFC_READ_REPORT",
                PROGRAM=program_name,
            )

            lines = result.get("PROGRAM", [])
            if lines:
                print(f"✓ Programa extraído com {len(lines)} linhas")
                code_lines = [str(line.get("LINE", "")) for line in lines]
                return code_lines
            else:
                print("✗ RFC_READ_REPORT retornou vazio")
        except Exception as e:
            print(f"✗ RFC_READ_REPORT falhou: {e}")

        # Tentativa 2: READ_REPOSITORY
        print(f"\n[TENTATIVA 2] READ_REPOSITORY para {program_name}")
        print("-" * 80)

        try:
            result = self.conn.call(
                "READ_REPOSITORY",
                PROGRAM=program_name,
            )

            lines = result.get("SOURCE", []) or result.get("PROGRAM", [])
            if lines:
                print(f"✓ Programa extraído com {len(lines)} linhas")
                code_lines = [str(line.get("LINE", "")) for line in lines]
                return code_lines
            else:
                print("✗ READ_REPOSITORY retornou vazio")
        except Exception as e:
            print(f"✗ READ_REPOSITORY falhou: {e}")

        # Tentativa 3: Ler de ZREPOSITORY (tabela padrão SAP de código)
        print(f"\n[TENTATIVA 3] Tabela D010S (Source Code Repository)")
        print("-" * 80)

        try:
            # D010S contém código-fonte de programas
            source_rows = self.read_table(
                "D010S",
                ["PROG", "CDAT", "CNAM", "CCNT"],
                [f"PROG = '{program_name}'"]
            )

            if source_rows:
                print(f"✓ Encontrados {len(source_rows)} blocos em D010S")
                return [row.get("CCNT", "") for row in source_rows]
            else:
                print("✗ Nenhum bloco encontrado em D010S")
        except Exception as e:
            print(f"✗ Erro ao consultar D010S: {e}")

        # Tentativa 4: Ler de REPSRC (Source Repository)
        print(f"\n[TENTATIVA 4] Tabela REPSRC")
        print("-" * 80)

        try:
            source_rows = self.read_table(
                "REPSRC",
                ["PROG", "BUKRS", "ATEXT"],
                [f"PROG = '{program_name}'"]
            )

            if source_rows:
                print(f"✓ Encontrados {len(source_rows)} registos em REPSRC")
                for row in source_rows:
                    print(f"  {row}")
                return []
            else:
                print("✗ Nenhum registo em REPSRC")
        except Exception as e:
            print(f"✗ Erro ao consultar REPSRC: {e}")

        return []

    def search_in_code(self, code_lines: list[str]) -> dict:
        """Procura por padrões relevantes no código."""

        print("\n" + "="*80)
        print("ANÁLISE DO CÓDIGO")
        print("="*80)

        findings = {
            "SELECT_EKKO": [],
            "SELECT_EKPO": [],
            "SELECT_EKET": [],
            "PSTYP_CONDITIONS": [],
            "STATU_CONDITIONS": [],
            "LIFNR_CONDITIONS": [],
            "DELETE_STATEMENTS": [],
            "FILTER_STATEMENTS": [],
        }

        if not code_lines:
            print("✗ Nenhum código disponível para análise")
            return findings

        code_text = "\n".join(code_lines)

        # Procura por SELECT
        for i, line in enumerate(code_lines):
            upper_line = line.upper()

            if "SELECT" in upper_line and "EKKO" in upper_line:
                findings["SELECT_EKKO"].append((i+1, line))

            if "SELECT" in upper_line and "EKPO" in upper_line:
                findings["SELECT_EKPO"].append((i+1, line))

            if "SELECT" in upper_line and "EKET" in upper_line:
                findings["SELECT_EKET"].append((i+1, line))

            # Procura por PSTYP
            if "PSTYP" in upper_line:
                findings["PSTYP_CONDITIONS"].append((i+1, line))

            # Procura por STATU
            if "STATU" in upper_line and "=" in upper_line:
                findings["STATU_CONDITIONS"].append((i+1, line))

            # Procura por LIFNR
            if "LIFNR" in upper_line:
                findings["LIFNR_CONDITIONS"].append((i+1, line))

            # Procura por DELETE
            if "DELETE" in upper_line:
                findings["DELETE_STATEMENTS"].append((i+1, line))

            # Procura por FILTER ou WHERE
            if ("FILTER" in upper_line or "WHERE" in upper_line) and ("EKPO" in upper_line or "EKKO" in upper_line):
                findings["FILTER_STATEMENTS"].append((i+1, line))

        return findings

    def print_findings(self, findings: dict, code_lines: list[str]):
        """Imprime os achados."""

        if findings["SELECT_EKKO"]:
            print("\n[ACHADO] SELECT com EKKO:")
            for line_no, line in findings["SELECT_EKKO"]:
                print(f"  Linha {line_no}: {line}")

        if findings["SELECT_EKPO"]:
            print("\n[ACHADO] SELECT com EKPO:")
            for line_no, line in findings["SELECT_EKPO"]:
                print(f"  Linha {line_no}: {line}")

        if findings["PSTYP_CONDITIONS"]:
            print("\n[ACHADO] Condições com PSTYP:")
            for line_no, line in findings["PSTYP_CONDITIONS"]:
                print(f"  Linha {line_no}: {line}")

                # Mostra contexto
                if code_lines and line_no > 0:
                    start = max(0, line_no - 3)
                    end = min(len(code_lines), line_no + 3)
                    print("    --- Contexto ---")
                    for i in range(start, end):
                        marker = ">>>" if i == line_no - 1 else "   "
                        print(f"    {marker} {i+1}: {code_lines[i]}")

        if findings["STATU_CONDITIONS"]:
            print("\n[ACHADO] Condições com STATU:")
            for line_no, line in findings["STATU_CONDITIONS"]:
                print(f"  Linha {line_no}: {line}")

        if findings["LIFNR_CONDITIONS"]:
            print("\n[ACHADO] Condições com LIFNR:")
            for line_no, line in findings["LIFNR_CONDITIONS"]:
                print(f"  Linha {line_no}: {line}")

        if findings["DELETE_STATEMENTS"]:
            print("\n[ACHADO] Declarações DELETE:")
            for line_no, line in findings["DELETE_STATEMENTS"]:
                print(f"  Linha {line_no}: {line}")

        if findings["FILTER_STATEMENTS"]:
            print("\n[ACHADO] Declarações FILTER:")
            for line_no, line in findings["FILTER_STATEMENTS"]:
                print(f"  Linha {line_no}: {line}")

    def close(self):
        """Fecha conexão."""
        if self.conn:
            try:
                self.conn.close()
            except:
                pass


def main():
    extractor = None
    try:
        extractor = ABAPSourceExtractor()
        extractor.connect()

        # Extrai código-fonte
        print("\n" + "="*80)
        print("EXTRAÇÃO DO CÓDIGO-FONTE: ZFI_PURCH_DOC_EX_RATE")
        print("="*80)

        code_lines = extractor.extract_program_source("ZFI_PURCH_DOC_EX_RATE")

        if code_lines:
            print(f"\n✓ Código extraído com sucesso ({len(code_lines)} linhas)")

            # Procura por AUTO_INST_ASSIGN
            auto_inst_found = False
            for i, line in enumerate(code_lines):
                if "AUTO_INST_ASSIGN" in line.upper():
                    print(f"\n✓ AUTO_INST_ASSIGN encontrado na linha {i+1}")
                    auto_inst_found = True

                    # Mostra contexto
                    start = max(0, i - 5)
                    end = min(len(code_lines), i + 50)
                    print("\n--- Trecho contendo AUTO_INST_ASSIGN ---")
                    for j in range(start, end):
                        marker = ">>>" if j == i else "   "
                        print(f"{marker} {j+1}: {code_lines[j]}")
                    break

            if not auto_inst_found:
                print("\n⚠️  AUTO_INST_ASSIGN não encontrado no código extraído")

            # Análise geral
            findings = extractor.search_in_code(code_lines)
            extractor.print_findings(findings, code_lines)
        else:
            print("\n✗ Não foi possível extrair código-fonte")
            print("\nAlternativas:")
            print("  1. Código pode estar em um include separado")
            print("  2. Pode ser uma classe ABAP (não programa direto)")
            print("  3. Acesso ao repository pode estar restrito")

    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if extractor:
            extractor.close()


if __name__ == "__main__":
    main()

"""
Investigação Automática - Encontrar a Causa Raiz

Objetivo: Testar cada INNER JOIN e WHERE do SELECT AUTO_INST_ASSIGN
até encontrar a primeira falha que elimina a PO 4300000002.

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


class RootCauseFinder:
    """Encontra a causa raiz testando cada filtro do AUTO_INST_ASSIGN."""

    def __init__(self):
        self.conn = None
        self.po = "4300000002"
        self.results = []

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

    def test_ekko(self):
        """Testa EKKO (base)."""
        print("\n[ETAPA 1.1] EKKO (Cabeçalho)")
        print("-" * 100)

        result = self.read_table(
            "EKKO",
            ["EBELN", "BUKRS", "WAERS", "LIFNR", "ZZCOORI", "ZZEXPVZ"],
            [f"EBELN = '{self.po}'"]
        )

        status = "PASSA" if result else "FALHA"
        print(f"RESULTADO: {len(result)} linha(s) - {status}")

        if result:
            row = result[0]
            print(f"  EBELN: {row.get('EBELN')}")
            print(f"  BUKRS: {row.get('BUKRS')}")
            print(f"  WAERS: {row.get('WAERS')}")
            print(f"  LIFNR: {row.get('LIFNR')}")
            print(f"  ZZCOORI: {row.get('ZZCOORI')}")
            print(f"  ZZEXPVZ: {row.get('ZZEXPVZ')}")

        self.results.append(("EKKO", "Base", "Existe", len(result), status))
        return result

    def test_eket(self):
        """Testa EKET JOIN."""
        print("\n[ETAPA 1.2] EKET (INNER JOIN EKET)")
        print("-" * 100)

        result = self.read_table(
            "EKET",
            ["EBELN", "EBELP"],
            [f"EBELN = '{self.po}'"]
        )

        status = "PASSA" if result else "FALHA"
        print(f"Condição ON: B~EBELN EQ A~EBELN")
        print(f"RESULTADO: {len(result)} linha(s) - {status}")

        if result:
            for row in result[:3]:
                print(f"  EBELN: {row.get('EBELN')}, EBELP: {row.get('EBELP')}")

        self.results.append(("EKET", "INNER JOIN", "B~EBELN = A~EBELN", len(result), status))
        return result

    def test_lfb1(self):
        """Testa LFB1 JOIN."""
        print("\n[ETAPA 1.3] LFB1 (INNER JOIN LFB1)")
        print("-" * 100)

        # Primeiro obter LIFNR e BUKRS de EKKO
        ekko = self.read_table("EKKO", ["LIFNR", "BUKRS"], [f"EBELN = '{self.po}'"])
        if not ekko:
            print("FALHA: EKKO não encontrada")
            self.results.append(("LFB1", "INNER JOIN", "D~LIFNR=A~LIFNR AND D~BUKRS=A~BUKRS", 0, "FALHA"))
            return []

        lifnr = ekko[0].get('LIFNR')
        bukrs = ekko[0].get('BUKRS')

        print(f"Condição ON: D~LIFNR EQ A~LIFNR AND D~BUKRS EQ A~BUKRS")
        print(f"Valores: LIFNR={lifnr}, BUKRS={bukrs}")

        result = self.read_table(
            "LFB1",
            ["LIFNR", "BUKRS", "ZTERM"],
            [f"LIFNR = '{lifnr}'", f"BUKRS = '{bukrs}'"]
        )

        status = "PASSA" if result else "FALHA"
        print(f"RESULTADO: {len(result)} linha(s) - {status}")

        if result:
            print(f"  ZTERM: {result[0].get('ZTERM')}")

        self.results.append(("LFB1", "INNER JOIN", f"LIFNR={lifnr} AND BUKRS={bukrs}", len(result), status))
        return result

    def test_zfi_pay_date_t(self):
        """Testa ZFI_PAY_DATE_T JOIN."""
        print("\n[ETAPA 1.4] ZFI_PAY_DATE_T (INNER JOIN ZFI_PAY_DATE_T)")
        print("-" * 100)

        # Primeiro obter ZZCOORI e ZZEXPVZ de EKKO
        ekko = self.read_table("EKKO", ["ZZCOORI", "ZZEXPVZ"], [f"EBELN = '{self.po}'"])
        if not ekko:
            print("FALHA: EKKO não encontrada")
            self.results.append(("ZFI_PAY_DATE_T", "INNER JOIN", "E~ZZCOORI=A~ZZCOORI AND E~ZZEXPVZ=A~ZZEXPVZ", 0, "FALHA"))
            return []

        zzcoori = ekko[0].get('ZZCOORI')
        zzexpvz = ekko[0].get('ZZEXPVZ')

        print(f"Condição ON: E~ZZCOORI EQ A~ZZCOORI AND E~ZZEXPVZ EQ A~ZZEXPVZ")
        print(f"Valores: ZZCOORI={zzcoori}, ZZEXPVZ={zzexpvz}")

        try:
            result = self.read_table(
                "ZFI_PAY_DATE_T",
                ["ZZCOORI", "ZZEXPVZ", "NDAYS"],
                [f"ZZCOORI = '{zzcoori}'", f"ZZEXPVZ = '{zzexpvz}'"]
            )

            status = "PASSA" if result else "FALHA"
            print(f"RESULTADO: {len(result)} linha(s) - {status}")

            if result:
                print(f"  NDAYS: {result[0].get('NDAYS')}")

            self.results.append(("ZFI_PAY_DATE_T", "INNER JOIN", f"ZZCOORI={zzcoori} AND ZZEXPVZ={zzexpvz}", len(result), status))
            return result

        except Exception as e:
            print(f"⚠️  Tabela ZFI_PAY_DATE_T não acessível: {e}")
            self.results.append(("ZFI_PAY_DATE_T", "INNER JOIN", f"ZZCOORI={zzcoori} AND ZZEXPVZ={zzexpvz}", 0, "ERRO"))
            return []

    def test_t052(self):
        """Testa T052 JOIN."""
        print("\n[ETAPA 1.5] T052 (INNER JOIN T052)")
        print("-" * 100)

        # Primeiro obter ZTERM de LFB1
        ekko = self.read_table("EKKO", ["LIFNR", "BUKRS"], [f"EBELN = '{self.po}'"])
        if not ekko:
            self.results.append(("T052", "INNER JOIN", "F~ZTERM = D~ZTERM", 0, "FALHA"))
            return []

        lfb1 = self.read_table("LFB1", ["ZTERM"], [f"LIFNR = '{ekko[0].get('LIFNR')}'", f"BUKRS = '{ekko[0].get('BUKRS')}'"])

        if not lfb1:
            print("FALHA: LFB1 não encontrada")
            self.results.append(("T052", "INNER JOIN", "F~ZTERM = D~ZTERM", 0, "FALHA"))
            return []

        zterm = lfb1[0].get('ZTERM')

        print(f"Condição ON: F~ZTERM EQ D~ZTERM")
        print(f"Valor: ZTERM={zterm}")

        result = self.read_table(
            "T052",
            ["ZTERM"],
            [f"ZTERM = '{zterm}'"]
        )

        status = "PASSA" if result else "FALHA"
        print(f"RESULTADO: {len(result)} linha(s) - {status}")

        self.results.append(("T052", "INNER JOIN", f"ZTERM={zterm}", len(result), status))
        return result

    def test_ekpo(self):
        """Testa EKPO JOIN."""
        print("\n[ETAPA 1.6] EKPO (INNER JOIN EKPO)")
        print("-" * 100)

        # Obter EBELP de EKET
        eket = self.read_table("EKET", ["EBELP"], [f"EBELN = '{self.po}'"])

        if not eket:
            print("FALHA: EKET não encontrada")
            self.results.append(("EKPO", "INNER JOIN", "G~EBELN = A~EBELN AND G~EBELP = B~EBELP", 0, "FALHA"))
            return []

        ebelp = eket[0].get('EBELP')

        print(f"Condição ON: G~EBELN = A~EBELN AND G~EBELP = B~EBELP")
        print(f"Valores: EBELN={self.po}, EBELP={ebelp}")

        result = self.read_table(
            "EKPO",
            ["EBELN", "EBELP", "NETWR"],
            [f"EBELN = '{self.po}'", f"EBELP = '{ebelp}'"]
        )

        status = "PASSA" if result else "FALHA"
        print(f"RESULTADO: {len(result)} linha(s) - {status}")

        if result:
            print(f"  NETWR: {result[0].get('NETWR')}")

        self.results.append(("EKPO", "INNER JOIN", f"EBELN={self.po} AND EBELP={ebelp}", len(result), status))
        return result

    def run(self):
        """Executa toda a investigação."""
        print("\n" + "="*100)
        print("INVESTIGAÇÃO AUTOMÁTICA - CAUSA RAIZ DE EXCLUSÃO DA PO 4300000002")
        print("="*100)

        print("\n[FASE 1] TESTANDO INNER JOINs")
        print("="*100)

        # Testa cada JOIN
        ekko = self.test_ekko()
        eket = self.test_eket()
        lfb1 = self.test_lfb1()
        zfi = self.test_zfi_pay_date_t()
        t052 = self.test_t052()
        ekpo = self.test_ekpo()

        # Resumo
        print("\n" + "="*100)
        print("RESUMO DOS INNER JOINs")
        print("="*100)

        print(f"\n{'Tabela':<20} | {'Condição':<45} | {'Resultado':<10}")
        print("-" * 100)

        falhas = []
        for etapa, _, condicao, count, status in self.results:
            print(f"{etapa:<20} | {condicao:<45} | {count:>3} linhas - {status}")
            if status == "FALHA":
                falhas.append((etapa, condicao, count, status))

        if falhas:
            print("\n" + "="*100)
            print("🎯 CAUSA RAIZ ENCONTRADA!")
            print("="*100)

            primeira_falha = falhas[0]
            print(f"\nPrimeira falha: {primeira_falha[0]}")
            print(f"Condição: {primeira_falha[1]}")
            print(f"Resultado: {primeira_falha[2]} linhas")

            self.show_root_cause()

        else:
            print("\n⚠️  Todos os INNER JOINs passaram")
            print("    Investigar condições WHERE...")

    def show_root_cause(self):
        """Exibe a causa raiz encontrada."""
        print("\n" + "="*100)
        print("CAUSA RAIZ FINAL")
        print("="*100)

        for etapa, tipo, condicao, count, status in self.results:
            if status == "FALHA":
                print(f"\nETAPA: {etapa}")
                print(f"TIPO: {tipo}")
                print(f"CONDIÇÃO ABAP: {condicao}")
                print(f"RESULTADO EM QAD: {count} linha(s)")
                print(f"STATUS: FALHA")

                print(f"\nEXPLICAÇÃO:")
                print(f"  O INNER JOIN com {etapa} falha porque não há correspondência.")
                print(f"  Uma PO que não possui registos em {etapa} com a condição")
                print(f"  '{condicao}' é automaticamente excluída do resultado do SELECT.")

                break

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
        finder = RootCauseFinder()
        finder.connect()
        print("✓ Conectado a S4Q/100")
        finder.run()
    except Exception as e:
        print(f"\n✗ ERRO: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if finder:
            finder.close()


if __name__ == "__main__":
    main()

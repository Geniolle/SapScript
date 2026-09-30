# -*- coding: utf-8 -*-
"""
Investigacao Completa: Por que POs nao aparecem em AUTO_INST_ASSIGN
Protocolo 15 fases (100% SOMENTE LEITURA)

POs a investigar:
- 4300000003
- 4000055997
"""

import os
import sys
from pathlib import Path
from pyrfc import Connection

project_root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)

def connect():
    env_prefix = "QAD"
    params = {
        "user": os.getenv(f"SAP_{env_prefix}_USER", "").strip(),
        "passwd": os.getenv(f"SAP_{env_prefix}_PASSWD", "").strip(),
        "ashost": os.getenv(f"SAP_{env_prefix}_ASHOST", "").strip(),
        "sysnr": os.getenv(f"SAP_{env_prefix}_SYSNR", "").strip(),
        "client": os.getenv(f"SAP_{env_prefix}_CLIENT", "").strip(),
    }
    return Connection(**params)

def read_table(conn, table, fields, where_clause):
    """
    where_clause: string unica com toda a condicao, ou lista de strings
    Se for lista, junta com AND
    """
    try:
        if isinstance(where_clause, list):
            where_clause = " AND ".join(where_clause)

        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": f} for f in fields],
            OPTIONS=[{"TEXT": where_clause}],
            ROWCOUNT=0
        )
        rows = result.get("DATA", [])
        data = []
        for row in rows:
            parts = [p.strip() for p in str(row.get("WA", "")).split("|")]
            data.append({field: (parts[i] if i < len(parts) else "") for i, field in enumerate(fields)})
        return data
    except Exception as e:
        print(f"  [ERRO RFC] {e}")
        return []

class POInvestigator:
    def __init__(self, po_number):
        self.po = po_number
        self.conn = None
        self.results = {}

    def connect(self):
        self.conn = connect()

    def close(self):
        if self.conn:
            self.conn.close()

    # FASE 1: EKKO
    def fase_ekko(self):
        print(f"\n{'='*110}")
        print(f"FASE 1: EKKO - Cabecalho da PO {self.po}")
        print(f"{'='*110}")

        fields = ["EBELN", "BUKRS", "WAERS", "LIFNR", "BEDAT", "AEDAT", "ZZCOORI", "ZZEXPVZ"]

        ekko = read_table(self.conn, "EKKO", fields, f"EBELN = '{self.po}'")

        if not ekko:
            print(f"[FALHA] PO {self.po} nao existe em EKKO")
            self.results["ekko"] = {"status": "FALHA", "data": None}
            return None

        data = ekko[0]
        print(f"[PASSA] PO encontrada em EKKO")
        print(f"  BUKRS:   {data.get('BUKRS')}")
        print(f"  WAERS:   {data.get('WAERS')}")
        print(f"  EBELN:   {data.get('EBELN')}")
        print(f"  BEDAT:   {data.get('BEDAT')}")
        print(f"  AEDAT:   {data.get('AEDAT')}")
        print(f"  LIFNR:   {data.get('LIFNR')}")
        zzcoori = data.get('ZZCOORI', '').strip()
        zzexpvz = data.get('ZZEXPVZ', '').strip()
        print(f"  ZZCOORI: '{zzcoori}' {'[VAZIO]' if not zzcoori else ''}")
        print(f"  ZZEXPVZ: '{zzexpvz}' {'[VAZIO]' if not zzexpvz else ''}")

        self.results["ekko"] = {"status": "PASSA", "data": data}
        return data

    # FASE 2: EKET
    def fase_eket(self, ekko_data):
        print(f"\n{'='*110}")
        print(f"FASE 2: EKET - Programacao de Entrega")
        print(f"{'='*110}")

        fields = ["EBELN", "EBELP", "ETENR", "EINDT", "MENGE", "WEMNG"]

        eket = read_table(self.conn, "EKET", fields, f"EBELN = '{self.po}'")

        if not eket:
            print(f"[FALHA] Nenhuma posicao em EKET")
            print(f"  >>> INNER JOIN EKET elimina a PO")
            self.results["eket"] = {"status": "FALHA", "data": None}
            return None

        print(f"[PASSA] {len(eket)} posicao(oes) encontrada(s)")
        for i, row in enumerate(eket, 1):
            print(f"  Posicao {i}: EBELP={row.get('EBELP')}, EINDT={row.get('EINDT')}, MENGE={row.get('MENGE')}")

        self.results["eket"] = {"status": "PASSA", "data": eket}
        return eket

    # FASE 3: LFB1
    def fase_lfb1(self, ekko_data):
        print(f"\n{'='*110}")
        print(f"FASE 3: LFB1 - Fornecedor por Empresa")
        print(f"{'='*110}")

        lifnr = ekko_data.get('LIFNR')
        bukrs = ekko_data.get('BUKRS')

        print(f"Procurando: LIFNR={lifnr}, BUKRS={bukrs}")

        fields = ["LIFNR", "BUKRS", "ZTERM"]

        where = f"LIFNR = '{lifnr}' AND BUKRS = '{bukrs}'"
        lfb1 = read_table(self.conn, "LFB1", fields, where)

        if not lfb1:
            print(f"[FALHA] Nenhuma correspondencia em LFB1")
            print(f"  >>> INNER JOIN LFB1 elimina a PO")
            self.results["lfb1"] = {"status": "FALHA", "data": None}
            return None

        data = lfb1[0]
        print(f"[PASSA] Fornecedor encontrado")
        print(f"  LIFNR: {data.get('LIFNR')}")
        print(f"  BUKRS: {data.get('BUKRS')}")
        print(f"  ZTERM: {data.get('ZTERM')}")

        self.results["lfb1"] = {"status": "PASSA", "data": data}
        return data

    # FASE 4: ZFI_PAY_DATE_T
    def fase_zfi_pay_date_t(self, ekko_data):
        print(f"\n{'='*110}")
        print(f"FASE 4: ZFI_PAY_DATE_T - Configuracao Critica")
        print(f"{'='*110}")

        zzcoori = ekko_data.get('ZZCOORI', '').strip()
        zzexpvz = ekko_data.get('ZZEXPVZ', '').strip()

        print(f"Procurando: ZZCOORI='{zzcoori}', ZZEXPVZ='{zzexpvz}'")

        if not zzcoori or not zzexpvz:
            print(f"[ATENCAO] Um ou ambos os campos estao vazios")

        fields = ["ZZCOORI", "ZZEXPVZ", "NDAYS"]

        where = f"ZZCOORI = '{zzcoori}' AND ZZEXPVZ = '{zzexpvz}'"
        zfi = read_table(self.conn, "ZFI_PAY_DATE_T", fields, where)

        if not zfi:
            print(f"[FALHA] Nenhuma correspondencia em ZFI_PAY_DATE_T")
            print(f"  >>> INNER JOIN ZFI_PAY_DATE_T elimina a PO")
            print(f"  >>> CAUSA RAIZ PROVAVEL")
            self.results["zfi_pay_date_t"] = {"status": "FALHA", "data": None}
            return None

        data = zfi[0]
        print(f"[PASSA] Configuracao encontrada")
        print(f"  ZZCOORI: {data.get('ZZCOORI')}")
        print(f"  ZZEXPVZ: {data.get('ZZEXPVZ')}")
        print(f"  NDAYS:   {data.get('NDAYS')}")

        self.results["zfi_pay_date_t"] = {"status": "PASSA", "data": data}
        return data

    # FASE 5: T052
    def fase_t052(self, lfb1_data):
        print(f"\n{'='*110}")
        print(f"FASE 5: T052 - Condicao de Pagamento")
        print(f"{'='*110}")

        zterm = lfb1_data.get('ZTERM')

        print(f"Procurando: ZTERM={zterm}")

        fields = ["ZTERM", "ZTAG1"]

        t052 = read_table(self.conn, "T052", fields, f"ZTERM = '{zterm}'")

        if not t052:
            print(f"[FALHA] ZTERM nao encontrado em T052")
            print(f"  >>> INNER JOIN T052 elimina a PO")
            self.results["t052"] = {"status": "FALHA", "data": None}
            return None

        data = t052[0]
        print(f"[PASSA] Condicao encontrada")
        print(f"  ZTERM: {data.get('ZTERM')}")
        print(f"  ZTAG1: {data.get('ZTAG1')}")

        self.results["t052"] = {"status": "PASSA", "data": data}
        return data

    # FASE 6: EKPO
    def fase_ekpo(self, eket_data):
        print(f"\n{'='*110}")
        print(f"FASE 6: EKPO - Posicoes")
        print(f"{'='*110}")

        fields = ["EBELN", "EBELP", "PSTYP", "NETWR", "ZZDAT02", "ZZDAT03"]

        ekpo_list = []
        for eket_row in eket_data:
            ebelp = eket_row.get('EBELP')

            where = f"EBELN = '{self.po}' AND EBELP = '{ebelp}'"
            ekpo = read_table(self.conn, "EKPO", fields, where)

            if not ekpo:
                print(f"[FALHA] Posicao {ebelp} nao encontrada em EKPO")
                print(f"  >>> INNER JOIN EKPO elimina a PO")
                self.results["ekpo"] = {"status": "FALHA", "data": None}
                return None

            data = ekpo[0]
            ekpo_list.append(data)
            pstyp = data.get('PSTYP')
            pstyp_desc = "[SERVICO]" if pstyp == "5" else "[MATERIAL]"
            print(f"[PASSA] EBELP={ebelp}, PSTYP={pstyp} {pstyp_desc}, NETWR={data.get('NETWR')}")

        self.results["ekpo"] = {"status": "PASSA", "data": ekpo_list}
        return ekpo_list

    # FASE 7: SELECT principal
    def fase_select_principal(self):
        print(f"\n{'='*110}")
        print(f"FASE 7: SELECT Principal (LT_AUTO)")
        print(f"{'='*110}")

        # Se chegou aqui, a PO passa por todos os JOINs
        if all(self.results.get(k, {}).get("status") == "PASSA" for k in
               ["ekko", "eket", "lfb1", "zfi_pay_date_t", "t052", "ekpo"]):
            print(f"[PASSA] PO entra em LT_AUTO (todos os JOINs passaram)")
            print(f"  >>> PO passou pela clausula SELECT completa")
            self.results["select_principal"] = {"status": "PASSA"}
            return True
        else:
            print(f"[FALHA] PO nao entra em LT_AUTO")
            self.results["select_principal"] = {"status": "FALHA"}
            return False

    # FASE 8: LOG
    def fase_log(self):
        print(f"\n{'='*110}")
        print(f"FASE 8: ZFI_DOC_EX_LOG_T - Log de Processamento")
        print(f"{'='*110}")

        ekko_data = self.results.get("ekko", {}).get("data")
        if not ekko_data:
            print(f"[SKIP] Sem dados EKKO")
            return None

        bukrs = ekko_data.get('BUKRS')

        fields = ["EBELN", "BUKRS"]

        where = f"EBELN = '{self.po}' AND BUKRS = '{bukrs}'"
        log = read_table(self.conn, "ZFI_DOC_EX_LOG_T", fields, where)

        if not log:
            print(f"[PASSA] PO nao esta no log (pode ser processada)")
            print(f"  >>> PO nao foi removida por existencia anterior")
            self.results["log"] = {"status": "PASSA", "data": None}
            return None
        else:
            print(f"[AVISO] PO existe no log")
            print(f"  Registos encontrados: {len(log)}")
            print(f"  >>> PO ja foi processada e removida de LT_AUTO")
            self.results["log"] = {"status": "AVISO", "data": log}
            return log

    def run(self):
        print(f"\n{'#'*110}")
        print(f"# INVESTIGACAO COMPLETA: PO {self.po}")
        print(f"# Modo: 100% SOMENTE LEITURA")
        print(f"{'#'*110}")

        try:
            self.connect()

            # Executar fases
            ekko = self.fase_ekko()
            if not ekko:
                self.diagnostico_final("PO nao existe no sistema")
                return

            eket = self.fase_eket(ekko)
            if not eket:
                self.diagnostico_final("PO falha em EKET (sem posicoes)")
                return

            lfb1 = self.fase_lfb1(ekko)
            if not lfb1:
                self.diagnostico_final("PO falha em LFB1 (fornecedor nao encontrado)")
                return

            zfi = self.fase_zfi_pay_date_t(ekko)
            if not zfi:
                self.diagnostico_final("PO falha em ZFI_PAY_DATE_T (configuracao nao encontrada) - CAUSA RAIZ")
                return

            t052 = self.fase_t052(lfb1)
            if not t052:
                self.diagnostico_final("PO falha em T052 (condicao pagamento nao encontrada)")
                return

            ekpo = self.fase_ekpo(eket)
            if not ekpo:
                self.diagnostico_final("PO falha em EKPO (posicao nao encontrada)")
                return

            self.fase_select_principal()
            self.fase_log()

        finally:
            self.close()

    def diagnostico_final(self, msg):
        print(f"\n{'='*110}")
        print(f"[DIAGNOSTICO] {msg}")
        print(f"{'='*110}")


def main():
    pos = ["4300000003", "4000055997"]

    for po in pos:
        investigator = POInvestigator(po)
        investigator.run()
        print("\n\n")


if __name__ == "__main__":
    main()

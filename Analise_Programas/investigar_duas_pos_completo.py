"""
Investigação Completa: Por que POs não aparecem em AUTO_INST_ASSIGN
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

def read_table(conn, table, fields, options):
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": f} for f in fields],
            OPTIONS=[{"TEXT": o} for o in options],
            ROWCOUNT=0
        )
        rows = result.get("DATA", [])
        data = []
        for row in rows:
            parts = [p.strip() for p in str(row.get("WA", "")).split("|")]
            data.append({field: (parts[i] if i < len(parts) else "") for i, field in enumerate(fields)})
        return data
    except Exception as e:
        print(f"  ⚠️ ERRO: {e}")
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
        print(f"\n{'='*120}")
        print(f"FASE 1: EKKO — Cabeçalho da PO {self.po}")
        print(f"{'='*120}")

        fields = ["EBELN", "BUKRS", "WAERS", "LIFNR", "BEDAT", "AEDAT", "ZZCOORI", "ZZEXPVZ",
                  "BSART", "BSTYP", "LOEKZ", "STATU", "WKURS", "KUFIX"]

        ekko = read_table(self.conn, "EKKO", fields, [f"EBELN = '{self.po}'"])

        if not ekko:
            print(f"❌ FALHA: PO {self.po} não existe em EKKO")
            self.results["ekko"] = {"status": "FALHA", "data": None}
            return None

        data = ekko[0]
        print(f"✅ PO encontrada em EKKO")
        print(f"\n[Campos Utilizados por AUTO_INST_ASSIGN]")
        print(f"  BUKRS:    {data.get('BUKRS')}")
        print(f"  WAERS:    {data.get('WAERS')}")
        print(f"  EBELN:    {data.get('EBELN')}")
        print(f"  BEDAT:    {data.get('BEDAT')}")
        print(f"  AEDAT:    {data.get('AEDAT')}")
        print(f"  LIFNR:    {data.get('LIFNR')}")
        print(f"  ZZCOORI:  '{data.get('ZZCOORI')}' {'[VAZIO]' if not data.get('ZZCOORI') else ''}")
        print(f"  ZZEXPVZ:  '{data.get('ZZEXPVZ')}' {'[VAZIO]' if not data.get('ZZEXPVZ') else ''}")

        print(f"\n[Campos Informativos (não filtrados)]")
        print(f"  BSART:    {data.get('BSART')}")
        print(f"  BSTYP:    {data.get('BSTYP')}")
        print(f"  LOEKZ:    {data.get('LOEKZ')}")
        print(f"  STATU:    {data.get('STATU')}")
        print(f"  WKURS:    {data.get('WKURS')}")
        print(f"  KUFIX:    {data.get('KUFIX')}")

        self.results["ekko"] = {"status": "PASSA", "data": data}
        return data

    # FASE 2: EKET
    def fase_eket(self, ekko_data):
        print(f"\n{'='*120}")
        print(f"FASE 2: EKET — Programação de Entrega")
        print(f"{'='*120}")

        fields = ["EBELN", "EBELP", "ETENR", "EINDT", "MENGE", "WEMNG"]

        eket = read_table(self.conn, "EKET", fields, [f"EBELN = '{self.po}'"])

        if not eket:
            print(f"❌ FALHA: Nenhuma posição em EKET")
            print(f"   INNER JOIN EKET elimina a PO")
            self.results["eket"] = {"status": "FALHA", "data": None}
            return None

        print(f"✅ PASSA: {len(eket)} posição(ões) encontrada(s)")
        for i, row in enumerate(eket, 1):
            print(f"\n  Posição {i}:")
            print(f"    EBELP: {row.get('EBELP')}")
            print(f"    ETENR: {row.get('ETENR')}")
            print(f"    EINDT: {row.get('EINDT')}")
            print(f"    MENGE: {row.get('MENGE')}")
            print(f"    WEMNG: {row.get('WEMNG')}")

        self.results["eket"] = {"status": "PASSA", "data": eket}
        return eket

    # FASE 3: LFB1
    def fase_lfb1(self, ekko_data):
        print(f"\n{'='*120}")
        print(f"FASE 3: LFB1 — Fornecedor por Empresa")
        print(f"{'='*120}")

        lifnr = ekko_data.get('LIFNR')
        bukrs = ekko_data.get('BUKRS')

        print(f"Procurando: LIFNR={lifnr}, BUKRS={bukrs}")

        fields = ["LIFNR", "BUKRS", "ZTERM"]

        lfb1 = read_table(self.conn, "LFB1", fields,
                         [f"LIFNR = '{lifnr}'", f"BUKRS = '{bukrs}'"])

        if not lfb1:
            print(f"❌ FALHA: Nenhuma correspondência em LFB1")
            print(f"   INNER JOIN LFB1 elimina a PO")
            self.results["lfb1"] = {"status": "FALHA", "data": None}
            return None

        data = lfb1[0]
        print(f"✅ PASSA: Fornecedor encontrado")
        print(f"  LIFNR: {data.get('LIFNR')}")
        print(f"  BUKRS: {data.get('BUKRS')}")
        print(f"  ZTERM: {data.get('ZTERM')}")

        self.results["lfb1"] = {"status": "PASSA", "data": data}
        return data

    # FASE 4: ZFI_PAY_DATE_T
    def fase_zfi_pay_date_t(self, ekko_data):
        print(f"\n{'='*120}")
        print(f"FASE 4: ZFI_PAY_DATE_T — Configuração Crítica")
        print(f"{'='*120}")

        zzcoori = ekko_data.get('ZZCOORI')
        zzexpvz = ekko_data.get('ZZEXPVZ')

        print(f"Procurando: ZZCOORI='{zzcoori}', ZZEXPVZ='{zzexpvz}'")

        if not zzcoori or not zzexpvz:
            print(f"⚠️ ATENÇÃO: Um ou ambos os campos estão vazios")

        fields = ["ZZCOORI", "ZZEXPVZ", "NDAYS"]

        zfi = read_table(self.conn, "ZFI_PAY_DATE_T", fields,
                        [f"ZZCOORI = '{zzcoori}'", f"ZZEXPVZ = '{zzexpvz}'"])

        if not zfi:
            print(f"❌ FALHA: Nenhuma correspondência em ZFI_PAY_DATE_T")
            print(f"   INNER JOIN ZFI_PAY_DATE_T elimina a PO")
            self.results["zfi_pay_date_t"] = {"status": "FALHA", "data": None}
            return None

        data = zfi[0]
        print(f"✅ PASSA: Configuração encontrada")
        print(f"  ZZCOORI: {data.get('ZZCOORI')}")
        print(f"  ZZEXPVZ: {data.get('ZZEXPVZ')}")
        print(f"  NDAYS:   {data.get('NDAYS')}")

        self.results["zfi_pay_date_t"] = {"status": "PASSA", "data": data}
        return data

    # FASE 5: T052
    def fase_t052(self, lfb1_data):
        print(f"\n{'='*120}")
        print(f"FASE 5: T052 — Condição de Pagamento")
        print(f"{'='*120}")

        zterm = lfb1_data.get('ZTERM')

        print(f"Procurando: ZTERM={zterm}")

        fields = ["ZTERM", "ZTAG1"]

        t052 = read_table(self.conn, "T052", fields, [f"ZTERM = '{zterm}'"])

        if not t052:
            print(f"❌ FALHA: ZTERM não encontrado em T052")
            print(f"   INNER JOIN T052 elimina a PO")
            self.results["t052"] = {"status": "FALHA", "data": None}
            return None

        data = t052[0]
        print(f"✅ PASSA: Condição encontrada")
        print(f"  ZTERM: {data.get('ZTERM')}")
        print(f"  ZTAG1: {data.get('ZTAG1')}")

        self.results["t052"] = {"status": "PASSA", "data": data}
        return data

    # FASE 6: EKPO
    def fase_ekpo(self, eket_data):
        print(f"\n{'='*120}")
        print(f"FASE 6: EKPO — Posições")
        print(f"{'='*120}")

        fields = ["EBELN", "EBELP", "PSTYP", "NETWR", "ZZDAT02", "ZZDAT03", "LOEKZ"]

        ekpo_list = []
        for eket_row in eket_data:
            ebelp = eket_row.get('EBELP')
            print(f"\nProcurando posição: EBELP={ebelp}")

            ekpo = read_table(self.conn, "EKPO", fields,
                             [f"EBELN = '{self.po}'", f"EBELP = '{ebelp}'"])

            if not ekpo:
                print(f"  ❌ FALHA: Posição não encontrada em EKPO")
                self.results["ekpo"] = {"status": "FALHA", "data": None}
                return None

            data = ekpo[0]
            ekpo_list.append(data)
            print(f"  ✅ PASSA: Posição encontrada")
            print(f"    PSTYP:    {data.get('PSTYP')} {'[SERVIÇO]' if data.get('PSTYP') == '5' else '[MATERIAL]'}")
            print(f"    NETWR:    {data.get('NETWR')}")
            print(f"    ZZDAT02:  {data.get('ZZDAT02')}")
            print(f"    ZZDAT03:  {data.get('ZZDAT03')}")
            print(f"    LOEKZ:    {data.get('LOEKZ')}")

        self.results["ekpo"] = {"status": "PASSA", "data": ekpo_list}
        return ekpo_list

    # FASE 7: SELECT principal
    def fase_select_principal(self):
        print(f"\n{'='*120}")
        print(f"FASE 7: SELECT Principal (LT_AUTO)")
        print(f"{'='*120}")

        # Se chegou aqui, a PO passa por todos os JOINs
        if all(self.results.get(k, {}).get("status") == "PASSA" for k in
               ["ekko", "eket", "lfb1", "zfi_pay_date_t", "t052", "ekpo"]):
            print(f"✅ PASSA: PO entra em LT_AUTO (todos os JOINs passaram)")
            self.results["select_principal"] = {"status": "PASSA"}
            return True
        else:
            print(f"❌ FALHA: PO não entra em LT_AUTO")
            self.results["select_principal"] = {"status": "FALHA"}
            return False

    # FASE 8: LOG
    def fase_log(self):
        print(f"\n{'='*120}")
        print(f"FASE 8: ZFI_DOC_EX_LOG_T — Log de Processamento")
        print(f"{'='*120}")

        ekko_data = self.results.get("ekko", {}).get("data")
        if not ekko_data:
            print(f"⚠️ SKIP: Sem dados EKKO")
            return None

        bukrs = ekko_data.get('BUKRS')

        fields = ["EBELN", "BUKRS"]

        log = read_table(self.conn, "ZFI_DOC_EX_LOG_T", fields,
                        [f"EBELN = '{self.po}'", f"BUKRS = '{bukrs}'"])

        if not log:
            print(f"✅ PASSA: PO não está no log (pode ser processada)")
            self.results["log"] = {"status": "PASSA", "data": None}
            return None
        else:
            print(f"⚠️ AVISO: PO existe no log")
            print(f"   Registos encontrados: {len(log)}")
            self.results["log"] = {"status": "AVISO", "data": log}
            return log

    # FASE 9: PAY_DATE
    def fase_pay_date(self):
        print(f"\n{'='*120}")
        print(f"FASE 9: PAY_DATE — Cálculo")
        print(f"{'='*120}")

        ekpo_data = self.results.get("ekpo", {}).get("data")
        zfi_data = self.results.get("zfi_pay_date_t", {}).get("data")
        t052_data = self.results.get("t052", {}).get("data")

        if not (ekpo_data and zfi_data and t052_data):
            print(f"⚠️ SKIP: Dados insuficientes para cálculo")
            return None

        print(f"[Cálculo: PAY_DATE = DATA_BASE - NDAYS + ZTAG1]")

        for ekpo in (ekpo_data if isinstance(ekpo_data, list) else [ekpo_data]):
            zzdat02 = ekpo.get('ZZDAT02')
            zzdat03 = ekpo.get('ZZDAT03')
            data_base = zzdat03 if zzdat03 else zzdat02
            ndays = zfi_data.get('NDAYS', '0')
            ztag1 = t052_data.get('ZTAG1', '0')

            print(f"\nPosição {ekpo.get('EBELP')}:")
            print(f"  ZZDAT02:  {zzdat02}")
            print(f"  ZZDAT03:  {zzdat03}")
            print(f"  DATA_BASE (usada): {data_base}")
            print(f"  NDAYS:    {ndays}")
            print(f"  ZTAG1:    {ztag1}")
            print(f"  → PAY_DATE = {data_base} (base de cálculo identificada)")

        self.results["pay_date"] = {"status": "INFO"}

    # FASE 10: Intervalo PAY_DATE
    def fase_intervalo_pay_date(self):
        print(f"\n{'='*120}")
        print(f"FASE 10: Intervalo PAY_DATE")
        print(f"{'='*120}")

        print(f"⚠️ NOTA: ID_IDATE e ID_EDATE precisam ser fornecidos pelo utilizador")
        print(f"          ou extraídos do contexto de execução do programa")
        print(f"\nSem estes valores, não é possível validar se PAY_DATE está dentro do intervalo")

        self.results["intervalo_pay_date"] = {"status": "INFO"}

    def run(self):
        print(f"\n\n{'#'*120}")
        print(f"# INVESTIGAÇÃO COMPLETA: PO {self.po}")
        print(f"# Modo: 100% SOMENTE LEITURA")
        print(f"{'#'*120}")

        try:
            self.connect()

            # Executar fases
            ekko = self.fase_ekko()
            if not ekko:
                return

            eket = self.fase_eket(ekko)
            if not eket:
                return

            lfb1 = self.fase_lfb1(ekko)
            if not lfb1:
                return

            zfi = self.fase_zfi_pay_date_t(ekko)
            if not zfi:
                return

            t052 = self.fase_t052(lfb1)
            if not t052:
                return

            ekpo = self.fase_ekpo(eket)
            if not ekpo:
                return

            self.fase_select_principal()
            self.fase_log()
            self.fase_pay_date()
            self.fase_intervalo_pay_date()

            # Relatório final
            self.relatorio_final()

        finally:
            self.close()

    def relatorio_final(self):
        print(f"\n\n{'='*120}")
        print(f"RELATÓRIO FINAL: PO {self.po}")
        print(f"{'='*120}")

        print(f"\n| Fase | Critério | Valor Encontrado | Status |")
        print(f"|---|---|---|---|")

        for fase, result in self.results.items():
            status = result.get("status", "?")
            print(f"| {fase} | — | — | {status} |")


def main():
    pos = ["4300000003", "4000055997"]

    for po in pos:
        investigator = POInvestigator(po)
        investigator.run()
        print("\n\n")


if __name__ == "__main__":
    main()

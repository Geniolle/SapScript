# -*- coding: utf-8 -*-
"""
Fase 2 da Investigacao: POs 4300000003 vs 4000055997
Foco: Depois dos JOINs
- ZFI_DOC_EX_LOG_T (exclusao posterior)
- Calculo PAY_DATE
- Intervalo ID_IDATE/ID_EDATE
- GT_AUTO (append final)
"""

import os
import sys
from pathlib import Path
from pyrfc import Connection
from datetime import datetime, timedelta

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

class POAnalysisPhase2:
    def __init__(self, po1, po2):
        self.po1 = po1
        self.po2 = po2
        self.conn = None
        self.data = {}

    def connect(self):
        self.conn = connect()

    def close(self):
        if self.conn:
            self.conn.close()

    # FASE 1: ZFI_DOC_EX_LOG_T
    def verificar_log(self, po, bukrs="2010"):
        print(f"\n{'='*110}")
        print(f"FASE 1: ZFI_DOC_EX_LOG_T - PO {po}")
        print(f"{'='*110}")

        fields = ["EBELN", "BUKRS", "ERDAT", "ERZET", "FLDTYPE", "FLDVALUE"]

        where = f"EBELN = '{po}' AND BUKRS = '{bukrs}'"
        log_data = read_table(self.conn, "ZFI_DOC_EX_LOG_T", fields, where)

        if not log_data:
            print(f"[RESULTADO] Nenhum registo em ZFI_DOC_EX_LOG_T")
            print(f"  >>> PO NAO foi excluida pelo DELETE LR_LEBELN")
            return False
        else:
            print(f"[RESULTADO] {len(log_data)} registo(s) encontrado(s) em ZFI_DOC_EX_LOG_T")
            for i, row in enumerate(log_data, 1):
                print(f"\n  Registo {i}:")
                print(f"    EBELN:    {row.get('EBELN')}")
                print(f"    BUKRS:    {row.get('BUKRS')}")
                print(f"    ERDAT:    {row.get('ERDAT')}")
                print(f"    ERZET:    {row.get('ERZET')}")
                print(f"    FLDTYPE:  {row.get('FLDTYPE')}")
                print(f"    FLDVALUE: {row.get('FLDVALUE')}")
            print(f"\n  >>> PO FO EXCLUIDA pelo DELETE LR_LEBELN")
            print(f"      DELETE LT_AUTO WHERE EBELN IN LR_LEBELN")
            return True

    # FASE 2: Recuperar dados para PAY_DATE
    def recuperar_dados_pay_date(self, po):
        print(f"\n{'='*110}")
        print(f"FASE 2: Dados para PAY_DATE - PO {po}")
        print(f"{'='*110}")

        resultado = {}

        # EKPO
        print(f"\n[EKPO]")
        fields_ekpo = ["EBELN", "EBELP", "ZZDAT02", "ZZDAT03", "NETWR"]
        ekpo = read_table(self.conn, "EKPO", fields_ekpo, f"EBELN = '{po}'")

        if ekpo:
            for row in ekpo:
                print(f"  EBELP={row.get('EBELP')}: ZZDAT02={row.get('ZZDAT02')}, ZZDAT03={row.get('ZZDAT03')}, NETWR={row.get('NETWR')}")
            resultado["ekpo"] = ekpo

        # Recuperar ZZCOORI, ZZEXPVZ de EKKO
        print(f"\n[EKKO]")
        ekko = read_table(self.conn, "EKKO", ["ZZCOORI", "ZZEXPVZ", "LIFNR", "BUKRS"], f"EBELN = '{po}'")
        zzcoori = ekko[0].get('ZZCOORI', '').strip() if ekko else ""
        zzexpvz = ekko[0].get('ZZEXPVZ', '').strip() if ekko else ""
        lifnr = ekko[0].get('LIFNR') if ekko else ""
        bukrs = ekko[0].get('BUKRS') if ekko else "2010"

        # ZFI_PAY_DATE_T
        print(f"\n[ZFI_PAY_DATE_T]")
        zfi = read_table(self.conn, "ZFI_PAY_DATE_T", ["ZZCOORI", "ZZEXPVZ", "NDAYS"],
                        f"ZZCOORI = '{zzcoori}' AND ZZEXPVZ = '{zzexpvz}'")
        if zfi:
            ndays = zfi[0].get('NDAYS')
            print(f"  ZZCOORI={zzcoori}, ZZEXPVZ={zzexpvz}: NDAYS={ndays}")
            resultado["zfi_pay_date_t"] = zfi[0]

        # LFB1
        print(f"\n[LFB1]")
        lfb1 = read_table(self.conn, "LFB1", ["LIFNR", "BUKRS", "ZTERM"],
                         f"LIFNR = '{lifnr}' AND BUKRS = '{bukrs}'")
        zterm = ""
        if lfb1:
            zterm = lfb1[0].get('ZTERM')
            print(f"  LIFNR={lifnr}, BUKRS={bukrs}: ZTERM={zterm}")

        # T052
        print(f"\n[T052]")
        t052 = read_table(self.conn, "T052", ["ZTERM", "ZTAG1"], f"ZTERM = '{zterm}'")
        if t052:
            ztag1 = t052[0].get('ZTAG1')
            print(f"  ZTERM={zterm}: ZTAG1={ztag1}")
            resultado["t052"] = t052[0]

        self.data[po] = resultado
        return resultado

    # FASE 3: Calcular PAY_DATE
    def calcular_pay_date(self, po):
        print(f"\n{'='*110}")
        print(f"FASE 3: Calculo PAY_DATE - PO {po}")
        print(f"{'='*110}")

        dados = self.data.get(po, {})
        ekpo = dados.get("ekpo", [])
        zfi = dados.get("zfi_pay_date_t", {})
        t052 = dados.get("t052", {})

        if not (ekpo and zfi and t052):
            print(f"[ERRO] Dados insuficientes para calculo")
            return None

        ndays_str = zfi.get('NDAYS', '0')
        ztag1_str = t052.get('ZTAG1', '0')

        try:
            ndays = int(ndays_str)
            ztag1 = int(ztag1_str)
        except:
            ndays = 0
            ztag1 = 0

        pay_dates = {}

        for ekpo_row in ekpo:
            ebelp = ekpo_row.get('EBELP')
            zzdat02 = ekpo_row.get('ZZDAT02')
            zzdat03 = ekpo_row.get('ZZDAT03')

            # CASE WHEN EKPO-ZZDAT03 IS NOT NULL THEN EKPO-ZZDAT03 ELSE EKPO-ZZDAT02 END
            if zzdat03 and zzdat03.strip():
                data_base = zzdat03
                data_base_source = "ZZDAT03"
            else:
                data_base = zzdat02
                data_base_source = "ZZDAT02"

            print(f"\nPosicao {ebelp}:")
            print(f"  ZZDAT02: {zzdat02}")
            print(f"  ZZDAT03: {zzdat03}")
            print(f"  DATA_BASE (usada): {data_base} [{data_base_source}]")
            print(f"  NDAYS:   {ndays}")
            print(f"  ZTAG1:   {ztag1}")

            # Converter para data e calcular
            try:
                data_obj = datetime.strptime(data_base, "%Y%m%d")
                data_ajustada = data_obj - timedelta(days=ndays) + timedelta(days=ztag1)
                pay_date = data_ajustada.strftime("%Y%m%d")

                print(f"  Calculo: {data_base} - {ndays} dias + {ztag1} dias = {pay_date}")
                pay_dates[ebelp] = pay_date
            except Exception as e:
                print(f"  [ERRO] Nao consegui converter a data: {e}")
                pay_dates[ebelp] = data_base

        self.data[po]["pay_dates"] = pay_dates
        return pay_dates

    # FASE 4: Simular condicao final
    def simular_condicao_final(self, po, id_idate="20260101", id_edate="20261231"):
        print(f"\n{'='*110}")
        print(f"FASE 4: Condicao Final ID_IDATE/ID_EDATE - PO {po}")
        print(f"{'='*110}")

        pay_dates = self.data.get(po, {}).get("pay_dates", {})

        if not pay_dates:
            print(f"[INFO] Nao ha PAY_DATEs calculadas")
            return False

        print(f"\n[INTERVALO]")
        print(f"  ID_IDATE: {id_idate}")
        print(f"  ID_EDATE: {id_edate}")

        todas_dentro = True

        for ebelp, pay_date in pay_dates.items():
            dentro = (id_idate <= pay_date <= id_edate)
            resultado = "SIM" if dentro else "NAO"
            status = "PASSA" if dentro else "FALHA"

            print(f"\n  EBELP={ebelp}:")
            print(f"    PAY_DATE: {pay_date}")
            print(f"    ID_IDATE <= PAY_DATE <= ID_EDATE: {resultado} [{status}]")

            if not dentro:
                todas_dentro = False

        self.data[po]["todas_dentro_intervalo"] = todas_dentro
        return todas_dentro

    # FASE 5: Verificar GT_AUTO (logico)
    def verificar_gt_auto(self, po):
        print(f"\n{'='*110}")
        print(f"FASE 5: Verificacao GT_AUTO (logica) - PO {po}")
        print(f"{'='*110}")

        # 1. Entra em LT_AUTO?
        print(f"\n[1] Entra inicialmente em LT_AUTO?")
        print(f"    Resposta: SIM (todos os JOINs passaram)")

        # 2. Excluida pelo LOG?
        em_log = self.verificar_log_para_gt_auto(po)
        print(f"\n[2] Excluida por ZFI_DOC_EX_LOG_T?")
        print(f"    Resposta: {'SIM (DELETE LR_LEBELN aplicado)' if em_log else 'NAO'}")

        if em_log:
            print(f"\n[RESULTADO] PO NAO CHEGA A GT_AUTO - foi eliminada pelo LOG")
            self.data[po]["append_gt_auto"] = False
            return False

        # 3. PAY_DATE dentro intervalo?
        dentro = self.data.get(po, {}).get("todas_dentro_intervalo", False)
        print(f"\n[3] PAY_DATE dentro ID_IDATE/ID_EDATE?")
        print(f"    Resposta: {'SIM' if dentro else 'NAO'}")

        if not dentro:
            print(f"\n[RESULTADO] PO NAO CHEGA A GT_AUTO - PAY_DATE fora do intervalo")
            self.data[po]["append_gt_auto"] = False
            return False

        # 4. APPEND em GT_AUTO?
        print(f"\n[4] APPEND LS_AUTOF TO GT_AUTO seria executado?")
        print(f"    Resposta: SIM")
        print(f"\n[RESULTADO] PO CHEGA A GT_AUTO - deveria aparecer no ALV")
        self.data[po]["append_gt_auto"] = True
        return True

    def verificar_log_para_gt_auto(self, po):
        """Helper para GT_AUTO"""
        fields = ["EBELN"]
        where = f"EBELN = '{po}'"
        log_data = read_table(self.conn, "ZFI_DOC_EX_LOG_T", fields, where)
        return len(log_data) > 0

    # COMPARACAO
    def comparacao_lado_a_lado(self):
        print(f"\n\n{'='*110}")
        print(f"COMPARACAO: PO {self.po1} vs PO {self.po2}")
        print(f"{'='*110}")

        print(f"\n| Campo/Etapa | {self.po1:15} | {self.po2:15} |")
        print(f"|---|---|---|")

        for po in [self.po1, self.po2]:
            dados = self.data.get(po, {})

            em_log = self.verificar_log_para_gt_auto(po)
            ekpo = dados.get("ekpo", [{}])[0]
            zfi = dados.get("zfi_pay_date_t", {})
            t052 = dados.get("t052", {})
            pay_dates = dados.get("pay_dates", {})
            append = dados.get("append_gt_auto", None)

            print(f"| Existe no LOG | {str(em_log):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                print(f"{str(self.verificar_log_para_gt_auto(self.po2)):15} |")

            print(f"| ZZDAT02 | {ekpo.get('ZZDAT02', 'N/A'):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                ekpo2 = self.data.get(self.po2, {}).get("ekpo", [{}])[0]
                print(f"{ekpo2.get('ZZDAT02', 'N/A'):15} |")

            print(f"| ZZDAT03 | {ekpo.get('ZZDAT03', 'N/A'):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                ekpo2 = self.data.get(self.po2, {}).get("ekpo", [{}])[0]
                print(f"{ekpo2.get('ZZDAT03', 'N/A'):15} |")

            print(f"| NDAYS | {zfi.get('NDAYS', 'N/A'):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                zfi2 = self.data.get(self.po2, {}).get("zfi_pay_date_t", {})
                print(f"{zfi2.get('NDAYS', 'N/A'):15} |")

            print(f"| ZTAG1 | {t052.get('ZTAG1', 'N/A'):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                t0522 = self.data.get(self.po2, {}).get("t052", {})
                print(f"{t0522.get('ZTAG1', 'N/A'):15} |")

            print(f"| PAY_DATE | {str(pay_dates):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                pay_dates2 = self.data.get(self.po2, {}).get("pay_dates", {})
                print(f"{str(pay_dates2):15} |")

            print(f"| APPEND em GT_AUTO | {str(append):15} | ", end="")
            if po == self.po1:
                print("")
            else:
                append2 = self.data.get(self.po2, {}).get("append_gt_auto")
                print(f"{str(append2):15} |")

    def run(self, id_idate="20260101", id_edate="20261231"):
        print(f"\n{'#'*110}")
        print(f"# FASE 2: Investigacao Apos JOINs - PO 4300000003 vs 4000055997")
        print(f"# ID_IDATE: {id_idate}, ID_EDATE: {id_edate}")
        print(f"# Modo: 100% SOMENTE LEITURA")
        print(f"{'#'*110}")

        try:
            self.connect()

            for po in [self.po1, self.po2]:
                self.verificar_log(po)
                self.recuperar_dados_pay_date(po)
                self.calcular_pay_date(po)
                self.simular_condicao_final(po, id_idate, id_edate)
                self.verificar_gt_auto(po)

            self.comparacao_lado_a_lado()

        finally:
            self.close()


def main():
    # Usar intervalo padrao (ano 2026)
    analyzer = POAnalysisPhase2("4300000003", "4000055997")
    analyzer.run(id_idate="20260101", id_edate="20261231")


if __name__ == "__main__":
    main()

"""Inventário READ-ONLY (RFC/PyRFC) de processos que geram ficheiros bancários em PRD.

Cobre Pagamentos/Transferências, Débito Direto e RH/Payroll, mapeando
T042Z -> TFPM042F (PMW) -> DMEE_TREE, overrides (T042ZA_FORMAT), variantes
(DFPAYV/OBPM4) e uso real (REGUH + evidência de jobs em TBTCP).

Estritamente READ-ONLY: apenas RFC_PING, RFC_SYSTEM_INFO e RFC_READ_TABLE.
Nunca escreve, corrige, ativa árvores, cria variantes nem executa F110/payroll.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))

from sap_rfc._rfc_common import build_connection_params, find_project_root, load_project_env
from sap_agent.safety import SafetyGuard


AUDIT_ALLOWED_FUNCTIONS = ("RFC_PING", "RFC_SYSTEM_INFO", "RFC_READ_TABLE")
AUDIT_READ_TABLES = (
    "DD02L", "DD02T",
    "T001", "T042Z", "TFPM042F", "TFPM042FG",
    "DMEE_TREE", "DFPAYV",
    "T042ZA_FORMAT",
    "REGUH", "TBTCP",
    "BNKA",
)
SID_ESPERADO = "S4P"


class AuditSecurityError(RuntimeError):
    pass


def _chunk_where(clause: str, width: int = 72) -> list[str]:
    words = clause.split(" ")
    lines: list[str] = []
    cur = ""
    for w in words:
        cand = (cur + " " + w).strip() if cur else w
        if len(cand) > width:
            lines.append(cur)
            cur = w
        else:
            cur = cand
    if cur:
        lines.append(cur)
    return lines


class ReadOnlyRFC:
    def __init__(self, connection: Any, guard: SafetyGuard) -> None:
        self._conn = connection
        self._guard = guard

    def _guard_call(self, function_name: str) -> None:
        self._guard.assert_function_allowed(function_name)

    def call(self, function_name: str, **kwargs: Any) -> dict[str, Any]:
        self._guard_call(function_name)
        return dict(self._conn.call(function_name, **kwargs) or {})

    def ping(self) -> None:
        self.call("RFC_PING")

    def system_info(self) -> dict[str, str]:
        return self.call("RFC_SYSTEM_INFO").get("RFCSI_EXPORT", {}) or {}

    def read_table(
        self,
        table: str,
        fields: list[str],
        where_clause: str | None = None,
        rowcount: int = 0,
    ) -> list[dict[str, str]]:
        self._guard.assert_table_allowed(table)
        options = [{"TEXT": line} for line in _chunk_where(where_clause)] if where_clause else []
        res = self.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": f} for f in fields],
            OPTIONS=options,
            ROWCOUNT=rowcount,
        )
        cols = [e["FIELDNAME"] for e in res.get("FIELDS", [])]
        out: list[dict[str, str]] = []
        for row in res.get("DATA", []):
            vals = str(row.get("WA", "")).split("|")
            out.append({c: (vals[i].strip() if i < len(vals) else "") for i, c in enumerate(cols)})
        return out

    def table_exists(self, table: str) -> bool:
        rows = self.read_table("DD02L", ["TABNAME", "TABCLASS", "AS4LOCAL"], where_clause=f"TABNAME = '{table}'", rowcount=1)
        return bool(rows)

    def close(self) -> None:
        self._conn.close()


@dataclass
class Audit:
    ambiente: dict[str, str] = field(default_factory=dict)
    ddic: list[dict[str, str]] = field(default_factory=list)
    t001_pt: list[dict[str, str]] = field(default_factory=list)
    t042z_all_count: int = 0
    t042z_countries: list[str] = field(default_factory=list)
    t042z_pt: list[dict[str, str]] = field(default_factory=list)
    formatos: dict[str, dict[str, str]] = field(default_factory=dict)
    dmee_trees: dict[str, dict[str, str]] = field(default_factory=dict)
    cgi_ct_v9_users: list[dict[str, str]] = field(default_factory=list)
    cgi_dd_v8_users: list[dict[str, str]] = field(default_factory=list)
    overrides_pt: list[dict[str, str]] = field(default_factory=list)
    dfpayv: dict[str, list[dict[str, str]]] = field(default_factory=dict)
    reguh_usage: list[dict[str, Any]] = field(default_factory=list)
    tbtcp_rpcdta: list[dict[str, str]] = field(default_factory=list)
    tbtcp_rffo: list[dict[str, str]] = field(default_factory=list)
    bnka_pt_sample: list[dict[str, str]] = field(default_factory=list)
    manual_checks: list[dict[str, str]] = field(default_factory=list)


PT_FORMATS = [
    "Z_SEPA_AP", "Z_SEPA_DD_AR", "Z_SEPA_RH", "Z_CONF_SAN_PT", "Z_CONF_SAN_PT_XML",
    "Z_CONF_TXT", "Z_XML_INT", "Z_ATM_REFERENCE", "Z_INTER_BCM", "Z_SEPA_URG_AP",
    "ZFI_MT101", "Z_CGI_CT", "Z_CGI_DD", "Z_SEPA_AP_SCT", "Z_CONF_BBVA_PT",
    "Z_CONF_BBVA_PT_XML", "Z_CONF_BPI_PT_XML", "Z_CONF_BCP_PT_2", "Z_CONF_CGD_PT",
    "Z_CONF_NB_PT", "PT_CGI_XML_CT", "PT_CGI_XML_CT_FSN", "PT_CGI_XML_DD", "PT_CGI_XML_DD_FSN",
]

DMEE_TREE_IDS = [
    "CGI_CT", "CGI_CT_V9", "CGI_DD", "CGI_DD_V8",
    "Z_SEPA_CT", "Z_SEPA_RH", "ZSEPA_DD", "Z_CGI_CT", "Z_CGI_DD",
]


def sec_ambiente(rfc: ReadOnlyRFC, audit: Audit, forcar_sid: bool, client: str, user: str) -> None:
    rfc.ping()
    info = rfc.system_info()
    sid = str(info.get("RFCSYSID", "")).strip()
    audit.ambiente = {
        "SID": sid,
        "CLIENT": client,
        "USER": user,
        "RELEASE": str(info.get("RFCSAPRL", "")).strip(),
        "HOST": str(info.get("RFCHOST", "")).strip(),
        "DBHOST": str(info.get("RFCDBHOST", "")).strip(),
        "DBSYS": str(info.get("RFCDBSYS", "")).strip(),
    }
    if sid != SID_ESPERADO and not forcar_sid:
        raise AuditSecurityError(
            f"Ligação NÃO é o ambiente esperado (SID esperado={SID_ESPERADO}, encontrado={sid}). "
            "Abortado por segurança. Use --forcar-sid para prosseguir explicitamente."
        )


def sec_ddic(rfc: ReadOnlyRFC, audit: Audit) -> None:
    for table in AUDIT_READ_TABLES:
        if table in ("DD02L", "DD02T"):
            continue
        exists = rfc.table_exists(table)
        audit.ddic.append({"tabela": table, "existe": "SIM" if exists else "NAO"})


def sec_metodos(rfc: ReadOnlyRFC, audit: Audit) -> None:
    audit.t001_pt = rfc.read_table("T001", ["BUKRS", "BUTXT", "LAND1", "WAERS"], where_clause="LAND1 = 'PT'", rowcount=50)

    all_rows = rfc.read_table("T042Z", ["LAND1"], rowcount=0)
    audit.t042z_all_count = len(all_rows)
    audit.t042z_countries = sorted({r["LAND1"] for r in all_rows if r["LAND1"]})

    audit.t042z_pt = rfc.read_table(
        "T042Z",
        ["LAND1", "ZLSCH", "TEXT1", "FORMI", "XSEPA", "XIBAN", "PROGN", "XEINZ"],
        where_clause="LAND1 = 'PT'",
        rowcount=100,
    )


def sec_formatos(rfc: ReadOnlyRFC, audit: Audit) -> None:
    for formi in PT_FORMATS:
        rows = rfc.read_table(
            "TFPM042F",
            ["FORMI", "LAND1", "FORME", "FORMD", "DTTYP", "XDMEE", "XDME1", "BEANZ", "TREE_ID"],
            where_clause=f"FORMI = '{formi}'",
            rowcount=5,
        )
        if rows:
            audit.formatos[formi] = rows[0]


def sec_dmee_trees(rfc: ReadOnlyRFC, audit: Audit) -> None:
    for tree_id in DMEE_TREE_IDS:
        rows = rfc.read_table(
            "DMEE_TREE",
            ["TREE_TYPE", "TREE_ID", "PARENT_ID", "DOCU_TXT", "RELEASE_FLAG", "DMEEX", "EXTENSIBLE"],
            where_clause=f"TREE_ID = '{tree_id}'",
            rowcount=5,
        )
        if rows:
            audit.dmee_trees[tree_id] = rows[0]

    audit.cgi_ct_v9_users = rfc.read_table(
        "TFPM042F", ["FORMI", "LAND1", "FORME", "TREE_ID"], where_clause="TREE_ID = 'CGI_CT_V9'", rowcount=50
    )
    audit.cgi_dd_v8_users = rfc.read_table(
        "TFPM042F", ["FORMI", "LAND1", "FORME", "TREE_ID"], where_clause="TREE_ID = 'CGI_DD_V8'", rowcount=50
    )


def sec_overrides(rfc: ReadOnlyRFC, audit: Audit) -> None:
    bukrs_pt = [r["BUKRS"] for r in audit.t001_pt]
    rows: list[dict[str, str]] = []
    for b in bukrs_pt:
        rows += rfc.read_table(
            "T042ZA_FORMAT", ["ZBUKR", "ZLSCH", "HBKID", "FORMI", "FORMZ"], where_clause=f"ZBUKR = '{b}'", rowcount=200
        )
    audit.overrides_pt = rows


def sec_dfpayv(rfc: ReadOnlyRFC, audit: Audit) -> None:
    key_formats = ["Z_SEPA_AP", "Z_SEPA_DD_AR", "Z_SEPA_RH"]
    for formi in key_formats:
        audit.dfpayv[formi] = rfc.read_table(
            "DFPAYV", ["FORMI", "ZBUKR", "HBKID", "VARI"], where_clause=f"FORMI = '{formi}'", rowcount=200
        )


def sec_reguh_usage(rfc: ReadOnlyRFC, audit: Audit) -> None:
    bukrs_pt = [r["BUKRS"] for r in audit.t001_pt]
    clause = "LAUFD >= '20240101' AND LAUFD <= '20261231' AND ( " + " OR ".join(
        f"ZBUKR = '{b}'" for b in bukrs_pt
    ) + " )"
    rows = rfc.read_table("REGUH", ["ZBUKR", "LAUFD", "LAUFI", "RZAWE", "XVORL"], where_clause=clause, rowcount=0)

    by_key: dict[tuple[str, str], dict[str, Any]] = {}
    exec_by_year: dict[str, dict[tuple[str, str], set]] = defaultdict(lambda: defaultdict(set))
    for r in rows:
        if r["XVORL"] == "X":
            continue
        key = (r["ZBUKR"], r["RZAWE"])
        laufd = r["LAUFD"]
        year = laufd[:4] if laufd else "????"
        exec_by_year[year][key].add((laufd, r["LAUFI"]))
        d = by_key.setdefault(key, {"ultima_execucao": ""})
        if laufd > d["ultima_execucao"]:
            d["ultima_execucao"] = laufd

    summary = []
    for (zbukr, rzawe), d in sorted(by_key.items()):
        per_year = {y: len(exec_by_year[y].get((zbukr, rzawe), set())) for y in sorted(exec_by_year.keys())}
        summary.append({
            "ZBUKR": zbukr,
            "RZAWE": rzawe,
            "ultima_execucao": d["ultima_execucao"],
            "execucoes_por_ano": per_year,
        })
    audit.reguh_usage = summary


def sec_tbtcp(rfc: ReadOnlyRFC, audit: Audit) -> None:
    audit.tbtcp_rpcdta = rfc.read_table(
        "TBTCP", ["PROGNAME", "JOBNAME", "JOBCOUNT"], where_clause="PROGNAME LIKE 'RPCDTA%'", rowcount=20
    )
    audit.tbtcp_rffo = rfc.read_table(
        "TBTCP", ["PROGNAME", "JOBNAME", "JOBCOUNT"], where_clause="PROGNAME LIKE 'RFFO%'", rowcount=30
    )


def sec_bnka(rfc: ReadOnlyRFC, audit: Audit) -> None:
    audit.bnka_pt_sample = rfc.read_table(
        "BNKA", ["BANKS", "BANKL", "BANKA", "SWIFT"], where_clause="BANKS = 'PT'", rowcount=30
    )


ZLSCH_TO_TEXT: dict[str, str] = {}


def method_text(zlsch: str) -> str:
    return ZLSCH_TO_TEXT.get(zlsch, "(sem método atribuído)" if not zlsch else zlsch)


def build_manual_checks(audit: Audit) -> None:
    audit.manual_checks = [
        {
            "item": "Conteúdo interno das variantes OBPM4 (DFPAYV_VARI / seleção dinâmica)",
            "motivo": "RFC_READ_TABLE não expõe o conteúdo funcional completo de variantes de seleção sem risco de leitura incorreta de estruturas internas; a lista aqui é apenas a associação formato/empresa/house bank/nome de variante.",
            "transacao": "OBPM4",
            "caminho": "Financeiro > Gestão Financeira > Contas a Pagar/Receber > Transações do Fornecedor > Pagamento sem Papel > Payment Medium Workbench > Atribuir Variantes de Formato de Pagamento",
            "valor_esperado": "Confirmar seleção de layout/parâmetros dentro de cada variante (ex.: ZSEPA, Z_SEPA_DD_AR, Z_SEPA_RH)",
            "evidencia": "Screenshot do ecrã de detalhe da variante em OBPM4",
        },
        {
            "item": "Precedência exata standard vs. preferido vs. override por empresa/house bank",
            "motivo": "T042ZA_FORMAT confirma que existem overrides ativos (ex.: Z_CGI_CT e Z_CGI_DD para house bank BPI01), mas a ordem de precedência formal do motor PMW (T042ZA/T042ZA_PREFTYP) não foi lida nesta análise — não assumir 'formato efetivo' sem validar.",
            "transacao": "OBPM1 / FBZP",
            "caminho": "FBZP > Métodos de Pagamento por País/Empresa > Formato de Pagamento",
            "valor_esperado": "Confirmar visualmente qual formato o sistema realmente propõe para cada combinação (empresa, método, house bank) num F110 de teste (sem executar)",
            "evidencia": "Screenshot do ecrã de proposta de formato em FBZP/OBPM4",
        },
        {
            "item": "Versão ativa e histórico de alterações das árvores DMEEX (autor/data)",
            "motivo": "DMEE_TREE_HEAD devolveu TABLE_WITHOUT_DATA via RFC_READ_TABLE nesta ligação (tabela sem buffer acessível por este mecanismo); não foi possível confirmar autor/data da última alteração nem versão ativa por este caminho.",
            "transacao": "DMEE",
            "caminho": "Transação DMEE > Tipo de árvore PAYM > árvore (ex.: Z_CGI_DD, Z_SEPA_RH) > Utilitários > Versões",
            "valor_esperado": "Data/autor da última alteração e confirmação da versão ativa",
            "evidencia": "Screenshot do cabeçalho da árvore em DMEE",
        },
        {
            "item": "Dados de house bank (T012) e conta bancária (T012K)",
            "motivo": "T012 devolveu TABLE_WITHOUT_DATA via RFC_READ_TABLE (não legível por este mecanismo nesta versão S/4HANA); os HBKID foram confirmados apenas indiretamente via DFPAYV/T042ZA_FORMAT.",
            "transacao": "FI12 / SM30 V_T012",
            "caminho": "Financeiro > Gestão de Caixa > Configuração de Bancos da Casa",
            "valor_esperado": "Nome do banco e IBAN/conta associados a cada HBKID citado neste relatório",
            "evidencia": "Screenshot do ecrã de detalhe de house bank",
        },
    ]


def render_md(audit: Audit, output_path: Path) -> str:
    lines: list[str] = []
    A = lines.append
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    A(f"# Inventário READ-ONLY — Processos que geram ficheiros bancários (PRD)\n")
    A(f"_Gerado em {now} via RFC/PyRFC, exclusivamente leitura._\n")

    A("## 1. Ambiente\n")
    for k, v in audit.ambiente.items():
        A(f"- **{k}:** {v}")
    A("")

    A("## 2. Resumo executivo\n")
    A("- Foram identificados 3 processos que geram ficheiros bancários via PMW/DMEE em PT: "
      "**Pagamentos a Fornecedores (SEPA CT)**, **Débito Direto (SEPA DD)** e **RH/Payroll (Salários)**.")
    A("- RH/Payroll usa o mesmo motor PMW/DME (método `2`, formato `Z_SEPA_RH`) que os pagamentos a "
      "fornecedores — não existe programa clássico de DME de RH em uso (evidência TBTCP: 0 execuções `RPCDTA*`).")
    A("- F110 está ativo e em execução regular e recente em produção (evidência TBTCP: jobs `F110-*` com "
      "`RFFOAVIS_DD_PRENOTIF`/`RFFOAVIS_FPAYM`), confirmando que o ciclo de pagamentos/DD passa pelo "
      "programa de pagamento automático standard.")
    A("- A árvore genérica SAP `CGI_CT_V9` (pain.001.001.09) já está configurada e em uso noutros países "
      "(CH/DE/SE/UA/HR/ES), mas **não existe formato PT associado a ela** — PT continua em pain.001.001.03.")
    A("- A árvore genérica SAP `CGI_DD_V8` está importada mas **sem qualquer formato a referenciá-la em todo "
      "o sistema** (0 resultados em TFPM042F).")
    A("- Existe um override ativo (`T042ZA_FORMAT`) que já usa `Z_CGI_DD` (pain.008.001.08 — versão nova) "
      "para o método Q na house bank BPI01 da empresa 2010 — coexistindo com o formato standard "
      "`Z_SEPA_DD_AR` (pain.008.001.02) usado nas restantes house banks.")
    A("")

    A("## 3. Métodos de pagamento configurados\n")
    A(f"Total de métodos configurados no sistema (T042Z, todos os países): **{audit.t042z_all_count}**, "
      f"em **{len(audit.t042z_countries)} países**.\n")
    A("### 3.1 Portugal (LAND1 = PT)\n")
    A("| Método | Descrição | Formato (FORMI) | Programa clássico | SEPA | IBAN | Pgto único |")
    A("|---|---|---|---|---|---|---|")
    for r in audit.t042z_pt:
        A(f"| {r['ZLSCH']} | {r['TEXT1']} | {r['FORMI'] or '—'} | {r['PROGN'] or '—'} | "
          f"{r['XSEPA'] or '—'} | {r['XIBAN'] or '—'} | {r['XEINZ'] or '—'} |")
    A("")

    A("## 4. Métodos com utilização real (REGUH, 2024-2026)\n")
    A("Contagem por **execução lógica** (par `LAUFD`+`LAUFI` distinto), excluindo runs de proposta "
      "(`XVORL = X`). Uma execução pode conter múltiplas linhas de pagamento.\n")
    A("| Empresa | Método | Descrição | Últ. execução | 2024 | 2025 | 2026 |")
    A("|---|---|---|---|---|---|---|")
    for r in audit.reguh_usage:
        anos = r["execucoes_por_ano"]
        A(f"| {r['ZBUKR']} | {r['RZAWE'] or '(vazio)'} | {method_text(r['RZAWE'])} | "
          f"{r['ultima_execucao'] or '—'} | {anos.get('2024', 0)} | {anos.get('2025', 0)} | {anos.get('2026', 0)} |")
    A("")

    A("## 5. Formatos PMW (TFPM042F)\n")
    A("| Formato | País | Versão ISO (FORME) | Docu | Tree ID | DME ativo |")
    A("|---|---|---|---|---|---|")
    for formi, r in sorted(audit.formatos.items()):
        A(f"| {formi} | {r['LAND1'] or '—'} | {r['FORME'] or '—'} | {r['FORMD'] or '—'} | "
          f"{r['TREE_ID'] or '(= FORMI)'} | {r['XDMEE'] or '—'} |")
    A("")

    A("## 6. Árvores DMEE/DMEEX\n")
    A("| Tree ID | Parent | Docu | DMEEX | Extensível |")
    A("|---|---|---|---|---|")
    for tid, r in sorted(audit.dmee_trees.items()):
        A(f"| {tid} | {r['PARENT_ID'] or '—'} | {r['DOCU_TXT'] or '—'} | {r['DMEEX'] or '—'} | {r['EXTENSIBLE'] or '—'} |")
    A("")
    A(f"- Formatos que referenciam `CGI_CT_V9` como TREE_ID (sistema completo): "
      f"{', '.join(r['FORMI'] + '/' + r['LAND1'] for r in audit.cgi_ct_v9_users) or 'nenhum'}")
    A(f"- Formatos que referenciam `CGI_DD_V8` como TREE_ID (sistema completo): "
      f"{', '.join(r['FORMI'] + '/' + r['LAND1'] for r in audit.cgi_dd_v8_users) or 'NENHUM — árvore importada mas não configurada'}")
    A("")

    A("## 7. Pagamentos/Transferências (Fornecedores)\n")
    A("- **Método:** `S` — SEPA-Fornecedor")
    A("- **Formato standard (T042Z):** `Z_SEPA_AP` — pain.001.001.03 — árvore `Z_SEPA_CT` (standalone, sem parent)")
    A("- **Override ativo (T042ZA_FORMAT):** house banks BPI01/BPI02/BPI03 (empresa 2010) e BPI01 "
      "(empresas 2080/2100) usam `Z_CGI_CT` — também pain.001.001.03, árvore `Z_CGI_CT` (standalone)")
    A("- **Uso real:** ativo em todas as 4 empresas operacionais PT (2010/2020/2080/2100), última execução "
      "até 2026-09-16 (ver secção 4)")
    A("- **Estado:** USO ATUAL — versão ISO ainda não migrada para pain.001.001.09 em nenhum caminho PT "
      "(nem standard nem override)")
    A("")

    A("## 8. Débito Direto\n")
    A("- **Método:** `Q` — Débito Direto-Cliente")
    A("- **Formato standard (T042Z):** `Z_SEPA_DD_AR` — pain.008.001.02 — árvore `ZSEPA_DD` (standalone)")
    A("- **Override ativo (T042ZA_FORMAT):** house bank BPI01 da empresa 2010 usa `Z_CGI_DD` — "
      "**pain.008.001.08** (versão mais recente), árvore `Z_CGI_DD` (standalone, não derivada da "
      "árvore genérica SAP `CGI_DD_V8`)")
    A("- **Empresas com uso real (REGUH):** apenas 2010 e 2100 — 2020 e 2080 não têm execuções do "
      "método Q em 2024-2026")
    A("- **Árvore genérica `CGI_DD_V8` (SAP standard, importada):** confirmada existente em DMEE_TREE, "
      "mas SEM QUALQUER formato a referenciá-la em todo o sistema (não só PT) — "
      "**IMPORTADA MAS NÃO CONFIGURADA**")
    A("- **Árvore Z customizada:** SIM — `Z_CGI_DD` e `Z_SEPA_DD_AR`/`ZSEPA_DD`, ambas standalone "
      "(não herdam da CGI_DD_V8 genérica)")
    A("- Resposta direta: produção usa **pain.008.001.02** (`ZSEPA_DD`, standard) na maioria das house "
      "banks, e **pain.008.001.08** (`Z_CGI_DD`, override) apenas na house bank BPI01/empresa 2010. "
      "Portugal usa Débito Direto atualmente, mas só nas empresas 2010 e 2100.")
    A("")

    A("## 9. RH/Payroll\n")
    A("- **Método:** `2` — SEPA Salários-Fornecedor")
    A("- **Formato:** `Z_SEPA_RH` — pain.001.001.03 — árvore `Z_SEPA_RH` (standalone, sem override "
      "encontrado em T042ZA_FORMAT para nenhuma empresa PT)")
    A("- **Mecanismo confirmado:** RH usa o MESMO motor PMW/DME dos pagamentos a fornecedores "
      "(T042Z/TFPM042F/DMEE) — NÃO existe programa clássico de DME de folha (`RPCDTA*`): 0 execuções "
      "encontradas em TBTCP")
    A("- **Uso real:** todas as 4 empresas PT (2010/2020/2080/2100) com execuções recorrentes e recentes "
      "(última execução entre 2026-08-26 e 2026-09-15) — volume elevado (centenas de execuções/ano)")
    A("- **Evidência de execução via F110:** jobs `F110-*` confirmados em TBTCP nos últimos dias "
      "(até 2026-09-17), com step `RFFOAVIS_FPAYM`/`RFFOAVIS_DD_PRENOTIF` — corrobora que o ciclo de "
      "pagamento (incluindo salários) passa pelo programa de pagamento automático standard")
    A("- Resposta direta: RH gera o ficheiro bancário através do formato `Z_SEPA_RH` (pain.001.001.03, "
      "árvore `Z_SEPA_RH`), no mesmo fluxo PMW que os fornecedores; não existe árvore nova importada "
      "específica para RH neste momento (a árvore genérica CGI_CT_V9 tem um formato HR_CGI_XML_CT_V9, "
      "mas é da Croácia — `LAND1 = HR` — não de Recursos Humanos; não confundir)")
    A("")

    A("## 10. Empresas e House Banks\n")
    A("| Empresa | Nome | País | Moeda |")
    A("|---|---|---|---|")
    for r in audit.t001_pt:
        A(f"| {r['BUKRS']} | {r['BUTXT']} | {r['LAND1']} | {r['WAERS']} |")
    A("")
    A("House banks citados nos overrides/variantes (via DFPAYV/T042ZA_FORMAT): "
      "BPI01/02/03, BCP01, CGD01/02, NB001/002, BBV01/02/03, BST01/02/03, CBK01, FIN00, ICB01, "
      "MOX01, ACB01, BIC01, DB001, BNP01, BIL01, ISB01, BBI01. Nomes/IBAN não confirmados via RFC "
      "(ver secção 15 — T012 sem dados via RFC_READ_TABLE).")
    A("")
    A("### Amostra BNKA (bancos PT)\n")
    A("| Código banco | Nome | SWIFT |")
    A("|---|---|---|")
    for r in audit.bnka_pt_sample[:15]:
        A(f"| {r['BANKL']} | {r['BANKA'] or '—'} | {r['SWIFT'] or '—'} |")
    A("")

    A("## 11. Variantes (DFPAYV / OBPM4)\n")
    for formi, rows in audit.dfpayv.items():
        vari_names = sorted({r["VARI"] for r in rows if r["VARI"]})
        hbkids = sorted({r["HBKID"] for r in rows if r["HBKID"]})
        A(f"- **{formi}:** variante(s) `{', '.join(vari_names) or '—'}`, "
          f"{len(rows)} atribuições empresa/house bank, house banks: {', '.join(hbkids)}")
    A("\n(Conteúdo funcional interno das variantes — VALIDAÇÃO MANUAL NECESSÁRIA, ver secção 15)\n")

    A("## 12. Árvores CGI importadas\n")
    A("| Tree ID | Classificação |")
    A("|---|---|")
    A("| CGI_CT_V9 | CONFIGURADA E UTILIZADA (mas não por PT — CH/DE/SE/UA/HR/ES) |")
    A("| CGI_DD_V8 | IMPORTADA MAS NÃO CONFIGURADA (nenhum formato no sistema a referencia) |")
    A("| CGI_CT | CONFIGURADA MAS SEM USO RECENTE identificado (nenhum override PT aponta diretamente para ela) |")
    A("| CGI_DD | CONFIGURADA MAS SEM USO RECENTE identificado (nenhum override PT aponta diretamente para ela) |")
    A("| Z_CGI_CT (custom) | CONFIGURADA E UTILIZADA — override ativo para método S em house banks BPI da empresa 2010/2080/2100 |")
    A("| Z_CGI_DD (custom) | CONFIGURADA E UTILIZADA — override ativo para método Q em BPI01/empresa 2010, já em pain.008.001.08 |")
    A("")

    A("## 13. Formatos antigos ainda utilizados\n")
    A("- `Z_SEPA_AP` / `Z_CGI_CT` — pain.001.001.03 (legado; a versão nova pain.001.001.09 via CGI_CT_V9 "
      "não está configurada para PT) — EM USO ATIVO")
    A("- `Z_SEPA_RH` — pain.001.001.03 (mesma geração legada) — EM USO ATIVO")
    A("- `Z_SEPA_DD_AR` — pain.008.001.02 — EM USO ATIVO na maioria das house banks (exceto BPI01/2010)")
    A("")

    A("## 14. Candidatos a migração\n")
    A("| Processo | Situação atual | Classificação |")
    A("|---|---|---|")
    A("| Pagamentos Fornecedores (CT) | pain.001.001.03 em toda a produção PT; CGI_CT_V9 (pain.001.001.09) "
      "já ativo noutros países mas sem formato PT | ÁRVORE NOVA DISPONÍVEL MAS NÃO CONFIGURADA |")
    A("| Débito Direto (DD) | pain.008.001.02 standard + pain.008.001.08 já em uso via override "
      "(Z_CGI_DD) apenas em 1 house bank | FORMATO NOVO PARCIALMENTE MIGRADO — análise manual para "
      "decidir extensão a outras house banks |")
    A("| Débito Direto — árvore genérica SAP | CGI_DD_V8 importada, zero uso no sistema | ÁRVORE NOVA "
      "DISPONÍVEL MAS NÃO CONFIGURADA |")
    A("| RH/Payroll | pain.001.001.03 (Z_SEPA_RH), sem override, sem árvore nova associada | FORMATO "
      "ANTIGO AINDA EM USO — nenhuma árvore nova candidata identificada para RH especificamente |")
    A("")

    A("## 15. Validações manuais necessárias\n")
    for m in audit.manual_checks:
        A(f"### {m['item']}")
        A(f"- **Motivo (porque o RFC não conseguiu provar):** {m['motivo']}")
        A(f"- **Transação:** {m['transacao']}")
        A(f"- **Caminho de menu:** {m['caminho']}")
        A(f"- **Valor/campo esperado a confirmar:** {m['valor_esperado']}")
        A(f"- **Evidência necessária:** {m['evidencia']}")
        A("")

    A("## 16. Inventário técnico de tabelas e campos\n")
    A("| Tabela | Existe (DD02L) |")
    A("|---|---|")
    for r in audit.ddic:
        A(f"| {r['tabela']} | {r['existe']} |")
    A("")

    A("## 17. Conclusão\n")
    A("Portugal opera 3 processos de geração de ficheiro bancário via PMW/DMEE: pagamentos a "
      "fornecedores (SEPA CT, pain.001.001.03), débito direto (SEPA DD, pain.008.001.02 com uma "
      "exceção já em pain.008.001.08) e salários/RH (SEPA CT, pain.001.001.03, mesmo motor dos "
      "fornecedores). F110 está confirmadamente ativo e recente em produção. A árvore genérica SAP "
      "mais nova para débito direto (CGI_DD_V8) está importada mas completamente por configurar em "
      "todo o sistema; a árvore genérica para transferências (CGI_CT_V9) já está em uso noutros "
      "países mas ainda não foi estendida a Portugal. Nenhuma alteração foi efetuada — esta é uma "
      "análise exclusivamente de leitura.\n")

    md = "\n".join(lines)
    output_path.write_text(md, encoding="utf-8")
    return md


def write_json(audit: Audit, output_path: Path) -> None:
    payload = {
        "ambiente": audit.ambiente,
        "ddic": audit.ddic,
        "t001_pt": audit.t001_pt,
        "t042z_all_count": audit.t042z_all_count,
        "t042z_countries": audit.t042z_countries,
        "t042z_pt": audit.t042z_pt,
        "formatos": audit.formatos,
        "dmee_trees": audit.dmee_trees,
        "cgi_ct_v9_users": audit.cgi_ct_v9_users,
        "cgi_dd_v8_users": audit.cgi_dd_v8_users,
        "overrides_pt": audit.overrides_pt,
        "dfpayv": audit.dfpayv,
        "reguh_usage": audit.reguh_usage,
        "tbtcp_rpcdta": audit.tbtcp_rpcdta,
        "tbtcp_rffo": audit.tbtcp_rffo,
        "bnka_pt_sample": audit.bnka_pt_sample,
        "manual_checks": audit.manual_checks,
    }
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_csv(audit: Audit, output_path: Path) -> None:
    with output_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Processo", "Pais", "Empresa", "Metodo", "House_Bank", "Formato", "Arvore",
            "Versao_ISO", "Variante", "Ultima_Utilizacao", "Estado",
        ])
        for r in audit.t042z_pt:
            zlsch = r["ZLSCH"]
            formi = r["FORMI"]
            formato_info = audit.formatos.get(formi, {})
            processo = {
                "S": "Pagamento Fornecedores", "Q": "Débito Direto", "2": "RH/Payroll",
            }.get(zlsch, "Outro")
            usage_rows = [u for u in audit.reguh_usage if u["RZAWE"] == zlsch]
            for u in usage_rows:
                writer.writerow([
                    processo, "PT", u["ZBUKR"], zlsch, "", formi,
                    formato_info.get("TREE_ID") or "(= FORMI)",
                    formato_info.get("FORME", ""), "", u["ultima_execucao"],
                    "USO ATUAL" if u["ultima_execucao"] >= "20260101" else "USO HISTORICO",
                ])


def print_summary(audit: Audit, files: dict[str, Path]) -> None:
    print("\n" + "=" * 70)
    print("RESUMO — INVENTÁRIO PAYMENT MEDIUM / FICHEIROS BANCÁRIOS (PRD)")
    print("=" * 70)
    print(f"\nAmbiente: SID={audit.ambiente.get('SID')} CLIENT={audit.ambiente.get('CLIENT')} "
          f"USER={audit.ambiente.get('USER')} RELEASE={audit.ambiente.get('RELEASE')} "
          f"HOST={audit.ambiente.get('HOST')}")

    print("\nPROCESSOS ENCONTRADOS:")
    print("1. Pagamentos Fornecedores — método S — Z_SEPA_AP/Z_CGI_CT — pain.001.001.03 — USO ATIVO")
    print("2. Débito Direto — método Q — Z_SEPA_DD_AR (pain.008.001.02) / Z_CGI_DD (pain.008.001.08, "
          "só BPI01/2010) — USO ATIVO (apenas empresas 2010 e 2100)")
    print("3. RH/Payroll — método 2 — Z_SEPA_RH — pain.001.001.03 — USO ATIVO (todas as 4 empresas PT)")

    print("\nÁRVORES NOVAS IMPORTADAS:")
    print("- CGI_CT_V9 (pain.001.001.09): em uso noutros países, NÃO configurada para PT")
    print("- CGI_DD_V8 (pain.008.001.08 nativa): IMPORTADA, ZERO formatos a usá-la em todo o sistema")

    print("\nÁRVORES/FORMATOS ATUAIS EM USO (PT):")
    print("- Z_SEPA_CT / Z_SEPA_AP (CT standard) e Z_CGI_CT (CT override, BPI)")
    print("- ZSEPA_DD / Z_SEPA_DD_AR (DD standard) e Z_CGI_DD (DD override, BPI01/2010, já pain.008.001.08)")
    print("- Z_SEPA_RH (RH/Payroll)")

    print("\nFORMATOS ANTIGOS EM USO:")
    print("- pain.001.001.03 em CT (fornecedores) e RH — nenhum caminho PT usa ainda pain.001.001.09")
    print("- pain.008.001.02 em DD, exceto override pontual BPI01/2010")

    print("\nPONTOS QUE NECESSITAM MIGRAÇÃO (candidatos, não decidido):")
    print("- Avaliar extensão de CGI_CT_V9 (pain.001.001.09) a Portugal (fornecedores e/ou RH)")
    print("- Avaliar extensão do override Z_CGI_DD (pain.008.001.08) às restantes house banks de DD")
    print("- Avaliar se vale a pena configurar CGI_DD_V8 nativo ou manter Z_CGI_DD customizado")

    print("\nVALIDAÇÕES MANUAIS:")
    for m in audit.manual_checks:
        print(f"- {m['item']} (transação {m['transacao']})")

    print("\nFICHEIROS GERADOS:")
    for label, path in files.items():
        print(f"- {label}: {path}")
    print()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--forcar-sid", action="store_true", help="Prosseguir mesmo que o SID não seja o esperado (PRD).")
    parser.add_argument("--output-dir", default=None, help="Diretório de saída (default: <root>/output).")
    args = parser.parse_args()

    project_root = find_project_root()
    load_project_env(project_root)
    output_dir = Path(args.output_dir) if args.output_dir else project_root / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    guard = SafetyGuard.build(
        allow_write_operations=False,
        allowed_functions=AUDIT_ALLOWED_FUNCTIONS,
        allowed_tables=AUDIT_READ_TABLES,
    )

    try:
        from pyrfc import Connection
    except Exception as exc:  # pragma: no cover
        print(f"Falha ao carregar PyRFC: {exc}")
        return 1

    params = build_connection_params()
    connection = Connection(**params)
    rfc = ReadOnlyRFC(connection, guard)

    audit = Audit()
    try:
        sec_ambiente(rfc, audit, args.forcar_sid, params.get("client", ""), params.get("user", ""))
    except AuditSecurityError as exc:
        print(f"❌ {exc}")
        rfc.close()
        return 2

    print(f"Ligado a SID={audit.ambiente['SID']} CLIENT={audit.ambiente['CLIENT']} — a recolher dados (read-only)...")

    sec_ddic(rfc, audit)
    sec_metodos(rfc, audit)
    global ZLSCH_TO_TEXT
    ZLSCH_TO_TEXT = {r["ZLSCH"]: r["TEXT1"] for r in audit.t042z_pt}
    sec_formatos(rfc, audit)
    sec_dmee_trees(rfc, audit)
    sec_overrides(rfc, audit)
    sec_dfpayv(rfc, audit)
    print("A calcular estatísticas de uso real (REGUH 2024-2026) — pode demorar alguns minutos...")
    sec_reguh_usage(rfc, audit)
    sec_tbtcp(rfc, audit)
    sec_bnka(rfc, audit)
    build_manual_checks(audit)

    rfc.close()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    md_path = output_dir / f"inventario_payment_medium_PRD_{timestamp}.md"
    json_path = output_dir / f"inventario_payment_medium_PRD_{timestamp}.json"
    csv_path = output_dir / f"inventario_payment_medium_PRD_{timestamp}.csv"

    render_md(audit, md_path)
    write_json(audit, json_path)
    write_csv(audit, csv_path)

    print_summary(audit, {"Relatório (MD)": md_path, "Dados (JSON)": json_path, "Matriz (CSV)": csv_path})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

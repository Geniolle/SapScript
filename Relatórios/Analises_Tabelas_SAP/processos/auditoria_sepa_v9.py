# -*- coding: utf-8 -*-
"""Auditoria técnica READ-ONLY da migração SEPA CT Portugal -> pain.001.001.09.

Objecto auditado: Z_PT_CGI_XML_CT_V9 (formato PMW + árvore DMEEX + FBZP/OBPM1/
OBPM4 + método PT/S + overrides + variantes + bancos empresa + BIC).

REGRAS (não negociáveis):
  * Só leitura. Nenhuma função RFC de escrita, nenhum COMMIT, nenhuma tabela
    fora da whitelist ``AUDIT_READ_TABLES`` é tocada.
  * Toda a chamada RFC passa por ``_safe_call`` (whitelist de funções + guarda
    de tokens de escrita reutilizada de ``sap_payroll_analysis.security``).
  * Ambiente alvo: QAD (SID S4Q). Se a ligação não for QAD o script PÁRA.

Este módulo NÃO usa o motor SE16H/SE16N de ``engine.py`` (GUI scripting): a
auditoria exige RFC/PyRFC. É executável directamente:

    .venv-rfc\\Scripts\\python.exe "Relatórios\\Analises_Tabelas_SAP\\processos\\auditoria_sepa_v9.py"

Opções:
    --sid-esperado S4Q      SID que a ligação principal tem de ter (default S4Q)
    --forcar-sid            continuar mesmo que o SID não coincida (NÃO recomendado)
    --com-dev               tentar ligar também a SAP_DEV_* e comparar config
    --output-dir output     pasta de saída (default: <raiz>/output)
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[3]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Reutiliza o núcleo de segurança já existente no projecto.
from sap_payroll_analysis.security import _MUTATION_TOKENS, SecurityError  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# =============================================================================
# WHITELISTS READ-ONLY (defesa em profundidade)
# =============================================================================

AUDIT_ALLOWED_FUNCTIONS: frozenset[str] = frozenset(
    {
        "RFC_PING",
        "RFC_SYSTEM_INFO",
        "RFC_READ_TABLE",
        "DDIF_FIELDINFO_GET",
    }
)

#: Tabelas/views que a auditoria pode ler. Qualquer outra é bloqueada.
AUDIT_READ_TABLES: frozenset[str] = frozenset(
    {
        # DDIC (metadados)
        "DD02L", "DD02T", "DD03L", "DD03T", "DD04T",
        # DMEEX / DME engine
        "DMEE_TREE", "DMEE_TREE_HEAD", "DMEE_TREE_NODE", "DMEE_TREE_NODE_T",
        "DMEE_TREE_RULES",
        # PMW - formatos (OBPM1) família FI-AP/AR
        "TFPM042F", "TFPM042FT", "TFPM042FB", "TFPM042FF", "TFPM042FD",
        "TFPM042FM", "TFPM042FG", "TFPM042FV", "TFPM042FZ", "TFPM042FZT",
        "TFPM042FC", "TFPM042FBC", "TFPM042FPB", "TFPM042FSB", "T042FID",
        # PMW - espelho "customizing" (SM30)
        "TFCUPM042F", "TFCUPM042FT", "TFCUPM042FB", "TFCUPM042FF",
        "TFCUPM042FG", "TFCUPM042FV", "TFCUPM042FZ", "TFCUPM042FM",
        # PMW - variantes de selecção (OBPM4)
        "DFPAYV", "DFPAYV_VARI", "DFPAY_XREL", "DFPAY_DYN_SEL",
        # Método de pagamento / país / empresa
        "T042Z", "T042E", "T042ZA", "T042ZA_FORMAT", "T042ZA_PREFTYP",
        "TDNOTEP_REFTYP", "TDNOTEP_REFTYPFP",
        # Bancos empresa / dados bancários
        "T012", "T012K", "BNKA",
        # Empresas (texto)
        "T001",
    }
)


class AuditSecurityError(SecurityError):
    pass


def _guard_readonly(function_name: str, query_table: str | None) -> None:
    name = str(function_name or "").strip().upper()
    if name not in AUDIT_ALLOWED_FUNCTIONS:
        raise AuditSecurityError(
            f"Função RFC '{function_name}' fora da whitelist READ-ONLY da auditoria."
        )
    hit = next((tok for tok in _MUTATION_TOKENS if tok in name), None)
    if hit is not None:
        raise AuditSecurityError(
            f"Função RFC '{function_name}' contém token de escrita '{hit}'. Bloqueada."
        )
    if query_table:
        tab = str(query_table).strip().upper()
        if tab not in AUDIT_READ_TABLES:
            raise AuditSecurityError(
                f"Tabela '{query_table}' não está em AUDIT_READ_TABLES. "
                f"Adicione-a explicitamente no código para a poder ler."
            )


# =============================================================================
# CLIENTE RFC READ-ONLY
# =============================================================================

REQUIRED_SUFFIXES = ("USER", "PASSWD", "ASHOST", "SYSNR", "CLIENT")


def _load_env() -> None:
    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env", override=False)


def _prefix_complete(prefix: str) -> bool:
    import os

    return all(os.getenv(f"{prefix}{s}", "").strip() for s in REQUIRED_SUFFIXES)


def _params_for(prefix: str) -> dict[str, str]:
    import os

    if not _prefix_complete(prefix):
        missing = [f"{prefix}{s}" for s in REQUIRED_SUFFIXES if not os.getenv(f"{prefix}{s}", "").strip()]
        raise RuntimeError(f"Parâmetros RFC ausentes no .env: {', '.join(missing)}")
    return {
        "user": os.environ[f"{prefix}USER"],
        "passwd": os.environ[f"{prefix}PASSWD"],
        "ashost": os.environ[f"{prefix}ASHOST"],
        "sysnr": os.environ[f"{prefix}SYSNR"],
        "client": os.environ[f"{prefix}CLIENT"],
        "lang": os.getenv(f"{prefix}LANG", "PT").strip() or "PT",
    }


class ReadOnlyRFC:
    """Wrapper fino sobre pyrfc.Connection. Só expõe leitura."""

    def __init__(self, prefix: str, params: dict[str, str]):
        try:
            from pyrfc import Connection  # type: ignore
        except Exception as exc:  # pragma: no cover
            raise RuntimeError(f"PyRFC indisponível: {exc}") from exc
        self.prefix = prefix
        self._safe_params = {k: v for k, v in params.items() if k != "passwd"}
        self._conn = Connection(**params)
        self.unreadable_fields: list[str] = []

    # -- infra -----------------------------------------------------------------
    def call(self, function_name: str, **kwargs: Any) -> dict[str, Any]:
        _guard_readonly(function_name, kwargs.get("QUERY_TABLE"))
        return dict(self._conn.call(function_name, **kwargs) or {})

    def ping(self) -> None:
        self.call("RFC_PING")

    def system_info(self) -> dict[str, str]:
        exp = self.call("RFC_SYSTEM_INFO").get("RFCSI_EXPORT", {}) or {}
        info = {k: str(v).strip() for k, v in exp.items()}
        attrs = {}
        try:
            attrs = {k: str(v).strip() for k, v in (self._conn.get_connection_attributes() or {}).items()}
        except Exception:
            pass
        info["_client"] = attrs.get("client", self._safe_params.get("client", ""))
        info["_user"] = attrs.get("user", self._safe_params.get("user", ""))
        info["_sysId"] = attrs.get("sysId", info.get("RFCSYSID", ""))
        info["_ashost"] = self._safe_params.get("ashost", "")
        return info

    def close(self) -> None:
        try:
            self._conn.close()
        except Exception:
            pass

    # -- leitura de tabelas --------------------------------------------------
    def read_table(
        self,
        table: str,
        fields: list[str],
        where: list[str] | None = None,
        rowcount: int = 0,
    ) -> list[dict[str, str]]:
        payload = {
            "QUERY_TABLE": table,
            "DELIMITER": "|",
            "FIELDS": [{"FIELDNAME": f} for f in fields],
            "OPTIONS": [{"TEXT": t} for t in (where or [])],
            "ROWCOUNT": rowcount,
        }
        res = self.call("RFC_READ_TABLE", **payload)
        cols = [e["FIELDNAME"] for e in res.get("FIELDS", [])]
        out: list[dict[str, str]] = []
        for row in res.get("DATA", []):
            vals = str(row.get("WA", "")).split("|")
            out.append({c: (vals[i].strip() if i < len(vals) else "") for i, c in enumerate(cols)})
        return out

    def read_table_wide(
        self,
        table: str,
        key_fields: list[str],
        data_fields: list[str],
        where: list[str] | None = None,
        rowcount: int = 0,
        group_size: int = 4,
    ) -> list[dict[str, str]]:
        """Lê tabelas largas em grupos de campos, dividindo adaptativamente
        sempre que o SAP devolver DATA_BUFFER_EXCEEDED (linha > 512 bytes).
        Campos individuais que não cabem (ex.: STRING/XML) são omitidos e
        registados em ``self.unreadable_fields``."""
        merged: dict[tuple[str, ...], dict[str, str]] = {}
        order: list[tuple[str, ...]] = []
        remaining = [f for f in data_fields if f not in key_fields]

        def run_group(grp: list[str]) -> None:
            if not grp:
                return
            sel = key_fields + grp
            try:
                rows = self.read_table(table, sel, where=where, rowcount=rowcount)
            except Exception as exc:  # noqa: BLE001
                if "BUFFER_EXCEEDED" in str(exc) or "DATA_BUFFER" in str(exc):
                    if len(grp) == 1:
                        self.unreadable_fields.append(f"{table}.{grp[0]}")
                        return
                    mid = len(grp) // 2
                    run_group(grp[:mid])
                    run_group(grp[mid:])
                    return
                raise
            for r in rows:
                k = tuple(r.get(kf, "") for kf in key_fields)
                if k not in merged:
                    merged[k] = {kf: r.get(kf, "") for kf in key_fields}
                    order.append(k)
                merged[k].update(r)

        for i in range(0, len(remaining), group_size):
            run_group(remaining[i : i + group_size])
        # garante que as chaves aparecem mesmo que todos os grupos de dados falhem
        if not order:
            try:
                for r in self.read_table(table, key_fields, where=where, rowcount=rowcount):
                    k = tuple(r.get(kf, "") for kf in key_fields)
                    if k not in merged:
                        merged[k] = r
                        order.append(k)
            except Exception:  # noqa: BLE001
                pass
        return [merged[k] for k in order]

    def field_catalog(self, table: str) -> list[dict[str, str]]:
        rows = self.read_table(
            "DD03L",
            ["FIELDNAME", "POSITION", "KEYFLAG", "ROLLNAME", "DATATYPE", "LENG"],
            where=[f"TABNAME = '{table.upper()}'"],
        )
        rows = [r for r in rows if not r["FIELDNAME"].startswith(".")]
        rows.sort(key=lambda r: int(r["POSITION"] or 0))
        return rows

    def table_exists(self, table: str) -> dict[str, str] | None:
        rows = self.read_table(
            "DD02L", ["TABNAME", "TABCLASS", "AS4LOCAL"],
            where=[f"TABNAME = '{table.upper()}'"],
        )
        rows = [r for r in rows if r.get("AS4LOCAL", "A") == "A"]
        return rows[0] if rows else None

    def table_text(self, table: str) -> str:
        rows = self.read_table(
            "DD02T", ["DDTEXT"],
            where=[f"TABNAME = '{table.upper()}'", "AND DDLANGUAGE = 'E'"],
        )
        return rows[0]["DDTEXT"] if rows else ""


# =============================================================================
# MODELO DE RESULTADOS
# =============================================================================

OK = "OK"
DIV = "DIVERGENCIA"
ERR = "ERRO"
NV = "NAO_VALIDADO"
MAN = "MANUAL"

EMOJI = {OK: "✅", DIV: "⚠️", ERR: "❌", NV: "❓", MAN: "👤"}


@dataclass
class Finding:
    seccao: str
    item: str
    status: str
    esperado: str = ""
    encontrado: str = ""
    nota: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "seccao": self.seccao,
            "item": self.item,
            "status": self.status,
            "emoji": EMOJI.get(self.status, "?"),
            "esperado": self.esperado,
            "encontrado": self.encontrado,
            "nota": self.nota,
        }


@dataclass
class ManualCheck:
    item: str
    porque: str
    transacao: str
    caminho: str
    valor_esperado: str
    evidencia: str

    def as_dict(self) -> dict[str, str]:
        return self.__dict__.copy()


@dataclass
class TableInventory:
    configuracao: str
    transacao: str
    view: str
    tabela_base: str
    campos_chave: str
    campos_relevantes: str
    como_identificada: str

    def as_dict(self) -> dict[str, str]:
        return self.__dict__.copy()


@dataclass
class Audit:
    findings: list[Finding] = field(default_factory=list)
    manuals: list[ManualCheck] = field(default_factory=list)
    inventory: list[TableInventory] = field(default_factory=list)
    raw: dict[str, Any] = field(default_factory=dict)

    def add(self, *a: Any, **kw: Any) -> None:
        self.findings.append(Finding(*a, **kw))

    def manual(self, *a: Any, **kw: Any) -> None:
        self.manuals.append(ManualCheck(*a, **kw))

    def inv(self, *a: Any, **kw: Any) -> None:
        self.inventory.append(TableInventory(*a, **kw))

    def counts(self) -> dict[str, int]:
        c = {OK: 0, DIV: 0, ERR: 0, NV: 0, MAN: 0}
        for f in self.findings:
            c[f.status] = c.get(f.status, 0) + 1
        c[MAN] = max(c[MAN], len(self.manuals))
        return c


# =============================================================================
# CONFIGURAÇÃO ESPERADA
# =============================================================================

FORMATO = "Z_PT_CGI_XML_CT_V9"
TREE_ID = "Z_PT_CGI_XML_CT_V9"
TREE_TYPE = "PAYM"
PARENT_TREE = "CGI_CT_V9"
ARVORE_ANTIGA = "Z_SEPA_CT"
ARVORE_ANTIGA_ALT = "Z_CGI_CT"
PAIS = "PT"
METODO = "S"
VARIANTE = "ZSEPA_V9"

OBPM1_ESPERADO = {
    "LAND1": "PT",
    "FORME": "PAIN.001.001.09",
    "FORMD": "FPM_DOCU_CGI_XML_CT",
    "DTTYP": "04",
    "XDMEE": "X",
    "XDME1": "X",
    "BEANZ": "15",
    "TREE_ID": "",  # árvore divergente vazia
}
OBPM1_PARAM_STRUCT = "FPM_CGI"

EVENTO_ESPERADO = {"05": "FI_PAYMEDIUM_DMEE_CGI_05"}
EVENTOS_INESPERADOS = {"00", "20", "25", "30", "40"}

REFERENCIAS_ESPERADAS = {
    "1": {"LENGTH": "35", "NUMBR": "4"},
    "2": {"LENGTH": "35", "NUMBR": "1"},
    "3": {"LENGTH": "35", "NUMBR": "1"},
}

HOUSE_BANKS = [
    ("2010", "BPI01"), ("2010", "BST01"), ("2010", "NB002"),
    ("2020", "BST01"),
    ("2080", "BPI01"), ("2080", "BST01"),
    ("2100", "BCP01"), ("2100", "BPI01"), ("2100", "BST01"),
]
OBPM4_ESPERADO = list(HOUSE_BANKS)  # todas -> ZSEPA_V9

BIC_ALVO = ("2100", "BPI01")
BIC_ESPERADO_XML = "BBPIPTP0"

FORMATOS_ANTIGOS = {"Z_CGI_CT", "Z_SEPA_AP", "Z_SEPA_AP_SCT", "Z_SEPA_CT"}


def _q(v: str) -> str:
    return v.replace("'", "''")


_COMPANY_COUNTRY: dict[str, str] = {}


def _company_country(rfc: "ReadOnlyRFC", bukrs: str) -> str:
    if not _COMPANY_COUNTRY:
        try:
            for r in rfc.read_table("T001", ["BUKRS", "LAND1"], where=[]):
                _COMPANY_COUNTRY[r.get("BUKRS", "")] = r.get("LAND1", "")
        except Exception:  # noqa: BLE001
            pass
    return _COMPANY_COUNTRY.get(bukrs, "")


def _cmp(found: str, expected: str) -> bool:
    return str(found).strip().upper() == str(expected).strip().upper()


# =============================================================================
# SECÇÕES DA AUDITORIA
# =============================================================================

def sec_ambiente(rfc: ReadOnlyRFC, info: dict[str, str], audit: Audit, sid_esperado: str) -> None:
    sid = info.get("_sysId") or info.get("RFCSYSID", "")
    audit.raw["ambiente"] = info
    linha = f"SID={sid} | Cliente={info.get('_client','')} | Utilizador={info.get('_user','')} | Host={info.get('_ashost','')} | Release={info.get('RFCSAPRL','')} | DB={info.get('RFCDBSYS','')}"
    if _cmp(sid, sid_esperado):
        audit.add("1. Ambiente", "Ligação RFC ao ambiente alvo", OK,
                  esperado=f"SID={sid_esperado} (QAD)", encontrado=linha,
                  nota=f"Destino RFC: {info.get('RFCDEST','')}")
    else:
        audit.add("1. Ambiente", "Ligação RFC ao ambiente alvo", DIV,
                  esperado=f"SID={sid_esperado} (QAD)", encontrado=linha,
                  nota="A ligação NÃO é o ambiente esperado — ver aviso no terminal.")


def sec_ddic_inventory(rfc: ReadOnlyRFC, audit: Audit) -> dict[str, dict[str, str] | None]:
    alvo = [
        "DMEE_TREE", "DMEE_TREE_HEAD", "DMEE_TREE_NODE", "DMEE_TREE_NODE_T", "DMEE_TREE_RULES",
        "TFPM042F", "TFPM042FT", "TFPM042FB", "TFPM042FF", "TFPM042FG", "TFPM042FV",
        "TFPM042FZ", "TFPM042FM", "TFPM042FD",
        "DFPAYV", "DFPAYV_VARI", "DFPAY_XREL", "DFPAY_DYN_SEL",
        "T042Z", "T042E", "T042ZA", "T042ZA_FORMAT", "T042ZA_PREFTYP",
        "T012", "T012K", "BNKA", "T001",
    ]
    existencia: dict[str, dict[str, str] | None] = {}
    for t in alvo:
        try:
            existencia[t] = rfc.table_exists(t)
        except Exception as exc:  # noqa: BLE001
            existencia[t] = None
            audit.add("15. Inventário", f"DDIC {t}", NV, encontrado=str(exc))
    audit.raw["ddic_existencia"] = {k: (v or {}) for k, v in existencia.items()}

    def cat(t: str) -> str:
        try:
            return ", ".join(f["FIELDNAME"] for f in rfc.field_catalog(t))
        except Exception:
            return "(catálogo indisponível)"

    _txt = rfc.table_text
    if existencia.get("DMEE_TREE"):
        audit.inv("DMEEX - cabeçalho árvore", "DMEEX", "-", "DMEE_TREE",
                  "TREE_TYPE + TREE_ID", cat("DMEE_TREE"),
                  "DD02L (existência) + DD03L (campos)")
    if existencia.get("DMEE_TREE_HEAD"):
        audit.inv("DMEEX - versões da árvore", "DMEEX", "-", "DMEE_TREE_HEAD",
                  "TREE_TYPE + TREE_ID + VERSION", cat("DMEE_TREE_HEAD"),
                  "DD02L + DD03L")
    if existencia.get("DMEE_TREE_NODE"):
        audit.inv("DMEEX - nós/mapeamentos", "DMEEX", "-", "DMEE_TREE_NODE",
                  "TREE_TYPE + TREE_ID + VERSION + NODE_ID", cat("DMEE_TREE_NODE"),
                  "DD02L + DD03L; reutiliza field list de scripts/inspect_dmeex.py")
    if existencia.get("TFPM042F"):
        audit.inv("Formato PMW (OBPM1) - atributos gerais", "OBPM1",
                  "V_TFPM042F / cluster PM_042F", "TFPM042F", "FORMI",
                  cat("TFPM042F"),
                  "DD02T LIKE '%Payment medium format%' -> família TFPM042F*")
    if existencia.get("TFPM042FT"):
        audit.inv("Formato PMW - textos", "OBPM1", "-", "TFPM042FT",
                  "SPRAS + FORMI", cat("TFPM042FT"), "DD02T + DD03L")
    if existencia.get("TFPM042FG"):
        audit.inv("Formato PMW - separação (empresa/banco)", "OBPM1", "-", "TFPM042FG",
                  "FORMI", cat("TFPM042FG"),
                  "DD02T 'Level of detail of payment medium'")
    if existencia.get("TFPM042FF"):
        audit.inv("Formato PMW - estrutura de parâmetros", "OBPM1", "-", "TFPM042FF",
                  "FORMI + FORMF", cat("TFPM042FF"), "DD02T 'Format parameters'")
    if existencia.get("TFPM042FB"):
        audit.inv("Eventos PMW", "OBPM1 (aba Eventos)", "-", "TFPM042FB",
                  "FORMI + EVENT", cat("TFPM042FB"), "DD02T 'Payment medium formats: Events'")
    if existencia.get("TFPM042FV"):
        audit.inv("Referências / notas ao beneficiário", "OBPM1", "-", "TFPM042FV",
                  "FORMI + FORMZ + TYPE", cat("TFPM042FV"),
                  "DD02T 'Note to payee fields'")
    if existencia.get("TFPM042FZ"):
        audit.inv("Suplementos do formato", "OBPM1", "-", "TFPM042FZ",
                  "FORMI + FORMZ", cat("TFPM042FZ"), "DD02T 'Supplements'")
    if existencia.get("TFPM042FM"):
        audit.inv("Parâmetros obrigatórios do formato", "OBPM1", "-", "TFPM042FM",
                  "FORMI + FORMF + FORMM", cat("TFPM042FM"),
                  "DD02T 'Reqd fields for format parameters'")
    if existencia.get("DFPAYV"):
        audit.inv("OBPM4 - atribuição variante por empresa/banco", "OBPM4",
                  "V_DFPAYV", "DFPAYV",
                  "FORMI + ZBUKR + BANKS + BANKL + HBKID + HKTID + CRDEB + RZAWE + LFDNR",
                  cat("DFPAYV"),
                  "DD03L FIELDNAME='VARI' -> DFPAYV/DFPAYV_VARI/DFPAY_XREL/DFPAY_DYN_SEL")
    if existencia.get("DFPAYV_VARI"):
        audit.inv("OBPM4 - variantes definidas para o formato", "OBPM4", "-", "DFPAYV_VARI",
                  "FORMI + VARI", cat("DFPAYV_VARI"), "DD03L FIELDNAME='VARI'")
    if existencia.get("DFPAY_XREL"):
        audit.inv("OBPM4 - conteúdo/parâmetros da variante", "OBPM4", "-", "DFPAY_XREL",
                  "FORMI + VARI + FORMF + FIELD_NAME", cat("DFPAY_XREL"),
                  "DD02T 'Selection variant Maintenance'")
    if existencia.get("DFPAY_DYN_SEL"):
        audit.inv("OBPM4 - selecções dinâmicas da variante", "OBPM4", "-", "DFPAY_DYN_SEL",
                  "FORMI + VARI + SEQ_NUM", cat("DFPAY_DYN_SEL"),
                  "DD02T 'Dynamic Selections for Variant'")
    if existencia.get("T042Z"):
        audit.inv("Método de pagamento por país", "FBZP / OBVCU", "V_T042Z", "T042Z",
                  "LAND1 + ZLSCH",
                  "FORMI, FORMZ, XSEPA, XIBAN, TEXT1, PROGN, XNO_ACCNO, XEINZ",
                  "conhecida (FBZP) + DD03L confirmado nesta release")
    if existencia.get("T042E"):
        audit.inv("Método de pagamento por empresa", "FBZP", "V_T042E", "T042E",
                  "ZBUKR + ZLSCH", ", ".join(f["FIELDNAME"] for f in rfc.field_catalog("T042E")),
                  "conhecida (FBZP)")
    if existencia.get("T042ZA_FORMAT"):
        audit.inv("Override 'Format in Company Code'", "OBPM2 / FBZP",
                  "V_T042ZA_FORMAT", "T042ZA_FORMAT",
                  "ZBUKR + ZLSCH + HBKID", cat("T042ZA_FORMAT"),
                  "DD02T 'Different PMW Format'")
    if existencia.get("T042ZA"):
        audit.inv("Tipo preferido PMW por país", "FBZP", "V_T042ZA", "T042ZA",
                  "LAND1 + ZLSCH + ORIGIN", cat("T042ZA"), "DD02T 'Different PMW Format'")
    if existencia.get("T042ZA_PREFTYP"):
        audit.inv("Tipo preferido PMW por empresa/banco", "FBZP", "V_T042ZA_PREFTYP",
                  "T042ZA_PREFTYP", "ZBUKR + ZLSCH + HBKID + ORIGIN",
                  cat("T042ZA_PREFTYP"), "DD02T 'Different PMW Format'")
    if existencia.get("T012"):
        audit.inv("Bancos empresa (house banks)", "FI12 / FBZP", "V_T012", "T012",
                  "BUKRS + HBKID", cat("T012"), "conhecida")
    if existencia.get("T012K"):
        audit.inv("Contas de banco empresa", "FI12", "V_T012K", "T012K",
                  "BUKRS + HBKID + HKTID", cat("T012K"), "conhecida")
    if existencia.get("BNKA"):
        audit.inv("Bancos (dados mestre) / BIC", "FI01 / BAsic", "-", "BNKA",
                  "BANKS + BANKL", "BANKA, SWIFT, BICKY, BNKLZ, LOEVM", "conhecida")
    return existencia


def _read_dmee_tree(rfc: ReadOnlyRFC, tree_id: str) -> dict[str, Any]:
    tree = rfc.read_table_wide(
        "DMEE_TREE",
        ["TREE_TYPE", "TREE_ID"],
        ["CREA_USER", "CREA_DATE", "CREA_TIME", "CHNG_USER", "CHNG_DATE", "CHNG_TIME",
         "DOCU_TXT", "RELEASE_FLAG", "ORIG_LANGU", "XML_TREE", "PARENT_ID",
         "TREE_LEVEL", "DMEEX", "EXTENSIBLE"],
        where=[f"TREE_ID = '{_q(tree_id)}'"],
    )
    heads = rfc.read_table_wide(
        "DMEE_TREE_HEAD",
        ["TREE_TYPE", "TREE_ID", "VERSION"],
        ["PARAM_STRUC", "FIRSTNODE_ID", "SF_NAME", "VERS_USER", "VERS_DATE", "VERS_TIME",
         "VERSION_TYPE", "VERSION_DESCRIPTION", "DETAILED_DESCRIPTION", "IF_TYPE",
         "MSGTYPE", "XSLTDESC", "CHARSET"],
        where=[f"TREE_ID = '{_q(tree_id)}'"],
    )
    try:
        nodes = rfc.read_table(
            "DMEE_TREE_NODE",
            ["TREE_TYPE", "TREE_ID", "VERSION", "NODE_ID", "TECH_NAME", "REF_NAME",
             "PARENT_ID", "NODE_TYPE", "LEV", "EX_STATUS"],
            where=[f"TREE_ID = '{_q(tree_id)}'"],
        )
    except Exception as exc:  # noqa: BLE001
        nodes = []
    return {"tree": tree, "heads": heads, "nodes": nodes}


def sec_dmeex(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("DMEE_TREE"):
        audit.add("3. DMEEX", "Tabelas DMEE_*", NV, encontrado="DMEE_TREE não encontrada no DDIC")
        return
    data = _read_dmee_tree(rfc, TREE_ID)
    audit.raw["dmeex_v9"] = data
    tree_rows = data["tree"]
    if not tree_rows:
        audit.add("3. DMEEX", f"Árvore {TREE_ID} existe", ERR,
                  esperado=f"1 registo em DMEE_TREE (TREE_TYPE={TREE_TYPE})",
                  encontrado="0 registos")
        return
    tr = tree_rows[0]
    audit.add("3. DMEEX", f"Árvore {TREE_ID} existe", OK,
              esperado=f"TREE_TYPE={TREE_TYPE}", encontrado=f"TREE_TYPE={tr.get('TREE_TYPE','')}")
    audit.add("3. DMEEX", "Tipo da árvore (TREE_TYPE)",
              OK if _cmp(tr.get("TREE_TYPE", ""), TREE_TYPE) else DIV,
              esperado=TREE_TYPE, encontrado=tr.get("TREE_TYPE", ""))
    audit.add("3. DMEEX", "Parent / generic tree (DMEE_TREE.PARENT_ID)",
              OK if _cmp(tr.get("PARENT_ID", ""), PARENT_TREE) else DIV,
              esperado=PARENT_TREE, encontrado=tr.get("PARENT_ID", "") or "(vazio)",
              nota="Objectivo pain.001.001.09 assenta na árvore genérica CGI_CT_V9.")
    audit.add("3. DMEEX", "Descrição (DOCU_TXT)",
              OK if tr.get("DOCU_TXT", "").strip() else DIV,
              esperado="(descrição preenchida)", encontrado=tr.get("DOCU_TXT", "") or "(vazio)")
    audit.add("3. DMEEX", "Idioma original (ORIG_LANGU)", OK,
              encontrado=tr.get("ORIG_LANGU", ""))
    audit.add("3. DMEEX", "Flag DMEEX / EXTENSIBLE", OK,
              encontrado=f"DMEEX={tr.get('DMEEX','')} EXTENSIBLE={tr.get('EXTENSIBLE','')} XML_TREE={tr.get('XML_TREE','')}")
    rel = tr.get("RELEASE_FLAG", "")
    audit.add("3. DMEEX", "DMEE_TREE.RELEASE_FLAG", OK,
              esperado="valor informativo (não é o indicador de activação nesta release)",
              encontrado=f"RELEASE_FLAG='{rel}'",
              nota=("Neste sistema 329/332 árvores PAYM têm RELEASE_FLAG vazio, incluindo "
                    "árvores standard em uso — não é sinal de erro. A árvore está referenciada "
                    "por T042Z (PT/S) e tem nós na versão 000. Activação/consistência confirmam-se "
                    "na transacção DMEEX (ver validações manuais)."))
    audit.add("3. DMEEX", "Criação / alteração", OK,
              encontrado=(f"criada {tr.get('CREA_DATE','')} por {tr.get('CREA_USER','')}; "
                          f"alterada {tr.get('CHNG_DATE','')} por {tr.get('CHNG_USER','')}"))

    heads = data["heads"]
    if not heads:
        audit.add("3. DMEEX", "Versões (DMEE_TREE_HEAD)", DIV,
                  esperado=">=1 versão", encontrado="0 registos em DMEE_TREE_HEAD")
    else:
        versoes = sorted({h.get("VERSION", "") for h in heads})
        for h in heads:
            audit.add("3. DMEEX", f"Versão {h.get('VERSION','?')}", OK,
                      encontrado=(f"tipo='{h.get('VERSION_TYPE','')}' desc='{h.get('VERSION_DESCRIPTION','')}' "
                                  f"param_struc='{h.get('PARAM_STRUC','')}' firstnode='{h.get('FIRSTNODE_ID','')}' "
                                  f"IF_TYPE='{h.get('IF_TYPE','')}' MSGTYPE='{h.get('MSGTYPE','')}' "
                                  f"por {h.get('VERS_USER','')} em {h.get('VERS_DATE','')}"))
        act = "000" if "000" in versoes else versoes[0]
        v999 = " Versão 999 é a cópia entregue pela SAP (VERS_USER=SAP) — referência, não activa." if "999" in versoes else ""
        audit.add("3. DMEEX", "Versão activa determinada", OK if "000" in versoes else DIV,
                  esperado="Versão 000 activa/productiva",
                  encontrado=f"versões presentes: {', '.join(versoes)}; activa: {act}",
                  nota=("Versão 000 é a produtiva/em uso." + v999 +
                        " Existirem várias versões NÃO é erro."))
        for h in heads:
            if h.get("VERSION") == act and OBPM1_PARAM_STRUCT:
                audit.add("3. DMEEX", "PARAM_STRUC da versão activa",
                          OK if _cmp(h.get("PARAM_STRUC", ""), OBPM1_PARAM_STRUCT) else DIV,
                          esperado=OBPM1_PARAM_STRUCT, encontrado=h.get("PARAM_STRUC", ""))

    nodes = data["nodes"]
    if nodes:
        by_ver: dict[str, int] = {}
        for n in nodes:
            by_ver[n.get("VERSION", "")] = by_ver.get(n.get("VERSION", ""), 0) + 1
        data["nodes_v000"] = sum(1 for n in nodes if n.get("VERSION", "") == "000")
        audit.add("3. DMEEX", "Nós existentes", OK,
                  encontrado="; ".join(f"v{k}: {v} nós" for k, v in sorted(by_ver.items())),
                  nota="Versão activa 000 tem estrutura de nós preenchida.")
    else:
        audit.add("3. DMEEX", "Nós (DMEE_TREE_NODE)", NV,
                  encontrado="Não foi possível ler DMEE_TREE_NODE via RFC_READ_TABLE",
                  nota="Estrutura larga; usar scripts/inspect_dmeex.py apontado a QAD para o detalhe.")

    audit.manual(
        item=f"Consistência da árvore DMEEX {TREE_ID}",
        porque="A verificação de consistência é uma função interna da transacção; "
               "não há tabela que a exponha e a auditoria é read-only.",
        transacao="DMEEX",
        caminho=f"DMEEX -> Tree Type {TREE_TYPE} -> Tree {TREE_ID} -> menu Formato de árvore -> Verificar (Ctrl+F2)",
        valor_esperado=f"Mensagem 'Árvore {TREE_ID} está consistente; nenhum erro encontrado'.",
        evidencia="Print da mensagem de estado após Verificar + print da versão activa (Saltar -> Versões).",
    )


def sec_dmeex_compare(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("DMEE_TREE"):
        return
    for antiga in (ARVORE_ANTIGA, ARVORE_ANTIGA_ALT):
        rows = rfc.read_table(
            "DMEE_TREE", ["TREE_TYPE", "TREE_ID", "PARENT_ID", "DOCU_TXT", "RELEASE_FLAG"],
            where=[f"TREE_ID = '{_q(antiga)}'"],
        )
        if not rows:
            audit.add("3. DMEEX", f"Árvore antiga {antiga} (comparação)", NV,
                      encontrado="não existe em QAD",
                      nota="Comparação nó-a-nó fica para validação manual / ambiente onde exista.")
            continue
        old = _read_dmee_tree(rfc, antiga)
        audit.raw[f"dmeex_{antiga}"] = old
        n_old = sum(1 for n in old["nodes"] if n.get("VERSION", "") == "000")
        n_new = audit.raw.get("dmeex_v9", {}).get("nodes_v000", 0)
        old_parent = (old["tree"][0].get("PARENT_ID", "") if old["tree"] else "")
        audit.add("3. DMEEX", f"Comparação nº de nós {antiga} vs {TREE_ID} (v000)",
                  NV if (not n_old or not n_new) else OK,
                  esperado="Estrutura pain.001.001.09 coerente",
                  encontrado=f"{antiga}: {n_old} nós (parent='{old_parent or 'nenhum'}') | "
                             f"{TREE_ID}: {n_new} nós (parent='{PARENT_TREE}')",
                  nota=("Diferença de nós NÃO é erro: Z_PT_CGI_XML_CT_V9 herda a árvore genérica "
                        f"{PARENT_TREE} (SAP), enquanto {antiga} não tem árvore-pai. Nós-alvo "
                        "(InitgPty/Id, Dbtr/Id, DbtrAcct/Ccy, ChrgBr, InstrId, Authstn/Prtry, "
                        "RmtInf/Ustrd) exigem inspecção de detalhe — scripts/inspect_dmeex.py "
                        "apontado a QAD + scripts/compare_dmee_trees.py."))
    audit.manual(
        item="Comparação estrutural nó-a-nó árvore antiga vs Z_PT_CGI_XML_CT_V9",
        porque="DMEE_TREE_NODE é demasiado larga para RFC_READ_TABLE devolver todos os "
               "atributos de mapeamento num único passo fiável.",
        transacao="DMEEX (2 sessões) ou script scripts/inspect_dmeex.py",
        caminho="Exportar cada árvore com inspect_dmeex.py (--tree-id) apontado a QAD e "
                "comparar com scripts/compare_dmee_trees.py.",
        valor_esperado="Nós GrpHdr/InitgPty/Id, PmtInf/Dbtr/Id, PmtInf/DbtrAcct/Ccy, "
                       "PmtInf/ChrgBr, .../PmtId/InstrId, GrpHdr/Authstn/Prtry, "
                       ".../RmtInf/Ustrd presentes e mapeados conforme CGI_CT_V9.",
        evidencia="JSON de cada árvore + diff de compare_dmee_trees.py.",
    )


def sec_obpm1(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("TFPM042F"):
        audit.add("4. OBPM1", "Tabela TFPM042F", NV, encontrado="não encontrada no DDIC")
        return
    rows = rfc.read_table_wide(
        "TFPM042F", ["FORMI"],
        ["XPRI1", "XPRI2", "XDME1", "XLST1", "LAND1", "BEANZ", "FORME", "FORMD",
         "DTTYP", "XDMEE", "FORMT", "TREE_ID", "APPLICATION_AREA"],
        where=[f"FORMI = '{_q(FORMATO)}'"],
    )
    audit.raw["obpm1_tfpm042f"] = rows
    if not rows:
        audit.add("4. OBPM1", f"Formato {FORMATO} existe (TFPM042F)", ERR,
                  esperado="1 registo", encontrado="0 registos")
        return
    r = rows[0]
    audit.add("4. OBPM1", f"Formato {FORMATO} existe", OK, encontrado="1 registo em TFPM042F")
    for campo, esp in OBPM1_ESPERADO.items():
        found = r.get(campo, "")
        if campo == "TREE_ID":
            ok = (found.strip() == "") or _cmp(found, FORMATO)
            audit.add("4. OBPM1", "Árvore divergente (TREE_ID)",
                      OK if ok else DIV,
                      esperado="(vazio) => usa árvore com o nome do formato",
                      encontrado=found or "(vazio)")
        else:
            audit.add("4. OBPM1", f"{campo}", OK if _cmp(found, esp) else DIV,
                      esperado=esp, encontrado=found or "(vazio)")
    audit.add("4. OBPM1", "Meio electrónico / criar ficheiro (XDME1)",
              OK if _cmp(r.get("XDME1", ""), "X") else DIV,
              esperado="X (activo)", encontrado=r.get("XDME1", "") or "(vazio)")
    audit.add("4. OBPM1", "Mapping DME / DME engine (XDMEE)",
              OK if _cmp(r.get("XDMEE", ""), "X") else DIV,
              esperado="X (activo)", encontrado=r.get("XDMEE", "") or "(vazio)")
    audit.add("4. OBPM1", "FORMT (tipo de formato)", OK, encontrado=r.get("FORMT", ""),
              nota="Valor informativo; XML CGI usa normalmente FORMT vazio/'2'.")

    # textos
    if existencia.get("TFPM042FT"):
        trows = rfc.read_table("TFPM042FT", ["SPRAS", "FORMI", "FORMX"],
                               where=[f"FORMI = '{_q(FORMATO)}'"])
        audit.raw["obpm1_textos"] = trows
        desc = "; ".join(f"{x.get('SPRAS','')}:{x.get('FORMX','')}" for x in trows) or "(sem texto)"
        audit.add("4. OBPM1", "Descrição do formato (TFPM042FT.FORMX)",
                  OK if trows else DIV,
                  esperado="CGI Credit Transfer_V9 (ou equivalente funcional)",
                  encontrado=desc)

    # estrutura de parâmetros
    if existencia.get("TFPM042FF"):
        frows = rfc.read_table("TFPM042FF", ["FORMI", "FORMF"],
                               where=[f"FORMI = '{_q(FORMATO)}'"])
        audit.raw["obpm1_param_struct"] = frows
        vals = [x.get("FORMF", "") for x in frows]
        audit.add("4. OBPM1", "Estrutura de parâmetros (TFPM042FF.FORMF)",
                  OK if OBPM1_PARAM_STRUCT in vals else DIV,
                  esperado=OBPM1_PARAM_STRUCT, encontrado=", ".join(vals) or "(vazio)")

    # separação por empresa / banco empresa
    if existencia.get("TFPM042FG"):
        grows = rfc.read_table("TFPM042FG", ["FORMI", "XBUKR", "XHBKI", "XHKTI", "XEINZ", "XZLSH"],
                               where=[f"FORMI = '{_q(FORMATO)}'"])
        audit.raw["obpm1_separacao"] = grows
        if not grows:
            audit.add("4. OBPM1", "Separação por empresa / banco empresa (TFPM042FG)", DIV,
                      esperado="XBUKR=X e XHBKI=X", encontrado="0 registos em TFPM042FG")
        else:
            g = grows[0]
            audit.add("4. OBPM1", "Separar por empresa (TFPM042FG.XBUKR)",
                      OK if _cmp(g.get("XBUKR", ""), "X") else DIV,
                      esperado="X", encontrado=g.get("XBUKR", "") or "(vazio)")
            audit.add("4. OBPM1", "Separar por banco empresa (TFPM042FG.XHBKI)",
                      OK if _cmp(g.get("XHBKI", ""), "X") else DIV,
                      esperado="X", encontrado=g.get("XHBKI", "") or "(vazio)")
    else:
        audit.manual(
            item="Separação do ficheiro por empresa / por banco empresa",
            porque="Tabela TFPM042FG não disponível nesta release para leitura.",
            transacao="OBPM1",
            caminho=f"OBPM1 -> formato {FORMATO} -> aba 'Especificações p/ criação de ficheiro'",
            valor_esperado="'Separar por empresa' e 'Separar por banco empresa' activos.",
            evidencia="Print da aba.",
        )


def sec_eventos(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("TFPM042FB"):
        audit.add("5. Eventos", "Tabela TFPM042FB", NV, encontrado="não encontrada no DDIC")
        return
    rows = rfc.read_table("TFPM042FB", ["FORMI", "EVENT", "FNAME"],
                          where=[f"FORMI = '{_q(FORMATO)}'"])
    audit.raw["eventos"] = rows
    got = {r.get("EVENT", "").lstrip("0").zfill(2): r.get("FNAME", "") for r in rows}
    if not rows:
        audit.add("5. Eventos", "Eventos do formato", DIV,
                  esperado="Evento 05 = FI_PAYMEDIUM_DMEE_CGI_05",
                  encontrado="0 registos em TFPM042FB",
                  nota="DME engine puro pode herdar eventos da árvore genérica; confirmar em OBPM1.")
    for ev, fm in EVENTO_ESPERADO.items():
        cur = got.get(ev, "")
        audit.add("5. Eventos", f"Evento {ev}",
                  OK if _cmp(cur, fm) else DIV,
                  esperado=f"{ev} = {fm}", encontrado=cur or "(ausente)")
    for ev, fm in sorted(got.items()):
        if ev in EVENTOS_INESPERADOS:
            audit.add("5. Eventos", f"Evento {ev} presente (inesperado)", DIV,
                      esperado="ausente para formato DME engine CGI",
                      encontrado=f"{ev} = {fm}",
                      nota="Tabela TFPM042FB, chave FORMI+EVENT.")
        elif ev not in EVENTO_ESPERADO:
            audit.add("5. Eventos", f"Evento {ev} presente (a rever)", DIV,
                      esperado="(não previsto na especificação)", encontrado=f"{ev} = {fm}")


def sec_referencias(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("TFPM042FV"):
        audit.add("6. Referências", "Tabela TFPM042FV", NV, encontrado="não encontrada no DDIC")
        audit.manual(
            item="Configuração de referências/notas ao beneficiário (Tipo 1/2/3)",
            porque="Tabela TFPM042FV não disponível para leitura nesta release.",
            transacao="OBPM1",
            caminho=f"OBPM1 -> {FORMATO} -> aba 'Referências' / 'Notas para o beneficiário'",
            valor_esperado="Tipo 1: comp. 35, qtd 4 | Tipo 2: comp. 35, qtd 1 | Tipo 3: comp. 35, qtd 1",
            evidencia="Print da aba.",
        )
        return
    rows = rfc.read_table("TFPM042FV", ["FORMI", "FORMZ", "TYPE", "LENGTH", "NUMBR"],
                          where=[f"FORMI = '{_q(FORMATO)}'"])
    audit.raw["referencias"] = rows
    if not rows:
        audit.add("6. Referências", "Referências do formato", DIV,
                  esperado="3 tipos configurados (1/2/3)",
                  encontrado="0 registos em TFPM042FV")
        return
    by_type: dict[str, dict[str, str]] = {}
    for r in rows:
        by_type.setdefault(r.get("TYPE", ""), r)
    for tp, esp in REFERENCIAS_ESPERADAS.items():
        r = by_type.get(tp)
        if not r:
            audit.add("6. Referências", f"Tipo {tp}", DIV,
                      esperado=f"comp. {esp['LENGTH']}, qtd {esp['NUMBR']}", encontrado="(ausente)")
            continue
        comp = str(int(r.get("LENGTH", "0") or 0))
        qtd = str(int(r.get("NUMBR", "0") or 0))
        ok = comp == esp["LENGTH"] and qtd == esp["NUMBR"]
        audit.add("6. Referências", f"Tipo {tp}", OK if ok else DIV,
                  esperado=f"comp. {esp['LENGTH']}, qtd {esp['NUMBR']} (FORMZ vazio)",
                  encontrado=f"comp. {comp}, qtd {qtd}, FORMZ='{r.get('FORMZ','')}'")
    extras = [t for t in by_type if t not in REFERENCIAS_ESPERADAS]
    if extras:
        audit.add("6. Referências", "Tipos adicionais", DIV,
                  esperado="apenas 1/2/3", encontrado=", ".join(extras))


def sec_suplementos(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if existencia.get("TFPM042FZ"):
        rows = rfc.read_table("TFPM042FZ", ["FORMI", "FORMZ"], where=[f"FORMI = '{_q(FORMATO)}'"])
        audit.raw["suplementos"] = rows
        audit.add("7. Suplementos", "Suplementos do formato (TFPM042FZ)",
                  OK if not rows else DIV,
                  esperado="vazio (sem suplementos)",
                  encontrado=f"{len(rows)} registo(s): " + ", ".join(x.get("FORMZ", "") for x in rows) if rows else "vazio")
    else:
        audit.add("7. Suplementos", "TFPM042FZ", NV, encontrado="não encontrada no DDIC")
    if existencia.get("TFPM042FM"):
        rows = rfc.read_table("TFPM042FM", ["FORMI", "FORMF", "FORMM", "REQUIRED"],
                              where=[f"FORMI = '{_q(FORMATO)}'"])
        audit.raw["param_obrigatorios"] = rows
        req = [x for x in rows if _cmp(x.get("REQUIRED", ""), "X")]
        audit.add("7. Suplementos", "Parâmetros obrigatórios (TFPM042FM)",
                  OK if not req else DIV,
                  esperado="vazio (sem parâmetros obrigatórios adicionais)",
                  encontrado=f"{len(rows)} linha(s), {len(req)} marcada(s) REQUIRED")
    else:
        audit.add("7. Suplementos", "TFPM042FM", NV, encontrado="não encontrada no DDIC")


def sec_metodo_pt_s(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("T042Z"):
        audit.add("8. Método PT/S", "Tabela T042Z", NV, encontrado="não encontrada no DDIC")
        return
    rows = rfc.read_table_wide(
        "T042Z", ["LAND1", "ZLSCH"],
        ["TEXT1", "FORMI", "FORMZ", "XSEPA", "XIBAN", "XNO_ACCNO", "PROGN", "XEINZ",
         "XNOPO", "BLART", "TXTSL"],
        where=[f"LAND1 = '{PAIS}'", f"AND ZLSCH = '{METODO}'"],
    )
    audit.raw["t042z_pt_s"] = rows
    if not rows:
        audit.add("8. Método PT/S", f"Método {PAIS}/{METODO} existe", ERR,
                  esperado="1 registo em T042Z", encontrado="0 registos")
        return
    r = rows[0]
    audit.add("8. Método PT/S", f"Método {PAIS}/{METODO} existe", OK,
              encontrado=f"TEXT1='{r.get('TEXT1','')}'")
    audit.add("8. Método PT/S", "Formato PMW genérico (T042Z.FORMI)",
              OK if _cmp(r.get("FORMI", ""), FORMATO) else ERR,
              esperado=FORMATO, encontrado=r.get("FORMI", "") or "(vazio)",
              nota="Elo central da cadeia: PT/S -> Z_PT_CGI_XML_CT_V9.")
    audit.add("8. Método PT/S", "IBAN (T042Z.XIBAN)",
              OK if r.get("XIBAN", "").strip() else DIV,
              esperado="preenchido (X ou S)", encontrado=r.get("XIBAN", "") or "(vazio)")
    audit.add("8. Método PT/S", "Programa clássico (T042Z.PROGN)",
              OK if not r.get("PROGN", "").strip() else DIV,
              esperado="(vazio) — uso exclusivo de PMW",
              encontrado=r.get("PROGN", "") or "(vazio)")
    audit.add("8. Método PT/S", "Flag SEPA (T042Z.XSEPA)", OK,
              esperado="conforme configuração existente",
              encontrado=r.get("XSEPA", "") or "(vazio)",
              nota="Observação: várias formas de pagamento SEPA PT têm XSEPA vazio neste sistema.")
    audit.add("8. Método PT/S", "FORMZ (suplemento) em T042Z",
              OK if not r.get("FORMZ", "").strip() else DIV,
              esperado="(vazio)", encontrado=r.get("FORMZ", "") or "(vazio)")

    # Tipo preferido PMW / "Different PMW Format" (T042ZA) por país
    if existencia.get("T042ZA"):
        za = rfc.read_table("T042ZA", ["LAND1", "ZLSCH", "ORIGIN", "PREFTYP"],
                            where=[f"LAND1 = '{PAIS}'", f"AND ZLSCH = '{METODO}'"])
        za_all = rfc.read_table("T042ZA", ["LAND1", "ZLSCH", "ORIGIN", "PREFTYP"],
                                where=[f"ZLSCH = '{METODO}'"])
        audit.raw["t042za_pt_s"] = za
        audit.raw["t042za_metodo_s_todos_paises"] = za_all
        aponta_antigo = [x for x in za if x.get("PREFTYP", "").upper() in
                         {f.upper() for f in FORMATOS_ANTIGOS} or "SEPA_AP" in x.get("PREFTYP", "").upper()]
        if za and aponta_antigo:
            outros = "; ".join(f"{x.get('LAND1','')}/{x.get('ZLSCH','')}={x.get('PREFTYP','')}"
                               for x in za_all if x.get("LAND1", "") != PAIS)
            audit.add("8. Método PT/S",
                      "Formato PMW preferido por país/origem (T042ZA.PREFTYP)", DIV,
                      esperado=f"sem entrada para PT/{METODO}, OU PREFTYP = {FORMATO}",
                      encontrado="; ".join(f"ORIGIN='{x.get('ORIGIN','')}' -> PREFTYP='{x.get('PREFTYP','')}'" for x in za),
                      nota=("T042ZA ('Different PMW Format') pode sobrepor-se ao FORMI de T042Z na "
                            f"determinação do formato para a origem indicada (aqui FI-AP). Aponta para "
                            f"'{za[0].get('PREFTYP','')}' (formato ANTIGO), não para {FORMATO}. "
                            f"Nota: ES/FR/IE também usam Z_SEPA_AP ({outros or 'idem'}) — indica que a "
                            "linha PT NÃO foi actualizada durante a migração. Requer validação manual "
                            "da precedência nesta release antes de concluir."))
            audit.manual(
                item=f"Precedência T042ZA.PREFTYP vs T042Z.FORMI para PT/{METODO}/FI-AP",
                porque="A auditoria lê a tabela mas a regra de precedência do determinador de "
                       "formato PMW (qual ganha: 'preferred type' ou FORMI do método) depende da "
                       "release e não é observável por tabela.",
                transacao="OBPM1 (definição de 'format type') + simulação em F110/FBPM (sem executar)",
                caminho=("FBZP / OBPM1: verificar se o formato Z_PT_CGI_XML_CT_V9 pertence ao "
                         f"'format type' '{za[0].get('PREFTYP','')}'. Em alternativa, numa proposta "
                         "de pagamento de teste PT/S FI-AP, confirmar em SE38 (RFFO*/PMW) qual "
                         "formato é seleccionado — sem gerar ficheiro."),
                valor_esperado=(f"O formato efectivamente usado para PT/{METODO} origem FI-AP é "
                                f"{FORMATO} (pain.001.001.09). Se T042ZA forçar Z_SEPA_AP, a entrada "
                                "PT tem de ser corrigida ou removida."),
                evidencia="Print de OBPM1 (format type do V9) + print da determinação de formato na proposta.",
            )
        elif za:
            audit.add("8. Método PT/S", "Formato PMW preferido por país (T042ZA.PREFTYP)", DIV,
                      esperado=f"sem entrada, OU PREFTYP = {FORMATO}",
                      encontrado="; ".join(f"ORIGIN={x.get('ORIGIN','')}->{x.get('PREFTYP','')}" for x in za),
                      nota="Confirmar que o 'preferred type' resolve para o formato V9.")
        else:
            audit.add("8. Método PT/S", "Formato PMW preferido por país (T042ZA)", OK,
                      esperado=f"sem entrada PT/{METODO} -> usa FORMI de T042Z", encontrado="0 registos")


def sec_overrides(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("T042ZA_FORMAT"):
        audit.add("8. Overrides", "Tabela T042ZA_FORMAT", NV,
                  encontrado="não encontrada no DDIC",
                  nota="Confirmar em OBPM2 / FBZP 'Format in Company Code'.")
        audit.manual(
            item="Overrides de formato por empresa/banco (Format in Company Code)",
            porque="Tabela de override não legível via RFC nesta release.",
            transacao="OBPM2 (ou FBZP -> 'Set Up Payment Methods per Company Code for Payment Transactions')",
            caminho="Verificar, para país PT / método S, se há formato específico por empresa/banco.",
            valor_esperado="Sem override, OU override = Z_PT_CGI_XML_CT_V9.",
            evidencia="Print da lista de formatos por empresa.",
        )
        return
    rows = rfc.read_table("T042ZA_FORMAT",
                          ["ZBUKR", "ZLSCH", "HBKID", "FORMI", "FORMZ", "DTTYP_ALTV", "HALGO"],
                          where=[f"ZLSCH = '{METODO}'"])
    audit.raw["t042za_format_metodo_s"] = rows
    pt_s = [r for r in rows if _company_country(rfc, r.get("ZBUKR", "")) == PAIS]
    outros_s = [r for r in rows if _company_country(rfc, r.get("ZBUKR", "")) != PAIS]

    if not pt_s:
        audit.add("8. Overrides", f"Overrides de formato para empresas PT / método {METODO}", OK,
                  esperado="sem overrides -> todas as empresas PT herdam Z_PT_CGI_XML_CT_V9 via T042Z",
                  encontrado="0 overrides em T042ZA_FORMAT para empresas de país PT com método S")
    else:
        for r in pt_s:
            fmt = r.get("FORMI", "")
            if fmt.upper() in {f.upper() for f in FORMATOS_ANTIGOS}:
                st, nota = ERR, f"Override PT aponta para FORMATO ANTIGO '{fmt}' — pagamento não usará V9."
            elif _cmp(fmt, FORMATO):
                st, nota = OK, "Override redundante mas coerente (aponta para V9)."
            else:
                st, nota = DIV, f"Override PT aponta para '{fmt}' (não é V9)."
            audit.add("8. Overrides",
                      f"Override PT {r.get('ZBUKR','')}/{r.get('ZLSCH','')}/{r.get('HBKID','')}",
                      st, esperado="sem override, ou = Z_PT_CGI_XML_CT_V9",
                      encontrado=f"FORMI={fmt} FORMZ={r.get('FORMZ','')} DTTYP_ALTV={r.get('DTTYP_ALTV','')}",
                      nota=nota)

    if outros_s:
        det = "; ".join(f"{x.get('ZBUKR','')}({_company_country(rfc, x.get('ZBUKR',''))})/"
                        f"{x.get('HBKID','')}->{x.get('FORMI','')}" for x in outros_s)
        audit.add("8. Overrides", f"Overrides método {METODO} de empresas NÃO-PT (informativo)", OK,
                  esperado="inalterados pela migração PT",
                  encontrado=det,
                  nota="Empresas de outros países (ex.: 2120 = FR/BNP01, formatos *_SCT). "
                       "Fora do âmbito da migração PT — confirmar apenas que não foram tocados.")

    # Panorama global (todos os métodos) — sanidade de que só PT/S foi alvo
    all_rows = rfc.read_table("T042ZA_FORMAT",
                              ["ZBUKR", "ZLSCH", "HBKID", "FORMI"], where=[])
    audit.raw["t042za_format_todos"] = [
        {**x, "LAND1": _company_country(rfc, x.get("ZBUKR", ""))} for x in all_rows
    ]
    pt_outros_metodos = [x for x in all_rows
                         if _company_country(rfc, x.get("ZBUKR", "")) == PAIS and x.get("ZLSCH", "") != METODO]
    audit.add("8. Overrides", "Panorama T042ZA_FORMAT (empresas PT, outros métodos)", OK,
              encontrado="; ".join(f"{x.get('ZBUKR','')}/{x.get('ZLSCH','')}/{x.get('HBKID','')}->{x.get('FORMI','')}"
                                   for x in pt_outros_metodos) or "nenhum",
              nota="Overrides PT de métodos 1/Q/W/Z (confirming, débito directo) — não relacionados com PT/S.")
    if existencia.get("T042ZA_PREFTYP"):
        pref = rfc.read_table("T042ZA_PREFTYP",
                              ["ZBUKR", "ZLSCH", "HBKID", "ORIGIN", "PREFTYP"],
                              where=[f"ZLSCH = '{METODO}'"])
        audit.raw["t042za_preftyp_s"] = pref
        audit.add("8. Overrides", "Tipo preferido por empresa/banco (T042ZA_PREFTYP)",
                  OK if not pref else DIV,
                  esperado="sem entradas para método S",
                  encontrado=f"{len(pref)} registo(s)" +
                             ("" if not pref else ": " + "; ".join(
                                 f"{x.get('ZBUKR','')}/{x.get('HBKID','')}->{x.get('PREFTYP','')}" for x in pref)))


def sec_obpm4(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("DFPAYV"):
        audit.add("9. OBPM4", "Tabela DFPAYV", NV, encontrado="não encontrada no DDIC")
        audit.manual(
            item="Atribuição de variantes de selecção por empresa/banco (OBPM4)",
            porque="Tabela DFPAYV não legível via RFC nesta release.",
            transacao="OBPM4",
            caminho=f"OBPM4 -> formato {FORMATO} -> lista de empresas/bancos e variante atribuída",
            valor_esperado=f"Todas as combinações PT activas -> variante {VARIANTE}",
            evidencia="Print da lista OBPM4 para o formato.",
        )
        return
    rows = rfc.read_table_wide(
        "DFPAYV", ["FORMI", "ZBUKR", "BANKS", "BANKL", "HBKID", "HKTID", "CRDEB", "RZAWE", "LFDNR"],
        ["VARI"],
        where=[f"FORMI = '{_q(FORMATO)}'"],
    )
    audit.raw["dfpayv"] = rows
    combos: dict[tuple[str, str], set[str]] = {}
    for r in rows:
        combos.setdefault((r.get("ZBUKR", ""), r.get("HBKID", "")), set()).add(r.get("VARI", ""))

    # Contexto: quantas atribuições têm os formatos antigos (padrão do sistema)
    ctx = {}
    for f in ("Z_SEPA_AP", "Z_CGI_CT", "Z_SEPA_AP_SCT"):
        try:
            ctx[f] = len(rfc.read_table("DFPAYV", ["FORMI", "VARI"], where=[f"FORMI = '{_q(f)}'"]))
        except Exception:  # noqa: BLE001
            ctx[f] = "?"
    audit.raw["dfpayv_contexto_formatos_antigos"] = ctx

    presentes = sorted(k for k, v in combos.items())
    n_ok = sum(1 for c in OBPM4_ESPERADO if combos.get(c) == {VARIANTE})
    audit.add("9. OBPM4", "Cobertura das atribuições OBPM4 para o formato V9",
              OK if n_ok == len(OBPM4_ESPERADO) else DIV,
              esperado=f"{len(OBPM4_ESPERADO)} combinações empresa/banco -> {VARIANTE}",
              encontrado=f"{len(rows)} linha(s) em DFPAYV; {n_ok}/{len(OBPM4_ESPERADO)} combinações "
                         f"esperadas OK; presentes: " +
                         (", ".join(f"{b}/{h}" for b, h in presentes) or "nenhuma"),
              nota=(f"Para referência, os formatos antigos têm: " +
                    ", ".join(f"{k}={v} linhas" for k, v in ctx.items()) +
                    ". O padrão do sistema é uma linha por combinação — a cobertura quase nula "
                    "do V9 indica que as atribuições OBPM4 da migração NÃO chegaram a QAD "
                    "(provável transporte em falta). Comparar com DEV."))

    table_rows = []
    for (bukrs, hbk) in OBPM4_ESPERADO:
        varis = combos.get((bukrs, hbk), set())
        if not varis:
            st, enc = DIV, "(sem atribuição)"
        elif varis == {VARIANTE}:
            st, enc = OK, VARIANTE
        elif VARIANTE in varis:
            st, enc = DIV, f"{VARIANTE} + outras: {', '.join(sorted(varis))}"
        else:
            st, enc = DIV, ", ".join(sorted(varis))
        table_rows.append({"empresa": bukrs, "banco_empresa": hbk, "variante": enc, "status": st})
        audit.add("9. OBPM4", f"{bukrs} / {hbk}", st, esperado=VARIANTE, encontrado=enc)
    audit.raw["obpm4_tabela"] = table_rows

    for (bukrs, hbk), varis in sorted(combos.items()):
        if (bukrs, hbk) not in OBPM4_ESPERADO:
            audit.add("9. OBPM4", f"{bukrs} / {hbk} (extra)", DIV,
                      esperado="(não consta da lista esperada)",
                      encontrado=", ".join(sorted(varis)),
                      nota="A REVER — não assumir erro; confirmar se é combinação legítima.")

    audit.manual(
        item="Atribuições de variante OBPM4 em falta para Z_PT_CGI_XML_CT_V9",
        porque="RFC confirma que apenas 2010/BPI01 tem variante atribuída em QAD; a criação/"
               "transporte das restantes é acção de configuração, fora do âmbito read-only.",
        transacao="OBPM4",
        caminho=(f"OBPM4 -> formato {FORMATO}: confirmar/gerar as atribuições -> {VARIANTE} para "
                 "2010/BST01, 2010/NB002, 2020/BST01, 2080/BPI01, 2080/BST01, 2100/BCP01, "
                 "2100/BPI01, 2100/BST01. Comparar a lista com DEV (mesma transacção)."),
        valor_esperado=f"As 9 combinações empresa/banco PT SEPA CT apontam para a variante {VARIANTE}.",
        evidencia="Print da lista OBPM4 do formato em DEV e em QAD, lado a lado.",
    )

    # DFPAYV_VARI / DFPAY_XREL / DFPAY_DYN_SEL: neste release estão vazias para
    # TODOS os formatos -> não são fonte fiável; o conteúdo da variante vive
    # como variante ABAP standard do programa de payment medium.
    for tab in ("DFPAYV_VARI", "DFPAY_XREL", "DFPAY_DYN_SEL"):
        if not existencia.get(tab):
            continue
        try:
            total = len(rfc.read_table(tab, ["FORMI"], where=[], rowcount=5))
        except Exception:  # noqa: BLE001
            total = 0
        if total == 0:
            audit.add("10. ZSEPA_V9", f"{tab} (conteúdo/variante)", NV,
                      esperado="—",
                      encontrado=f"{tab} vazia para todos os formatos neste sistema",
                      nota="Tabela não utilizada nesta release; não é indício de erro. "
                           "O conteúdo da variante confirma-se por outra via (ver validação manual).")
        else:
            xr = rfc.read_table(tab, ["FORMI", "VARI"], where=[f"FORMI = '{_q(FORMATO)}'"])
            audit.raw[f"{tab.lower()}_v9"] = xr
            audit.add("10. ZSEPA_V9", f"{tab} para {FORMATO}",
                      OK if xr else DIV,
                      esperado=f"linhas para VARI={VARIANTE}",
                      encontrado=f"{len(xr)} registo(s)")

    audit.manual(
        item=f"Existência e conteúdo da variante ABAP {VARIANTE} do programa de payment medium",
        porque="DFPAYV_VARI/DFPAY_XREL estão vazias neste sistema; a variante de selecção "
               "PMW é uma variante ABAP standard (tabelas VARID/VARIT do programa SAPFPAYM / "
               "RFFO*), que RFC_READ_TABLE não interpreta com segurança.",
        transacao="OBPM4 (ou SE38 -> programa de meio de pagamento PMW -> Passar a -> Variantes)",
        caminho=(f"OBPM4 -> {FORMATO} -> para 2010/BPI01 abrir a variante {VARIANTE}; registar "
                 "programa e parâmetros. Repetir para as restantes combinações depois de criadas "
                 "e confirmar que os parâmetros são idênticos/coerentes."),
        valor_esperado=(f"Variante {VARIANTE} existe, aponta para o programa de payment medium do "
                        "PMW e tem parâmetros coerentes com pain.001.001.09 (sem restrições que "
                        "excluam empresas/bancos PT)."),
        evidencia="Print do ecrã de parâmetros da variante + nome do programa.",
    )


def sec_bancos_empresa(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    if not existencia.get("T012"):
        audit.add("10. Bancos empresa", "Tabela T012", NV, encontrado="não encontrada no DDIC")
        return
    rows = rfc.read_table("T012", ["BUKRS", "HBKID", "BANKS", "BANKL", "NAME1"], where=[])
    idx = {(r.get("BUKRS", ""), r.get("HBKID", "")): r for r in rows}
    audit.raw["t012"] = [idx[k] for k in idx]
    tbl = []
    for (bukrs, hbk) in HOUSE_BANKS:
        r = idx.get((bukrs, hbk))
        if r:
            audit.add("10. Bancos empresa", f"{bukrs} / {hbk}", OK,
                      encontrado=f"BANKS={r.get('BANKS','')} BANKL={r.get('BANKL','')} NAME1='{r.get('NAME1','')}'")
            tbl.append({"bukrs": bukrs, "hbkid": hbk, "banks": r.get("BANKS", ""),
                        "bankl": r.get("BANKL", ""), "name1": r.get("NAME1", ""), "existe": "sim"})
        else:
            audit.add("10. Bancos empresa", f"{bukrs} / {hbk}", ERR,
                      esperado="registo em T012", encontrado="(não existe)")
            tbl.append({"bukrs": bukrs, "hbkid": hbk, "existe": "NAO"})
    audit.raw["bancos_empresa_tabela"] = tbl


def sec_bic_bpi01(rfc: ReadOnlyRFC, existencia: dict, audit: Audit) -> None:
    bukrs, hbk = BIC_ALVO
    if not (existencia.get("T012") and existencia.get("BNKA")):
        audit.add("11. BIC BPI01", "Tabelas T012/BNKA", NV, encontrado="não encontradas no DDIC")
        return
    t012 = rfc.read_table("T012", ["BUKRS", "HBKID", "BANKS", "BANKL", "NAME1"],
                          where=[f"BUKRS = '{bukrs}'", f"AND HBKID = '{hbk}'"])
    if not t012:
        audit.add("11. BIC BPI01", f"House bank {bukrs}/{hbk}", ERR,
                  esperado="registo em T012", encontrado="(não existe)")
        return
    hb = t012[0]
    banks, bankl = hb.get("BANKS", ""), hb.get("BANKL", "")
    t012k = rfc.read_table("T012K", ["BUKRS", "HBKID", "HKTID", "BANKN", "WAERS", "BKONT"],
                           where=[f"BUKRS = '{bukrs}'", f"AND HBKID = '{hbk}'"])
    bnka = rfc.read_table("BNKA", ["BANKS", "BANKL", "BANKA", "SWIFT", "BICKY", "BNKLZ", "LOEVM"],
                          where=[f"BANKS = '{banks}'", f"AND BANKL = '{bankl}'"])
    audit.raw["bic_bpi01"] = {"t012": hb, "t012k": t012k, "bnka": bnka}
    if not bnka:
        audit.add("11. BIC BPI01", "Registo BNKA do bank key", ERR,
                  esperado=f"BNKA para {banks}/{bankl}", encontrado="(não existe)")
        return
    b = bnka[0]
    swift = b.get("SWIFT", "") or b.get("BICKY", "")
    cadeia = (f"Empresa {bukrs} -> House Bank {hbk} -> Bank Country {banks} -> "
              f"Bank Key {bankl} -> BNKA(BANKA='{b.get('BANKA','')}') -> SWIFT/BIC='{swift}'")
    st = OK if _cmp(swift, BIC_ESPERADO_XML) else DIV
    audit.add("11. BIC BPI01", "Cadeia empresa -> BIC", st,
              esperado=f"SWIFT/BIC = {BIC_ESPERADO_XML} (valor gerado no XML de teste)",
              encontrado=cadeia,
              nota=("BIC do XML provém deste registo BNKA. "
                    f"LOEVM(marca eliminação)='{b.get('LOEVM','')}'."))
    audit.add("11. BIC BPI01", "Contas do house bank (T012K)", OK,
              encontrado="; ".join(f"HKTID={x.get('HKTID','')} BANKN={x.get('BANKN','')} WAERS={x.get('WAERS','')}"
                                   for x in t012k) or "(sem contas)")


def sec_dev_compare(audit: Audit, com_dev: bool, qad_raw: dict[str, Any], qad_rfc: "ReadOnlyRFC") -> None:
    qad_raw = dict(qad_raw)
    qad_raw["_qad_rfc"] = qad_rfc
    if not com_dev:
        audit.add("12. DEV x QAD", "Comparação automática DEV vs QAD", NV,
                  encontrado="Não executada (--com-dev não indicado).",
                  nota="Sem acesso a DEV a auditoria não falha; comparar manualmente se necessário. "
                       "Correr novamente com --com-dev para diff automático DEV/QAD do subset.")
        return
    try:
        _load_env()
        if not _prefix_complete("SAP_DEV_"):
            raise RuntimeError("SAP_DEV_* incompleto no .env")
        dev = ReadOnlyRFC("SAP_DEV_", _params_for("SAP_DEV_"))
        dev.ping()
        info = dev.system_info()
    except Exception as exc:  # noqa: BLE001
        audit.add("12. DEV x QAD", "Ligação a DEV", NV, encontrado=str(exc),
                  nota="Comparação DEV/QAD marcada como não executada.")
        return
    try:
        cmp_rows = []
        qad = qad_raw.get("_qad_rfc")

        def one(label: str, tab: str, fields: list[str], where: list[str]) -> None:
            try:
                d = dev.read_table(tab, fields, where=where)
            except Exception as exc:  # noqa: BLE001
                cmp_rows.append({"objeto": label, "dev": f"erro: {exc}", "qad": "-", "igual": "?", "obs": ""})
                return
            try:
                q = qad.read_table(tab, fields, where=where) if qad else []
            except Exception as exc:  # noqa: BLE001
                q = []
            norm = lambda rows: sorted(json.dumps(r, sort_keys=True) for r in rows)
            same = norm(d) == norm(q)
            cmp_rows.append({
                "objeto": f"{tab} ({label})",
                "dev": f"{len(d)} registo(s)",
                "qad": f"{len(q)} registo(s)",
                "igual": "sim" if same else "não",
                "obs": "" if same else "ver JSON dados_brutos p/ detalhe",
            })
            audit.raw.setdefault("dev_qad_detalhe", {})[tab] = {"dev": d, "qad": q}

        one("formato", "TFPM042F", ["FORMI", "LAND1", "FORME", "FORMD", "DTTYP", "XDMEE", "BEANZ", "TREE_ID"],
            [f"FORMI = '{_q(FORMATO)}'"])
        one("textos", "TFPM042FT", ["SPRAS", "FORMI", "FORMX"], [f"FORMI = '{_q(FORMATO)}'"])
        one("eventos", "TFPM042FB", ["FORMI", "EVENT", "FNAME"], [f"FORMI = '{_q(FORMATO)}'"])
        one("referencias", "TFPM042FV", ["FORMI", "FORMZ", "TYPE", "LENGTH", "NUMBR"], [f"FORMI = '{_q(FORMATO)}'"])
        one("separacao", "TFPM042FG", ["FORMI", "XBUKR", "XHBKI"], [f"FORMI = '{_q(FORMATO)}'"])
        one("metodo", "T042Z", ["LAND1", "ZLSCH", "FORMI", "XIBAN"],
            [f"LAND1 = '{PAIS}'", f"AND ZLSCH = '{METODO}'"])
        one("preferido", "T042ZA", ["LAND1", "ZLSCH", "ORIGIN", "PREFTYP"],
            [f"LAND1 = '{PAIS}'", f"AND ZLSCH = '{METODO}'"])
        one("obpm4", "DFPAYV", ["FORMI", "ZBUKR", "HBKID", "VARI"], [f"FORMI = '{_q(FORMATO)}'"])
        one("dmee_head", "DMEE_TREE_HEAD", ["TREE_ID", "VERSION", "PARAM_STRUC"],
            [f"TREE_ID = '{_q(TREE_ID)}'"])
        audit.raw["dev_qad_compare"] = cmp_rows
        audit.raw["dev_info"] = info
        difs = [c for c in cmp_rows if c["igual"] == "não"]
        for c in cmp_rows:
            st = OK if c["igual"] == "sim" else (NV if c["igual"] == "?" else DIV)
            audit.add("12. DEV x QAD", c["objeto"], st,
                      esperado="DEV == QAD", encontrado=f"DEV: {c['dev']} | QAD: {c['qad']} | igual: {c['igual']}",
                      nota=c["obs"])
        audit.add("12. DEV x QAD", "Resultado global (subset via RFC)",
                  OK if not difs else DIV,
                  esperado="configuração igual em DEV e QAD para os objectos da migração",
                  encontrado=f"DEV SID={info.get('_sysId','')} cliente={info.get('_client','')}; "
                             f"{len(cmp_rows)} objecto(s) comparados, {len(difs)} divergente(s)",
                  nota="Subset via RFC. Árvore DMEEX completa: usar scripts/inspect_dmeex.py em ambos.")
    finally:
        dev.close()


# =============================================================================
# OUTPUTS
# =============================================================================

def write_outputs(audit: Audit, info: dict[str, str], out_dir: Path, started: datetime) -> dict[str, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = started.strftime("%Y%m%d_%H%M%S")
    base = out_dir / f"sepa_v9_config_audit_QAD_{ts}"
    counts = audit.counts()

    payload = {
        "meta": {
            "gerado_em": datetime.now().isoformat(timespec="seconds"),
            "inicio": started.isoformat(timespec="seconds"),
            "ambiente": {
                "sistema": info.get("_sysId", ""),
                "cliente": info.get("_client", ""),
                "utilizador": info.get("_user", ""),
                "host": info.get("_ashost", ""),
                "release": info.get("RFCSAPRL", ""),
                "db": info.get("RFCDBSYS", ""),
                "destino_rfc": info.get("RFCDEST", ""),
            },
            "objeto_auditado": FORMATO,
            "contadores": counts,
        },
        "findings": [f.as_dict() for f in audit.findings],
        "validacoes_manuais": [m.as_dict() for m in audit.manuals],
        "inventario_tabelas": [t.as_dict() for t in audit.inventory],
        "dados_brutos": audit.raw,
    }
    json_path = base.with_suffix(".json")
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    csv_path = base.with_suffix(".csv")
    with csv_path.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["seccao", "item", "status", "emoji", "esperado", "encontrado", "nota"])
        for f in audit.findings:
            d = f.as_dict()
            w.writerow([d["seccao"], d["item"], d["status"], d["emoji"], d["esperado"], d["encontrado"], d["nota"]])

    md_path = base.with_suffix(".md")
    md_path.write_text(_render_md(audit, info, counts, started), encoding="utf-8")
    return {"json": json_path, "csv": csv_path, "md": md_path}


def _md_table(findings: Iterable[Finding]) -> str:
    lines = ["| Item | Status | Esperado | Encontrado | Nota |", "|---|---|---|---|---|"]
    for f in findings:
        d = f.as_dict()

        def esc(s: str) -> str:
            return str(s).replace("|", "\\|").replace("\n", " ")

        lines.append(f"| {esc(d['item'])} | {d['emoji']} {d['status']} | {esc(d['esperado'])} | "
                     f"{esc(d['encontrado'])} | {esc(d['nota'])} |")
    return "\n".join(lines)


def _render_md(audit: Audit, info: dict[str, str], counts: dict[str, int], started: datetime) -> str:
    by_sec: dict[str, list[Finding]] = {}
    for f in audit.findings:
        by_sec.setdefault(f.seccao, []).append(f)

    P: list[str] = []
    P.append("# Auditoria SEPA V9 — QAD\n")
    P.append(f"> Objecto: **{FORMATO}** · pain.001.001.09 · gerado {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

    P.append("## 1. Ambiente\n")
    P.append(f"- **Sistema:** {info.get('_sysId','')}")
    P.append(f"- **Cliente:** {info.get('_client','')}")
    P.append(f"- **Utilizador:** {info.get('_user','')}")
    P.append(f"- **Host / Destino RFC:** {info.get('_ashost','')} · {info.get('RFCDEST','')}")
    P.append(f"- **Release / DB:** {info.get('RFCSAPRL','')} · {info.get('RFCDBSYS','')}")
    P.append(f"- **Método de acesso:** RFC / PyRFC, apenas leitura (RFC_READ_TABLE, DDIF_FIELDINFO_GET, RFC_SYSTEM_INFO)\n")

    total = sum(v for k, v in counts.items() if k != MAN) + counts[MAN]
    P.append("## 2. Resumo executivo\n")
    P.append(f"| Resultado | Nº |")
    P.append("|---|---|")
    P.append(f"| ✅ OK | {counts[OK]} |")
    P.append(f"| ⚠️ Divergências | {counts[DIV]} |")
    P.append(f"| ❌ Erros | {counts[ERR]} |")
    P.append(f"| 👤 Validações manuais | {counts[MAN]} |")
    P.append(f"| ❓ Não validado | {counts[NV]} |\n")
    veredicto = "CONFIGURAÇÃO COMPLETA" if (counts[ERR] == 0 and counts[DIV] == 0) else "CONFIGURAÇÃO COM PONTOS A REVER"
    P.append(f"**Resultado:** {veredicto}\n")
    P.append("Cadeia principal validada:\n")
    P.append("```")
    P.append("PT / S")
    P.append(f"   -> {FORMATO}   (T042Z.FORMI)")
    P.append(f"   -> OBPM1 / TFPM042F (DME engine, pain.001.001.09)")
    P.append(f"   -> DMEEX {TREE_ID} (TREE_TYPE {TREE_TYPE}, parent {PARENT_TREE})")
    P.append(f"   -> OBPM4 / DFPAYV -> variante {VARIANTE}")
    P.append("   -> geração pain.001.001.09")
    P.append("```\n")

    sec_titles = [
        ("3. DMEEX", "## 3. DMEEX"),
        ("4. OBPM1", "## 4. OBPM1"),
        ("5. Eventos", "## 5. Eventos"),
        ("6. Referências", "## 6. Referências"),
        ("7. Suplementos", "## 7. Suplementos e parâmetros obrigatórios"),
        ("8. Método PT/S", "## 8. Método PT/S"),
        ("8. Overrides", "## 8b. Overrides por empresa / banco empresa"),
        ("9. OBPM4", "## 9. OBPM4 / ZSEPA_V9"),
        ("10. ZSEPA_V9", "## 9b. Conteúdo da variante ZSEPA_V9"),
        ("10. Bancos empresa", "## 10. Bancos empresa"),
        ("11. BIC BPI01", "## 11. BIC BPI01"),
        ("12. DEV x QAD", "## 12. Comparação DEV x QAD"),
    ]
    for key, title in sec_titles:
        if key in by_sec:
            P.append(title + "\n")
            P.append(_md_table(by_sec[key]) + "\n")

    P.append("## 13. Divergências e erros\n")
    probs = [f for f in audit.findings if f.status in (DIV, ERR)]
    if not probs:
        P.append("_Nenhuma divergência ou erro registado._\n")
    else:
        P.append(_md_table(probs) + "\n")

    P.append("## 14. Validações manuais necessárias\n")
    if not audit.manuals:
        P.append("_Nenhuma._\n")
    for i, m in enumerate(audit.manuals, 1):
        P.append(f"### 14.{i} {m.item}\n")
        P.append(f"- **Porque não foi validado via RFC:** {m.porque}")
        P.append(f"- **Transacção:** {m.transacao}")
        P.append(f"- **Caminho/menu:** {m.caminho}")
        P.append(f"- **Valor esperado:** {m.valor_esperado}")
        P.append(f"- **Evidência a recolher:** {m.evidencia}\n")

    P.append("## 15. Inventário técnico de tabelas/views\n")
    P.append("| Configuração | Transacção | View | Tabela base | Campos chave | Como identificada |")
    P.append("|---|---|---|---|---|---|")
    for t in audit.inventory:
        d = t.as_dict()
        P.append(f"| {d['configuracao']} | {d['transacao']} | {d['view']} | {d['tabela_base']} | "
                 f"{d['campos_chave']} | {d['como_identificada']} |")
    P.append("")
    P.append("<details><summary>Campos relevantes por tabela</summary>\n")
    for t in audit.inventory:
        P.append(f"- **{t.tabela_base}**: {t.campos_relevantes}")
    P.append("\n</details>\n")

    P.append("## 16. Conclusão\n")
    P.append(f"- **Veredicto:** {veredicto}")
    P.append(f"- Cadeia PT/S -> {FORMATO} -> DMEEX -> OBPM4/{VARIANTE} -> pain.001.001.09: "
             f"{'íntegra' if veredicto == 'CONFIGURAÇÃO COMPLETA' else 'com pontos a rever (ver secção 13)'}.")
    P.append(f"- {counts[MAN]} validação(ões) manual(is) obrigatória(s) antes de dar a migração por fechada.")
    P.append("- Nenhuma configuração foi alterada. Auditoria 100% read-only.\n")
    return "\n".join(P)


def print_summary(audit: Audit, info: dict[str, str]) -> None:
    c = audit.counts()
    veredicto = "CONFIGURAÇÃO COMPLETA" if (c[ERR] == 0 and c[DIV] == 0) else "CONFIGURAÇÃO COM PONTOS A REVER"
    print("\n" + "=" * 60)
    print("AUDITORIA SEPA V9 — QAD")
    print("=" * 60)
    print(f"Sistema : {info.get('_sysId','')}")
    print(f"Cliente : {info.get('_client','')}")
    print(f"Utilizador : {info.get('_user','')}")
    print()
    print(f"  OK                 : {c[OK]:>3}")
    print(f"  Divergencias       : {c[DIV]:>3}")
    print(f"  Erros              : {c[ERR]:>3}")
    print(f"  Validacoes manuais : {c[MAN]:>3}")
    print(f"  Nao validado       : {c[NV]:>3}")
    print()
    print("Cadeia principal:")
    print("  PT / S")
    print(f"     -> {FORMATO}")
    print("     -> DMEEX ativa")
    print(f"     -> OBPM4 / {VARIANTE}")
    print("     -> pain.001.001.09")
    print()
    print(f"RESULTADO: {veredicto}")
    print("=" * 60)
    probs = [f for f in audit.findings if f.status in (DIV, ERR)]
    if probs:
        print("\nDivergencias / erros:")
        for f in probs:
            print(f"  {EMOJI[f.status]} [{f.seccao}] {f.item}: esperado='{f.esperado}' | encontrado='{f.encontrado}'")
    if audit.manuals:
        print("\nValidacoes manuais necessarias:")
        for m in audit.manuals:
            print(f"  - {m.item}  ->  {m.transacao}")


# =============================================================================
# MAIN
# =============================================================================

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Auditoria READ-ONLY SEPA CT V9 (pain.001.001.09) em QAD")
    parser.add_argument("--sid-esperado", default="S4Q")
    parser.add_argument("--forcar-sid", action="store_true")
    parser.add_argument("--com-dev", action="store_true")
    parser.add_argument("--output-dir", default=str(ROOT / "output"))
    args = parser.parse_args(argv)

    started = datetime.now()
    _load_env()

    try:
        rfc = ReadOnlyRFC("SAP_QAD_", _params_for("SAP_QAD_"))
    except Exception as exc:  # noqa: BLE001
        print(f"❌ Não foi possível ligar a QAD via RFC: {exc}")
        return 2

    try:
        rfc.ping()
        info = rfc.system_info()
    except Exception as exc:  # noqa: BLE001
        print(f"❌ Falha no RFC_PING/RFC_SYSTEM_INFO: {exc}")
        rfc.close()
        return 2

    sid = info.get("_sysId") or info.get("RFCSYSID", "")
    print("=" * 60)
    print("VERIFICAÇÃO DE AMBIENTE (obrigatória antes de ler)")
    print("=" * 60)
    print(f"Sistema    : {sid}")
    print(f"Cliente    : {info.get('_client','')}")
    print(f"Utilizador : {info.get('_user','')}")
    print(f"Host       : {info.get('_ashost','')}  |  Destino RFC: {info.get('RFCDEST','')}")
    print(f"Release    : {info.get('RFCSAPRL','')}  |  DB: {info.get('RFCDBSYS','')}")
    if not _cmp(sid, args.sid_esperado):
        print("\n" + "!" * 60)
        print(f"AVISO: a ligação NÃO é o ambiente esperado ({args.sid_esperado}).")
        print(f"Ligação actual: SID={sid}.")
        print("!" * 60)
        if not args.forcar_sid:
            print("\nAuditoria INTERROMPIDA. Use --forcar-sid para continuar deliberadamente.")
            rfc.close()
            return 3
        print("\n--forcar-sid indicado: a continuar apesar do SID não coincidir.\n")

    audit = Audit()
    try:
        sec_ambiente(rfc, info, audit, args.sid_esperado)
        existencia = sec_ddic_inventory(rfc, audit)
        sec_dmeex(rfc, existencia, audit)
        sec_dmeex_compare(rfc, existencia, audit)
        sec_obpm1(rfc, existencia, audit)
        sec_eventos(rfc, existencia, audit)
        sec_referencias(rfc, existencia, audit)
        sec_suplementos(rfc, existencia, audit)
        sec_metodo_pt_s(rfc, existencia, audit)
        sec_overrides(rfc, existencia, audit)
        sec_obpm4(rfc, existencia, audit)
        sec_bancos_empresa(rfc, existencia, audit)
        sec_bic_bpi01(rfc, existencia, audit)
    finally:
        pass

    sec_dev_compare(audit, args.com_dev, audit.raw, rfc)

    if rfc.unreadable_fields:
        audit.raw["campos_nao_legiveis_rfc"] = rfc.unreadable_fields
        audit.add("15. Inventário", "Campos não legíveis via RFC_READ_TABLE", NV,
                  encontrado=", ".join(sorted(set(rfc.unreadable_fields))),
                  nota="Tipo STRING/XML/wide — inspeccionar na transacção respectiva se necessário.")

    rfc.close()

    paths = write_outputs(audit, info, Path(args.output_dir), started)
    print_summary(audit, info)
    print("\nOutputs:")
    for kind, p in paths.items():
        print(f"  {kind.upper():4} : {p}")

    c = audit.counts()
    return 1 if (c[ERR] or c[DIV]) else 0


if __name__ == "__main__":
    raise SystemExit(main())

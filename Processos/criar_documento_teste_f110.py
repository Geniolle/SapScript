# -*- coding: utf-8 -*-
"""
Cria um documento FI de fornecedor via RFC para teste da F110 / SEPA V9.

Fluxo:
    1. Carrega configuracao SAP a partir do .env
    2. Liga ao SAP via PyRFC
    3. Bloqueia execucao fora de QAD
    4. Monta o documento FI
    5. Executa BAPI_ACC_DOCUMENT_CHECK
    6. Confirma com o utilizador antes de postar
    7. Executa BAPI_ACC_DOCUMENT_POST
    8. Faz COMMIT ou ROLLBACK conforme o resultado

O script nao usa SAP GUI Scripting.
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

try:
    from pyrfc import Connection
    HAS_PYRFC = True
except Exception as exc:  # pragma: no cover - depende do ambiente local
    HAS_PYRFC = False
    PYRFC_IMPORT_ERROR = exc

try:
    from sap_agent.config import SapConnectionConfig
except Exception:  # pragma: no cover - caminho alternativo de execucao
    SapConnectionConfig = None  # type: ignore[assignment]


# =============================================================================
# (1) CONFIGURACAO BASE
# =============================================================================

ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_SYSTEM_KEY = "QAD"
ALLOWED_SYSTEM_IDS = {"QAD", "S4Q"}
ALLOWED_SYSTEM_KEYS = {"QAD", "S4Q", "S4QCLNT100"}
DEFAULT_CURRENCY = "EUR"
DEFAULT_DOC_TEXT = "TESTE SEPA V9"
DEFAULT_REFERENCE_PREFIX = "SEPA-V9-TEST"
DEFAULT_DOC_TYPE = "KR"
DEFAULT_PAYMENT_METHOD = "S"
DEFAULT_DAYS_BEFORE_DOC_DATE = 60
DEFAULT_POSTING_DATE_OFFSET_DAYS = 0


if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _flush_print(*args: Any, **kwargs: Any) -> None:
    kwargs.setdefault("flush", True)
    print(*args, **kwargs)


# =============================================================================
# (2) CARREGAMENTO DO .ENV E RESOLUCAO RFC
# =============================================================================

def load_project_dotenv() -> None:
    """Carrega o .env do projeto sem imprimir segredos."""

    candidates = [
        ROOT_DIR / ".env",
        Path.cwd() / ".env",
    ]

    for env_path in candidates:
        if env_path.exists():
            load_dotenv(env_path, override=False)


def _first_env(*names: str, default: str = "", required: bool = True) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value

    if required:
        joined = ", ".join(names)
        raise RuntimeError(f"Falta definir uma das variaveis de ambiente: {joined}")

    return default


def _normalize_system_key(raw_key: str | None) -> str:
    value = str(raw_key or "").strip().upper()
    if not value:
        return DEFAULT_SYSTEM_KEY
    return value


def _candidate_env_keys(system_key: str) -> list[str]:
    normalized = _normalize_system_key(system_key)
    mapping = {
        "QAD": ["QAD", "S4Q", "S4QCLNT100"],
        "S4Q": ["S4Q", "QAD", "S4QCLNT100"],
        "S4QCLNT100": ["S4QCLNT100", "S4Q", "QAD"],
    }
    keys = mapping.get(normalized, [normalized])
    if normalized not in keys:
        keys.insert(0, normalized)
    return keys


def resolve_connection_params(system_key: str | None = None) -> dict[str, str]:
    """
    Resolve os parametros de ligacao SAP sem expor passwords.

    Reutiliza SapConnectionConfig quando existem variaveis genericas SAP_*
    e cai para o padrao por alias quando o projeto usa SAP_ASHOST_QAD,
    SAP_USER_S4Q, SAP_PASSWORD_S4QCLNT100, etc.
    """

    requested_key = _normalize_system_key(system_key or os.getenv("SAP_SYSTEM") or os.getenv("WORKFLOW_SAP_KEY"))
    keys_to_try = _candidate_env_keys(requested_key)

    generic_ready = all(
        os.getenv(name, "").strip()
        for name in ("SAP_USER", "SAP_PASSWD", "SAP_ASHOST", "SAP_SYSNR", "SAP_CLIENT")
    )
    if generic_ready and SapConnectionConfig is not None:
        cfg = SapConnectionConfig.from_env()
        return cfg.as_pyrfc_params()

    ashost = _first_env(*(f"SAP_ASHOST_{key}" for key in keys_to_try), "SAP_ASHOST")
    sysnr = _first_env(*(f"SAP_SYSNR_{key}" for key in keys_to_try), "SAP_SYSNR", default="00", required=False) or "00"
    client = _first_env(*(f"SAP_CLIENT_{key}" for key in keys_to_try), "SAP_CLIENT")
    user = _first_env(*(f"SAP_USER_{key}" for key in keys_to_try), "SAP_USER")
    passwd = _first_env(
        *(f"SAP_PASSWORD_{key}" for key in keys_to_try),
        *(f"SAP_PASSWD_{key}" for key in keys_to_try),
        "SAP_PASSWORD",
        "SAP_PASSWD",
    )
    lang = _first_env(*(f"SAP_LANGUAGE_{key}" for key in keys_to_try), "SAP_LANG", "SAP_LANGUAGE", default="PT", required=False) or "PT"

    return {
        "ashost": ashost,
        "sysnr": sysnr,
        "client": client,
        "user": user,
        "passwd": passwd,
        "lang": lang,
    }


def open_rfc_connection(system_key: str | None = None) -> tuple[Connection, str]:
    if not HAS_PYRFC:
        raise RuntimeError(f"A biblioteca pyrfc nao esta disponivel: {PYRFC_IMPORT_ERROR}")

    params = resolve_connection_params(system_key)
    conn = Connection(**params)
    return conn, params["user"]


def get_system_info(conn: Connection) -> tuple[str, str]:
    result = conn.call("RFC_SYSTEM_INFO")
    info = result.get("RFCSI_EXPORT", {}) or {}
    system_id = str(info.get("RFCSYSID") or info.get("RFCSYSTEMID") or "").strip().upper()
    client = str(info.get("RFCCLIENT") or "").strip()
    return system_id, client


# =============================================================================
# (3) ENTRADAS E VALIDACOES
# =============================================================================

def _prompt_text(label: str, default: str = "", *, required: bool = False) -> str:
    suffix = f" [{default}]" if default else ""
    while True:
        value = input(f"{label}{suffix}: ").strip()
        if value:
            return value
        if default:
            return default
        if not required:
            return ""
        _flush_print("Informe um valor valido.")


def _prompt_decimal(label: str, default: str = "") -> Decimal:
    while True:
        raw = _prompt_text(label, default=default, required=True)
        normalized = raw.replace(" ", "").replace(".", "").replace(",", ".")
        try:
            value = Decimal(normalized)
        except (InvalidOperation, ValueError):
            _flush_print("Informe um valor numerico valido. Exemplo: 10,00")
            continue
        if value <= 0:
            _flush_print("Informe um valor positivo maior que zero.")
            continue
        return value.quantize(Decimal("0.01"))


def _parse_decimal_cli(raw: str) -> Decimal:
    text = str(raw or "").strip()
    normalized = text.replace(" ", "").replace(".", "").replace(",", ".")
    value = Decimal(normalized)
    if value <= 0:
        raise argparse.ArgumentTypeError("O valor deve ser positivo.")
    return value.quantize(Decimal("0.01"))


def _alpha(value: str, size: int = 10) -> str:
    text = str(value or "").strip()
    if text.isdigit():
        return text.zfill(size)
    return text


def _safe_date(value: str, default_value: str) -> str:
    raw = (value or "").strip()
    if not raw:
        return default_value
    try:
        datetime.strptime(raw, "%Y%m%d")
    except ValueError as exc:
        raise ValueError(f"Data invalida '{raw}'. Use o formato YYYYMMDD.") from exc
    return raw


@dataclass(frozen=True)
class DocumentInput:
    company_code: str
    vendor: str
    gl_account: str
    amount: Decimal
    currency: str
    doc_date: str
    posting_date: str
    payment_terms: str
    payment_method: str
    item_text: str
    header_text: str
    reference: str
    cost_center: str = ""
    profit_center: str = ""


def collect_document_input(args: argparse.Namespace) -> DocumentInput:
    today = date.today()
    default_doc_date = (today - timedelta(days=DEFAULT_DAYS_BEFORE_DOC_DATE)).strftime("%Y%m%d")
    default_posting_date = (today + timedelta(days=DEFAULT_POSTING_DATE_OFFSET_DAYS)).strftime("%Y%m%d")
    default_reference = f"{DEFAULT_REFERENCE_PREFIX}-{datetime.now().strftime('%H%M%S')}"[:16]

    company_code = (args.company_code or _prompt_text("Company code", required=True)).strip()
    vendor = _alpha(args.vendor or _prompt_text("Vendor", required=True))
    gl_account = _alpha(args.gl_account or _prompt_text("GL account", required=True))
    amount = args.amount if args.amount is not None else _prompt_decimal("Amount", default="10,00")
    currency = (args.currency or _prompt_text("Currency", default=DEFAULT_CURRENCY)).strip().upper()
    doc_date = _safe_date(args.doc_date or _prompt_text("Document date YYYYMMDD", default=default_doc_date), default_doc_date)
    posting_date = _safe_date(args.posting_date or _prompt_text("Posting date YYYYMMDD", default=default_posting_date), default_posting_date)
    payment_terms = (args.payment_terms or _prompt_text("Payment terms", default="")).strip()
    payment_method = (args.payment_method or _prompt_text("Payment method", default=DEFAULT_PAYMENT_METHOD)).strip().upper()
    item_text = (args.item_text or _prompt_text("Item text", default=DEFAULT_DOC_TEXT)).strip()
    header_text = (args.header_text or _prompt_text("Header text", default=item_text or DEFAULT_DOC_TEXT)).strip()
    reference = (args.reference or _prompt_text("Reference", default=default_reference)).strip()[:16]
    cost_center = (args.cost_center or _prompt_text("Cost center", default="")).strip()
    profit_center = (args.profit_center or _prompt_text("Profit center", default="")).strip()

    return DocumentInput(
        company_code=company_code,
        vendor=vendor,
        gl_account=gl_account,
        amount=amount,
        currency=currency,
        doc_date=doc_date,
        posting_date=posting_date,
        payment_terms=payment_terms,
        payment_method=payment_method,
        item_text=item_text,
        header_text=header_text,
        reference=reference,
        cost_center=cost_center,
        profit_center=profit_center,
    )


# =============================================================================
# (4) MONTAGEM DO DOCUMENTO SAP
# =============================================================================

def build_bapi_payload(data: DocumentInput, username: str) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    vendor = _alpha(data.vendor)
    gl_account = _alpha(data.gl_account)

    document_header = {
        "USERNAME": username,
        "BUS_ACT": "RFBU",
        "COMP_CODE": data.company_code,
        "DOC_DATE": data.doc_date,
        "PSTNG_DATE": data.posting_date,
        "DOC_TYPE": DEFAULT_DOC_TYPE,
        "HEADER_TXT": data.header_text,
        "REF_DOC_NO": data.reference,
    }

    accountgl_line: dict[str, Any] = {
        "ITEMNO_ACC": "0000000001",
        "GL_ACCOUNT": gl_account,
        "COMP_CODE": data.company_code,
        "ITEM_TEXT": data.item_text,
    }
    if data.cost_center:
        accountgl_line["COSTCENTER"] = data.cost_center
    if data.profit_center:
        accountgl_line["PROFIT_CTR"] = data.profit_center

    accountpayable_line: dict[str, Any] = {
        "ITEMNO_ACC": "0000000002",
        "VENDOR_NO": vendor,
        "COMP_CODE": data.company_code,
        "ITEM_TEXT": data.item_text,
        "PYMT_METH": data.payment_method,
    }
    if data.payment_terms:
        accountpayable_line["PMNTTRMS"] = data.payment_terms

    account_gl = [accountgl_line]
    account_payable = [accountpayable_line]
    currency_amount = [
        {
            "ITEMNO_ACC": "0000000001",
            "CURRENCY": data.currency,
            "AMT_DOCCUR": data.amount,
        },
        {
            "ITEMNO_ACC": "0000000002",
            "CURRENCY": data.currency,
            "AMT_DOCCUR": -data.amount,
        },
    ]

    return document_header, account_gl, account_payable, currency_amount


def _print_bapi_return(messages: list[dict[str, Any]] | None) -> bool:
    has_error = False
    for msg in messages or []:
        msg_type = str(msg.get("TYPE", "")).strip().upper()
        msg_id = str(msg.get("ID", "")).strip()
        msg_no = str(msg.get("NUMBER", "")).strip()
        msg_text = str(msg.get("MESSAGE", "")).strip()

        if msg_type in {"E", "A", "X"}:
            has_error = True

        icon = {
            "S": "[OK]",
            "W": "[WARN]",
            "I": "[INFO]",
            "E": "[ERR]",
            "A": "[ERR]",
            "X": "[ERR]",
        }.get(msg_type, "[MSG]")

        _flush_print(f"{icon} {msg_type} {msg_id}{msg_no}: {msg_text}")

    return has_error


def _parse_obj_key(obj_key: str) -> tuple[str, str, str] | None:
    value = str(obj_key or "").strip()
    if len(value) < 18:
        return None
    belnr = value[:10]
    bukrs = value[10:14]
    gjahr = value[14:18]
    return belnr, bukrs, gjahr


def _safe_rollback(conn: Connection | None) -> None:
    if conn is None:
        return
    try:
        conn.call("BAPI_TRANSACTION_ROLLBACK")
    except Exception:
        pass


def _safe_commit(conn: Connection) -> None:
    conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")


# =============================================================================
# (5) EXECUCAO PRINCIPAL
# =============================================================================

def main(argv: list[str] | None = None) -> int:
    load_project_dotenv()

    parser = argparse.ArgumentParser(
        description="Criar documento FI de fornecedor via RFC para teste da F110 / SEPA V9."
    )
    parser.add_argument("--sap-system", default=os.getenv("SAP_SYSTEM") or os.getenv("WORKFLOW_SAP_KEY") or DEFAULT_SYSTEM_KEY, help="Chave SAP do ambiente. Padrao: QAD")
    parser.add_argument("--company-code", default=os.getenv("SAP_COMPANY_CODE") or os.getenv("FI_COMPANY_CODE") or os.getenv("BUKRS", ""), help="Codigo da empresa (BUKRS)")
    parser.add_argument("--vendor", default=os.getenv("SAP_VENDOR") or os.getenv("FI_VENDOR", ""), help="Numero do fornecedor")
    parser.add_argument("--gl-account", default=os.getenv("SAP_GL_ACCOUNT") or os.getenv("FI_GL_ACCOUNT", ""), help="Conta de contrapartida GL")
    parser.add_argument("--amount", type=_parse_decimal_cli, default=None, help="Valor do documento")
    parser.add_argument("--currency", default=os.getenv("SAP_CURRENCY") or os.getenv("FI_CURRENCY") or DEFAULT_CURRENCY, help="Moeda")
    parser.add_argument("--doc-date", default=os.getenv("SAP_DOC_DATE") or os.getenv("FI_DOC_DATE", ""), help="Data do documento em YYYYMMDD")
    parser.add_argument("--posting-date", default=os.getenv("SAP_POSTING_DATE") or os.getenv("FI_POSTING_DATE", ""), help="Data de lancamento em YYYYMMDD")
    parser.add_argument("--payment-terms", default=os.getenv("SAP_PAYMENT_TERMS") or os.getenv("FI_PAYMENT_TERMS", ""), help="Condicao de pagamento")
    parser.add_argument("--payment-method", default=os.getenv("SAP_PAYMENT_METHOD") or os.getenv("FI_PAYMENT_METHOD") or DEFAULT_PAYMENT_METHOD, help="Metodo de pagamento")
    parser.add_argument("--item-text", default=os.getenv("SAP_ITEM_TEXT") or os.getenv("FI_ITEM_TEXT") or DEFAULT_DOC_TEXT, help="Texto da linha")
    parser.add_argument("--header-text", default=os.getenv("SAP_HEADER_TEXT") or os.getenv("FI_HEADER_TEXT"), help="Texto do cabecalho")
    parser.add_argument("--reference", default=os.getenv("SAP_REFERENCE") or os.getenv("FI_REFERENCE"), help="Referencia externa")
    parser.add_argument("--cost-center", default=os.getenv("SAP_COST_CENTER") or os.getenv("FI_COST_CENTER", ""), help="Centro de custo opcional")
    parser.add_argument("--profit-center", default=os.getenv("SAP_PROFIT_CENTER") or os.getenv("FI_PROFIT_CENTER", ""), help="Centro de lucro opcional")

    args = parser.parse_args(argv)

    _flush_print("=" * 84)
    _flush_print("SAP - CRIAR DOCUMENTO FI FORNECEDOR PARA TESTE F110 / SEPA V9")
    _flush_print("=" * 84)

    conn: Connection | None = None

    try:
        conn, user = open_rfc_connection(args.sap_system)
        conn.call("RFC_PING")
        system_id, client = get_system_info(conn)

        if system_id not in ALLOWED_SYSTEM_IDS:
            raise RuntimeError(
                f"Execucao bloqueada: o sistema ligado e '{system_id or 'DESCONHECIDO'}'. "
                "Este script e exclusivo para QAD."
            )

        _flush_print(f"Conexao SAP ok | System={system_id} | Client={client or '?'} | User={user}")

        document_data = collect_document_input(args)

        if _normalize_system_key(args.sap_system) not in ALLOWED_SYSTEM_KEYS:
            _flush_print(
                "Aviso: a chave informada nao e um alias esperado. "
                "A ligacao foi validada pelo system id retornado pelo SAP."
            )

        document_header, account_gl, account_payable, currency_amount = build_bapi_payload(document_data, user)

        _flush_print("")
        _flush_print("Documento a validar")
        _flush_print("-" * 84)
        _flush_print(f"Company code   : {document_data.company_code}")
        _flush_print(f"Vendor         : {document_data.vendor}")
        _flush_print(f"GL account     : {document_data.gl_account}")
        _flush_print(f"Amount         : {document_data.amount} {document_data.currency}")
        _flush_print(f"Doc date       : {document_data.doc_date}")
        _flush_print(f"Posting date    : {document_data.posting_date}")
        _flush_print(f"Payment terms   : {document_data.payment_terms or '(master data)'}")
        _flush_print(f"Payment method  : {document_data.payment_method}")
        _flush_print(f"Reference       : {document_data.reference}")
        _flush_print(f"Header text     : {document_data.header_text}")

        _flush_print("")
        _flush_print("A executar BAPI_ACC_DOCUMENT_CHECK...")
        check_result = conn.call(
            "BAPI_ACC_DOCUMENT_CHECK",
            DOCUMENTHEADER=document_header,
            ACCOUNTGL=account_gl,
            ACCOUNTPAYABLE=account_payable,
            CURRENCYAMOUNT=currency_amount,
        )

        if _print_bapi_return(check_result.get("RETURN", [])):
            _flush_print("Documento nao passou na validacao SAP. A fazer rollback.")
            _safe_rollback(conn)
            return 1

        confirm = input("\nPara criar o documento, escreva exatamente POSTAR: ").strip()
        if confirm != "POSTAR":
            _flush_print("Lancamento cancelado pelo utilizador.")
            _safe_rollback(conn)
            return 0

        _flush_print("")
        _flush_print("A executar BAPI_ACC_DOCUMENT_POST...")
        post_result = conn.call(
            "BAPI_ACC_DOCUMENT_POST",
            DOCUMENTHEADER=document_header,
            ACCOUNTGL=account_gl,
            ACCOUNTPAYABLE=account_payable,
            CURRENCYAMOUNT=currency_amount,
        )

        if _print_bapi_return(post_result.get("RETURN", [])):
            _flush_print("SAP rejeitou o lancamento. A fazer rollback.")
            _safe_rollback(conn)
            return 1

        _safe_commit(conn)

        obj_key = str(post_result.get("OBJ_KEY", "")).strip()
        parsed = _parse_obj_key(obj_key)

        _flush_print("")
        _flush_print("=" * 84)
        _flush_print("Documento criado com sucesso")
        _flush_print("=" * 84)
        _flush_print(f"OBJ_KEY SAP : {obj_key or '(vazio)'}")
        if parsed:
            belnr, bukrs, gjahr = parsed
            _flush_print(f"Documento   : {belnr}")
            _flush_print(f"Empresa     : {bukrs}")
            _flush_print(f"Exercicio   : {gjahr}")
        _flush_print("Pode validar na FB03 e seguir com a F110 no ambiente QAD.")
        return 0

    except KeyboardInterrupt:
        _flush_print("")
        _flush_print("Execucao interrompida pelo utilizador. A fazer rollback.")
        _safe_rollback(conn)
        return 130
    except Exception as exc:
        _flush_print(f"Erro: {exc}")
        _safe_rollback(conn)
        return 1
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())

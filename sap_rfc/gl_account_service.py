from __future__ import annotations

import os
import re
from typing import Any

from sap_rfc._rfc_common import build_connection_params_for, load_project_env, find_project_root

ACCOUNT_RE = re.compile(r"^\d{1,10}$")
COMPANY_RE = re.compile(r"^[A-Z0-9]{4}$")


def _account(value: str) -> str:
    raw = str(value or "").strip()
    if not ACCOUNT_RE.fullmatch(raw):
        raise ValueError("A conta deve conter entre 1 e 10 algarismos.")
    return raw.zfill(10)


def _company(value: str) -> str:
    raw = str(value or "").strip().upper()
    if not COMPANY_RE.fullmatch(raw):
        raise ValueError("A empresa deve conter exatamente 4 caracteres alfanuméricos.")
    return raw


def _messages_with_error(messages: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [item for item in messages if str(item.get("TYPE", "")).upper() in {"E", "A", "X"}]


def _read_one(connection: Any, table: str, fields: tuple[str, ...], options: list[str]) -> dict[str, str] | None:
    result = connection.call(
        "RFC_READ_TABLE", QUERY_TABLE=table, DELIMITER="|",
        FIELDS=[{"FIELDNAME": field} for field in fields],
        OPTIONS=[{"TEXT": text} for text in options], ROWCOUNT=1,
    )
    rows = result.get("DATA", []) or []
    if not rows:
        return None
    return dict(zip(fields, [part.strip() for part in str(rows[0].get("WA", "")).split("|")]))


def create_company_account_by_model(
    *, environment: str, account: str, target_company: str,
    model_company: str, alternative_account: str = "", test_only: bool = False,
    connection_factory: Any = None,
) -> dict[str, Any]:
    """Estende uma conta existente no plano operacional a outra empresa.

    Copia apenas o segmento empresarial. O segmento do plano e a conta de grupo
    têm de existir previamente. Para empresas fora de PT, a conta alternativa é
    obrigatória e tem de existir no plano alternativo da empresa.
    """
    env = str(environment or "").strip().upper()
    if env not in {"DEV", "QAD", "PRD"}:
        raise ValueError("Ambiente inválido. Utilize DEV, QAD ou PRD.")
    saknr = _account(account)
    target = _company(target_company)
    model = _company(model_company)
    if target == model:
        raise ValueError("A empresa de destino deve ser diferente da empresa-modelo.")

    load_project_env(find_project_root())
    if connection_factory is None:
        from pyrfc import Connection
        connection_factory = Connection
    connection = connection_factory(**build_connection_params_for(env))
    try:
        connection.call("RFC_PING")
        target_info = _read_one(connection, "T001", ("BUKRS", "LAND1", "KTOPL", "KTOP2"), [f"BUKRS = '{target}'"])
        model_info = _read_one(connection, "T001", ("BUKRS", "LAND1", "KTOPL", "KTOP2"), [f"BUKRS = '{model}'"])
        if not target_info or not model_info:
            raise RuntimeError("Empresa de destino ou empresa-modelo não encontrada.")
        if target_info["KTOPL"] != model_info["KTOPL"]:
            raise RuntimeError("As empresas não utilizam o mesmo plano de contas operacional.")

        coa = _read_one(connection, "SKA1", ("KTOPL", "SAKNR", "BILKT", "XLOEV"), [f"KTOPL = '{target_info['KTOPL']}'", f"AND SAKNR = '{saknr}'"])
        if not coa or coa.get("XLOEV") == "X":
            raise RuntimeError("A conta não existe ou está marcada para eliminação no plano operacional.")
        if not coa.get("BILKT"):
            raise RuntimeError("A conta operacional não possui conta de grupo atribuída.")
        group = _read_one(connection, "SKA1", ("KTOPL", "SAKNR", "XLOEV"), ["KTOPL = 'GC01'", f"AND SAKNR = '{coa['BILKT']}'"])
        if not group or group.get("XLOEV") == "X":
            raise RuntimeError(f"A conta de grupo GC01/{coa['BILKT']} não existe ou está inativa.")

        exists = _read_one(connection, "SKB1", ("BUKRS", "SAKNR"), [f"BUKRS = '{target}'", f"AND SAKNR = '{saknr}'"])
        if exists:
            raise RuntimeError(f"A conta {saknr} já existe na empresa {target}.")

        alt = _account(alternative_account) if str(alternative_account or "").strip() else ""
        if target_info.get("LAND1") != "PT":
            if not alt:
                raise ValueError(f"A conta alternativa é obrigatória para a empresa {target} ({target_info.get('LAND1')}).")
            alt_master = _read_one(connection, "SKA1", ("KTOPL", "SAKNR", "XLOEV"), [f"KTOPL = '{target_info['KTOP2']}'", f"AND SAKNR = '{alt}'"])
            if not alt_master or alt_master.get("XLOEV") == "X":
                raise RuntimeError(f"A conta alternativa {target_info['KTOP2']}/{alt} não existe ou está inativa.")

        source = connection.call(
            "GL_ACCT_MASTER_GET_CCODE_RFC",
            ACCOUNT_CCODE={"KEYY": {"BUKRS": model, "SAKNR": saknr}}, RETURN=[],
        )
        source_errors = _messages_with_error(list(source.get("RETURN", []) or []))
        if source_errors:
            raise RuntimeError("Não foi possível ler a empresa-modelo: " + source_errors[0].get("MESSAGE", ""))
        row = {
            "KEYY": {"BUKRS": target, "SAKNR": saknr},
            "DATA": dict(source["ACCOUNT_CCODE"]["DATA"], ALTKT=alt),
            "INFO": {}, "ACTION": "I",
        }
        preview = connection.call(
            "GL_ACCT_MASTER_SAVE_RFC", ACCOUNT_CCODES=[row], RETURN=[],
            NO_AUTHORITY_CHECK="", NO_SAVE_AT_WARNING="", TESTMODE="X",
        )
        preview_messages = list(preview.get("RETURN", []) or [])
        if _messages_with_error(preview_messages):
            return {"ok": False, "status": "PREVIEW_ERROR", "environment": env, "messages": preview_messages}
        base = {
            "ok": True, "environment": env, "account": saknr, "target_company": target,
            "model_company": model, "country": target_info.get("LAND1", ""),
            "alternative_chart": target_info.get("KTOP2", ""), "alternative_account": alt,
            "group_account": coa.get("BILKT", ""), "preview_messages": preview_messages,
        }
        if test_only:
            return dict(base, status="PREVIEW_READY")

        saved = connection.call(
            "GL_ACCT_MASTER_SAVE_RFC", ACCOUNT_CCODES=[row], RETURN=[],
            NO_AUTHORITY_CHECK="", NO_SAVE_AT_WARNING="", TESTMODE="",
        )
        save_messages = list(saved.get("RETURN", []) or [])
        if _messages_with_error(save_messages):
            connection.call("BAPI_TRANSACTION_ROLLBACK")
            return dict(base, ok=False, status="SAVE_ERROR", messages=save_messages)
        commit = connection.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
        commit_return = dict(commit.get("RETURN", {}) or {})
        if _messages_with_error([commit_return]):
            raise RuntimeError("Falha no commit: " + str(commit_return.get("MESSAGE", "")))
        verified = _read_one(connection, "SKB1", ("BUKRS", "SAKNR", "ALTKT", "FSTAG", "ZUAWA"), [f"BUKRS = '{target}'", f"AND SAKNR = '{saknr}'"])
        if not verified:
            raise RuntimeError("A gravação terminou sem erro, mas a conta não foi encontrada na verificação final.")
        return dict(base, status="CREATED", messages=save_messages, verified=verified)
    finally:
        connection.close()

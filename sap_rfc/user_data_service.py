"""Read-only: dados de um utilizador SAP no ambiente pedido.

kind = "master"   -> dados mestre (USR02 + USR01): tipo, grupo, validade,
                     bloqueio, ultimo logon, defaults.
kind = "personal" -> dados pessoais (USR21 -> ADRP / ADCP / ADR6): nome,
                     departamento, funcao, telefone, email.

O ambiente vem de PFCG_TARGET_ENV (posto pelo worker); default PRD.
"""

from __future__ import annotations

import os
import re
from typing import Any

from sap_rfc._rfc_common import (
    build_connection_params_for,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_option_eq,
    make_read_only_guard,
    read_table,
    resolve_target_env,
)

ALLOWED_TABLES = ("USR02", "USR01", "USR21", "ADRP", "ADCP", "ADR6")

_USER_RE = re.compile(r"^[A-Za-z0-9_.\-]{1,12}$")

_USTYP_LABELS = {
    "A": "Dialog",
    "B": "System (BDC)",
    "C": "Communication",
    "L": "Reference",
    "S": "Service",
}


def validate_username(raw: str) -> str:
    value = str(raw or "").strip().upper()
    if not value:
        raise ValueError("Indique um utilizador SAP.")
    if not _USER_RE.match(value):
        raise ValueError("Utilizador invalido: use apenas letras, numeros, _, . ou - (max. 12).")
    return value


def _error_result(username: str, kind: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "username": username,
        "kind": kind,
        "error_type": error_type,
        "message": message,
        "system": resolve_target_env(),
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _fmt_date(raw: str) -> str | None:
    raw = str(raw or "").strip()
    if len(raw) == 8 and raw.isdigit() and raw != "00000000":
        return f"{raw[6:8]}/{raw[4:6]}/{raw[0:4]}"
    return None


def _one(rows: list[list[str]]) -> list[str] | None:
    return rows[0] if rows else None


def _collect_master(connection: Any, guard: Any, username: str) -> tuple[list[dict[str, str]], list[str]]:
    warnings: list[str] = []
    fields: list[dict[str, str]] = []

    try:
        row = _one(read_table(
            connection, guard, table_name="USR02",
            fields=["BNAME", "USTYP", "CLASS", "GLTGV", "GLTGB", "TRDAT", "UFLAG"],
            options=make_option_eq("BNAME", username), rowcount=1,
        ))
    except Exception as exc:
        raise
    if not row:
        return [], warnings  # utilizador nao existe

    ustyp = row[1].strip()
    uflag = row[6].strip()
    fields += [
        {"label": "Tipo", "value": f"{ustyp} - {_USTYP_LABELS.get(ustyp, ustyp)}" if ustyp else "-"},
        {"label": "Grupo", "value": row[2].strip() or "-"},
        {"label": "Valido de", "value": _fmt_date(row[3]) or "-"},
        {"label": "Valido ate", "value": _fmt_date(row[4]) or "sem limite"},
        {"label": "Ultimo logon", "value": _fmt_date(row[5]) or "nunca"},
        {"label": "Estado", "value": "Bloqueado" if (uflag and uflag not in ("0", "")) else "Ativo"},
    ]

    try:
        d = _one(read_table(
            connection, guard, table_name="USR01",
            fields=["BNAME", "DATFM", "DCPFM", "SPLD", "LANGU"],
            options=make_option_eq("BNAME", username), rowcount=1,
        ))
        if d:
            fields += [
                {"label": "Formato de data", "value": d[1].strip() or "-"},
                {"label": "Formato decimal", "value": d[2].strip() or "-"},
                {"label": "Impressora", "value": d[3].strip() or "-"},
                {"label": "Idioma", "value": d[4].strip() or "-"},
            ]
    except Exception as exc:
        warnings.append(
            "Sem autorizacao para consultar USR01."
            if is_authorization_error(exc) else f"USR01: {classify_rfc_error(exc)[1]}"
        )
    return fields, warnings


def _collect_personal(connection: Any, guard: Any, username: str) -> tuple[list[dict[str, str]], list[str]]:
    warnings: list[str] = []

    link = _one(read_table(
        connection, guard, table_name="USR21",
        fields=["BNAME", "PERSNUMBER", "ADDRNUMBER"],
        options=make_option_eq("BNAME", username), rowcount=1,
    ))
    if not link:
        return [], warnings
    persnumber = link[1].strip()
    addrnumber = link[2].strip()

    fields: list[dict[str, str]] = []
    try:
        person = _one(read_table(
            connection, guard, table_name="ADRP",
            fields=["PERSNUMBER", "NAME_FIRST", "NAME_LAST", "NAME_TEXT"],
            options=make_option_eq("PERSNUMBER", persnumber), rowcount=1,
        ))
        if person:
            full = person[3].strip() or f"{person[1].strip()} {person[2].strip()}".strip()
            fields.append({"label": "Nome", "value": full or "-"})
    except Exception as exc:
        warnings.append(
            "Sem autorizacao para consultar ADRP."
            if is_authorization_error(exc) else f"ADRP: {classify_rfc_error(exc)[1]}"
        )

    try:
        comp = _one(read_table(
            connection, guard, table_name="ADCP",
            fields=["PERSNUMBER", "ADDRNUMBER", "DEPARTMENT", "FUNCTION", "TEL_NUMBER"],
            options=make_option_eq("PERSNUMBER", persnumber), rowcount=1,
        ))
        if comp:
            fields += [
                {"label": "Departamento", "value": comp[2].strip() or "-"},
                {"label": "Funcao", "value": comp[3].strip() or "-"},
                {"label": "Telefone", "value": comp[4].strip() or "-"},
            ]
    except Exception as exc:
        warnings.append(
            "Sem autorizacao para consultar ADCP."
            if is_authorization_error(exc) else f"ADCP: {classify_rfc_error(exc)[1]}"
        )

    try:
        mail_rows = read_table(
            connection, guard, table_name="ADR6",
            fields=["PERSNUMBER", "SMTP_ADDR"],
            options=make_option_eq("PERSNUMBER", persnumber), rowcount=5,
        )
        emails = sorted({r[1].strip() for r in mail_rows if r[1].strip()})
        if emails:
            fields.append({"label": "Email", "value": ", ".join(emails)})
    except Exception as exc:
        warnings.append(
            "Sem autorizacao para consultar ADR6."
            if is_authorization_error(exc) else f"ADR6: {classify_rfc_error(exc)[1]}"
        )
    return fields, warnings


def analyze_user_data(username: str, kind: str) -> dict[str, Any]:
    kind = str(kind or "").strip().lower()
    if kind not in ("master", "personal"):
        return _error_result(str(username or "").strip().upper(), kind, "INVALID_INPUT", "Tipo de analise invalido.")
    try:
        norm_user = validate_username(username)
    except ValueError as exc:
        return _error_result(str(username or "").strip().upper(), kind, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        target_env = resolve_target_env()
        params = build_connection_params_for(target_env)
    except Exception as exc:
        return _error_result(norm_user, kind, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(norm_user, kind, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(norm_user, kind, error_type, message, details=format_exception(exc))

    try:
        try:
            collector = _collect_master if kind == "master" else _collect_personal
            fields, warnings = collector(connection, guard, norm_user)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(norm_user, kind, "AUTHORIZATION_ERROR", "Sem autorizacao para ler os dados do utilizador.", details=format_exception(exc))
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_user, kind, error_type, message, details=format_exception(exc))

        if not fields:
            return {
                "ok": True, "status": "NAO_ENCONTRADO", "username": norm_user, "kind": kind,
                "fields": [], "system": target_env, "client": params["client"],
            }

        payload: dict[str, Any] = {
            "ok": True, "status": "OK", "username": norm_user, "kind": kind,
            "fields": fields, "system": target_env, "client": params["client"],
        }
        if warnings:
            payload["warning"] = " ".join(warnings)
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

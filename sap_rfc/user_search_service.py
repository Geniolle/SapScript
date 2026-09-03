"""Read-only: pesquisa utilizadores SAP por nome (ou parte do nome) no ambiente
pedido. Devolve nome completo + user SAP.

ADRP (nomes de pessoa) filtrado por LIKE em NAME_LAST/NAME_TEXT; liga a USR21
(BNAME <-> PERSNUMBER) para obter o utilizador.
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
    make_option_in,
    make_read_only_guard,
    read_table,
    resolve_target_env,
)

ALLOWED_TABLES = ("ADRP", "USR21")

_MAX_RESULTS = 100
_MAX_QUERY = 30


def sanitize_query(raw: str) -> str:
    value = re.sub(r"[%_'\"\\]", "", str(raw or "")).strip()
    if len(value) < 2:
        raise ValueError("Indique pelo menos 2 caracteres do nome a pesquisar.")
    return value[:_MAX_QUERY]


def _error_result(query: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "query": query,
        "error_type": error_type,
        "message": message,
        "system": resolve_target_env(),
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _like_options(query: str) -> list[dict[str, str]]:
    variants: list[str] = []
    for v in (query, query.upper(), query.title()):
        if v not in variants:
            variants.append(v)
    fields = ("NAME_LAST", "NAME_TEXT")
    rows: list[dict[str, str]] = []
    for field in fields:
        for v in variants:
            prefix = "OR " if rows else ""
            rows.append({"TEXT": f"{prefix}{field} LIKE '%{v}%'"})
    return rows


def search_users_by_name(query: str) -> dict[str, Any]:
    try:
        norm_query = sanitize_query(query)
    except ValueError as exc:
        return _error_result(str(query or "").strip(), "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        target_env = resolve_target_env()
        params = build_connection_params_for(target_env)
    except Exception as exc:
        return _error_result(norm_query, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(norm_query, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(norm_query, error_type, message, details=format_exception(exc))

    try:
        try:
            person_rows = read_table(
                connection, guard, table_name="ADRP",
                fields=["PERSNUMBER", "NAME_FIRST", "NAME_LAST", "NAME_TEXT"],
                options=_like_options(norm_query), rowcount=_MAX_RESULTS * 3,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(norm_query, "ADRP_AUTHORIZATION_ERROR", "Sem autorizacao para consultar ADRP.", details=format_exception(exc))
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_query, f"ADRP_{error_type}", message, details=format_exception(exc))

        by_pers: dict[str, str] = {}
        for persnumber, first, last, text in person_rows:
            pn = persnumber.strip()
            if not pn:
                continue
            full = text.strip() or f"{first.strip()} {last.strip()}".strip()
            by_pers.setdefault(pn, full)

        if not by_pers:
            return {
                "ok": True, "status": "NAO_ENCONTRADO", "query": norm_query,
                "count": 0, "users": [], "system": target_env, "client": params["client"],
            }

        warnings: list[str] = []
        try:
            link_rows = read_table(
                connection, guard, table_name="USR21",
                fields=["BNAME", "PERSNUMBER"],
                options=make_option_in("PERSNUMBER", list(by_pers.keys())), rowcount=0,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(norm_query, "USR21_AUTHORIZATION_ERROR", "Sem autorizacao para consultar USR21.", details=format_exception(exc))
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_query, f"USR21_{error_type}", message, details=format_exception(exc))

        seen: set[str] = set()
        users: list[dict[str, str]] = []
        for bname, persnumber in link_rows:
            bname = bname.strip()
            pn = persnumber.strip()
            if not bname or bname in seen or pn not in by_pers:
                continue
            seen.add(bname)
            users.append({"username": bname, "full_name": by_pers[pn]})

        users.sort(key=lambda u: (u["full_name"].upper(), u["username"]))
        truncated = len(users) > _MAX_RESULTS
        users = users[:_MAX_RESULTS]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK" if users else "NAO_ENCONTRADO",
            "query": norm_query,
            "count": len(users),
            "users": users,
            "system": target_env,
            "client": params["client"],
        }
        if truncated:
            payload["warning"] = f"Mais de {_MAX_RESULTS} resultados; mostro apenas os primeiros {_MAX_RESULTS}."
        if warnings:
            payload["warning"] = ((payload.get("warning", "") + " ") + " ".join(warnings)).strip()
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

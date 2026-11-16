"""Read-only: pesquisa utilizadores SAP por nome, email ou username no ambiente
pedido. Devolve nome completo + user SAP (+ email quando aplicável).

ADRP (nomes de pessoa) filtrado por LIKE em NAME_LAST/NAME_TEXT;
ADR6 (endereços de email) filtrado por LIKE em SMTP_ADDR;
USR21 (BNAME <-> PERSNUMBER) liga a pessoa ao utilizador;
USR02 (BNAME) valida utilizador direto.
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
    make_option_in,
    make_read_only_guard,
    read_table,
    resolve_target_env,
)

ALLOWED_TABLES = ("ADRP", "USR21", "ADR6", "USR02")

_MAX_RESULTS = 100
_MAX_QUERY = 50


def sanitize_query(raw: str, search_type: str = "name") -> str:
    raw_str = str(raw or "").strip()
    if search_type == "email":
        # Remove caracteres perigosos para SAP options mas preserva @ . - _ +
        value = re.sub(r"['\"\\]", "", raw_str).strip()
        if len(value) < 3 or "@" not in value:
            raise ValueError("Indique um email válido para pesquisar (mínimo 3 caracteres, incluindo '@').")
        return value[:_MAX_QUERY]
    elif search_type == "username":
        value = re.sub(r"[^A-Za-z0-9_.\-]", "", raw_str).strip().upper()
        if not value:
            raise ValueError("Indique o utilizador SAP a pesquisar.")
        return value[:12]
    else:
        value = re.sub(r"[%_'\"\\]", "", raw_str).strip()
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


def _like_options_name(query: str) -> list[dict[str, str]]:
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


def _like_options_email(query: str) -> list[dict[str, str]]:
    variants: list[str] = []
    for v in (query, query.lower(), query.upper()):
        if v not in variants:
            variants.append(v)
    rows: list[dict[str, str]] = []
    for v in variants:
        prefix = "OR " if rows else ""
        rows.append({"TEXT": f"{prefix}SMTP_ADDR LIKE '%{v}%'"})
    return rows


def search_users(query: str, search_type: str = "name") -> dict[str, Any]:
    search_type = str(search_type or "name").strip().lower()
    if search_type not in ("name", "email", "username"):
        search_type = "name"

    try:
        norm_query = sanitize_query(query, search_type)
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
        users: list[dict[str, Any]] = []
        warnings: list[str] = []

        if search_type == "username":
            # Pesquisa direta por BNAME em USR02 / USR21
            try:
                user_rows = read_table(
                    connection, guard, table_name="USR02",
                    fields=["BNAME", "UFLAG"],
                    options=make_option_eq("BNAME", norm_query), rowcount=1,
                )
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(norm_query, "USR02_AUTHORIZATION_ERROR", "Sem autorização para consultar USR02.", details=format_exception(exc))
                error_type, message = classify_rfc_error(exc)
                return _error_result(norm_query, f"USR02_{error_type}", message, details=format_exception(exc))

            if not user_rows:
                return {
                    "ok": True, "status": "NAO_ENCONTRADO", "query": norm_query,
                    "search_type": search_type, "count": 0, "users": [], "system": target_env, "client": params["client"],
                }

            bname = user_rows[0][0].strip()
            # Resolver nome e email via USR21 -> ADRP, ADR6
            full_name = bname
            email = ""
            try:
                link_rows = read_table(
                    connection, guard, table_name="USR21",
                    fields=["BNAME", "PERSNUMBER"],
                    options=make_option_eq("BNAME", bname), rowcount=1,
                )
                if link_rows and link_rows[0][1].strip():
                    persnumber = link_rows[0][1].strip()
                    adrp_rows = read_table(
                        connection, guard, table_name="ADRP",
                        fields=["PERSNUMBER", "NAME_FIRST", "NAME_LAST", "NAME_TEXT"],
                        options=make_option_eq("PERSNUMBER", persnumber), rowcount=1,
                    )
                    if adrp_rows:
                        first, last, text = adrp_rows[0][1].strip(), adrp_rows[0][2].strip(), adrp_rows[0][3].strip()
                        full_name = text or f"{first} {last}".strip() or bname

                    adr6_rows = read_table(
                        connection, guard, table_name="ADR6",
                        fields=["PERSNUMBER", "SMTP_ADDR"],
                        options=make_option_eq("PERSNUMBER", persnumber), rowcount=1,
                    )
                    if adr6_rows:
                        email = adr6_rows[0][1].strip()
            except Exception:
                pass

            users.append({"username": bname, "full_name": full_name, "email": email})

        elif search_type == "email":
            # Pesquisa por email em ADR6 -> USR21 -> ADRP
            try:
                adr6_rows = read_table(
                    connection, guard, table_name="ADR6",
                    fields=["PERSNUMBER", "SMTP_ADDR"],
                    options=_like_options_email(norm_query), rowcount=_MAX_RESULTS * 2,
                )
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(norm_query, "ADR6_AUTHORIZATION_ERROR", "Sem autorização para consultar ADR6.", details=format_exception(exc))
                error_type, message = classify_rfc_error(exc)
                return _error_result(norm_query, f"ADR6_{error_type}", message, details=format_exception(exc))

            by_pers_email: dict[str, str] = {}
            for persnumber, smtp_addr in adr6_rows:
                pn = persnumber.strip()
                if not pn:
                    continue
                by_pers_email.setdefault(pn, smtp_addr.strip())

            if not by_pers_email:
                return {
                    "ok": True, "status": "NAO_ENCONTRADO", "query": norm_query,
                    "search_type": search_type, "count": 0, "users": [], "system": target_env, "client": params["client"],
                }

            try:
                link_rows = read_table(
                    connection, guard, table_name="USR21",
                    fields=["BNAME", "PERSNUMBER"],
                    options=make_option_in("PERSNUMBER", list(by_pers_email.keys())), rowcount=0,
                )
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(norm_query, "USR21_AUTHORIZATION_ERROR", "Sem autorização para consultar USR21.", details=format_exception(exc))
                error_type, message = classify_rfc_error(exc)
                return _error_result(norm_query, f"USR21_{error_type}", message, details=format_exception(exc))

            pns_to_fetch = list({persnumber.strip() for _, persnumber in link_rows if persnumber.strip()})
            by_pers_name: dict[str, str] = {}
            if pns_to_fetch:
                try:
                    adrp_rows = read_table(
                        connection, guard, table_name="ADRP",
                        fields=["PERSNUMBER", "NAME_FIRST", "NAME_LAST", "NAME_TEXT"],
                        options=make_option_in("PERSNUMBER", pns_to_fetch), rowcount=0,
                    )
                    for persnumber, first, last, text in adrp_rows:
                        pn = persnumber.strip()
                        full = text.strip() or f"{first.strip()} {last.strip()}".strip()
                        by_pers_name[pn] = full
                except Exception:
                    pass

            seen: set[str] = set()
            for bname, persnumber in link_rows:
                bname = bname.strip()
                pn = persnumber.strip()
                if not bname or bname in seen:
                    continue
                seen.add(bname)
                users.append({
                    "username": bname,
                    "full_name": by_pers_name.get(pn, bname),
                    "email": by_pers_email.get(pn, ""),
                })

        else:
            # Pesquisa por nome (NAME_LAST / NAME_TEXT em ADRP) -> USR21 -> ADR6 (opcional)
            try:
                person_rows = read_table(
                    connection, guard, table_name="ADRP",
                    fields=["PERSNUMBER", "NAME_FIRST", "NAME_LAST", "NAME_TEXT"],
                    options=_like_options_name(norm_query), rowcount=_MAX_RESULTS * 3,
                )
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(norm_query, "ADRP_AUTHORIZATION_ERROR", "Sem autorização para consultar ADRP.", details=format_exception(exc))
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
                    "search_type": search_type, "count": 0, "users": [], "system": target_env, "client": params["client"],
                }

            try:
                link_rows = read_table(
                    connection, guard, table_name="USR21",
                    fields=["BNAME", "PERSNUMBER"],
                    options=make_option_in("PERSNUMBER", list(by_pers.keys())), rowcount=0,
                )
            except Exception as exc:
                if is_authorization_error(exc):
                    return _error_result(norm_query, "USR21_AUTHORIZATION_ERROR", "Sem autorização para consultar USR21.", details=format_exception(exc))
                error_type, message = classify_rfc_error(exc)
                return _error_result(norm_query, f"USR21_{error_type}", message, details=format_exception(exc))

            seen = set()
            for bname, persnumber in link_rows:
                bname = bname.strip()
                pn = persnumber.strip()
                if not bname or bname in seen or pn not in by_pers:
                    continue
                seen.add(bname)
                users.append({"username": bname, "full_name": by_pers[pn]})

        users.sort(key=lambda u: (str(u.get("full_name") or "").upper(), str(u.get("username") or "")))
        truncated = len(users) > _MAX_RESULTS
        users = users[:_MAX_RESULTS]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK" if users else "NAO_ENCONTRADO",
            "query": norm_query,
            "search_type": search_type,
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


def search_users_by_name(query: str) -> dict[str, Any]:
    """Compatibilidade com chamadas legadas."""
    return search_users(query, search_type="name")


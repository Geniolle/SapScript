"""Read-only: dado um utilizador SAP, devolve TODAS as funcoes PFCG que lhe
estao atribuidas em SAP PRD (sem filtro Z*, ao contrario dos fluxos
transacao/objeto -> funcoes).

Inverso de `pfcg_role_users_service` (role -> utilizadores). Consulta `AGR_USERS`
filtrando por `UNAME`. Enriquece com `AGR_TEXTS` (descricao) e classifica a
atribuicao (ATIVO / FUTURO / EXPIRADO) pelas datas.
"""

from __future__ import annotations

import os
import re
from datetime import date
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params,
    choose_best_text,
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
)

ALLOWED_TABLES = ("AGR_USERS", "AGR_TEXTS")

_USER_RE = re.compile(r"^[A-Za-z0-9_.\-]{1,12}$")
_STATUS_RANK = {"ATIVO": 0, "FUTURO": 1, "EXPIRADO": 2}


def validate_username(raw: str) -> str:
    value = str(raw or "").strip().upper()
    if not value:
        raise ValueError("Indique um utilizador SAP.")
    if not _USER_RE.match(value):
        raise ValueError("Utilizador invalido: use apenas letras, numeros, _, . ou - (max. 12).")
    return value


def _classify_assignment_status(from_dat: str, to_dat: str, today: date) -> str:
    today_int = today.year * 10000 + today.month * 100 + today.day
    try:
        from_int = int(from_dat) if from_dat.strip().isdigit() else 0
    except ValueError:
        from_int = 0
    try:
        to_int = int(to_dat) if to_dat.strip().isdigit() else 99991231
    except ValueError:
        to_int = 99991231
    if today_int < from_int:
        return "FUTURO"
    if today_int > to_int:
        return "EXPIRADO"
    return "ATIVO"


def _fmt_date(raw: str) -> str | None:
    raw = str(raw or "").strip()
    if len(raw) == 8 and raw.isdigit() and raw != "00000000":
        return f"{raw[6:8]}/{raw[4:6]}/{raw[0:4]}"
    return None


def _error_result(username: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "username": username,
        "error_type": error_type,
        "message": message,
        "system": SYSTEM_NAME,
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _fetch_role_assignments(connection: Any, guard: Any, username: str) -> dict[str, dict[str, Any]]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_USERS",
        fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT"],
        options=make_option_eq("UNAME", username),
        rowcount=0,
    )
    today = date.today()
    by_role: dict[str, dict[str, Any]] = {}
    for agr_name, _uname, from_dat, to_dat in rows:
        role = agr_name.strip()
        if not role:
            continue
        status = _classify_assignment_status(from_dat, to_dat, today)
        entry = {
            "role": role,
            "valid_from": _fmt_date(from_dat),
            "valid_to": _fmt_date(to_dat),
            "assignment_status": status,
        }
        prev = by_role.get(role)
        if prev is None or _STATUS_RANK.get(status, 9) < _STATUS_RANK.get(prev["assignment_status"], 9):
            by_role[role] = entry
    return by_role


def _fetch_role_descriptions(connection: Any, guard: Any, roles: list[str]) -> dict[str, str]:
    if not roles:
        return {}
    rows = read_table(
        connection,
        guard,
        table_name="AGR_TEXTS",
        fields=["AGR_NAME", "SPRAS", "TEXT"],
        options=make_option_in("AGR_NAME", roles),
        rowcount=0,
    )
    by_role: dict[str, list[dict[str, str]]] = {}
    for agr_name, spras, text in rows:
        by_role.setdefault(agr_name.strip(), []).append({"SPRAS": spras, "TEXT": text})
    out: dict[str, str] = {}
    for role, candidates in by_role.items():
        best = choose_best_text(candidates)
        if best and str(best.get("TEXT", "")).strip():
            out[role] = str(best["TEXT"]).strip()
    return out


def analyze_user_roles_prd(username: str) -> dict[str, Any]:
    try:
        norm_user = validate_username(username)
    except ValueError as exc:
        return _error_result(str(username or "").strip().upper(), "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params()
    except Exception as exc:
        return _error_result(norm_user, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(norm_user, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(norm_user, error_type, message, details=format_exception(exc))

    try:
        try:
            assignments = _fetch_role_assignments(connection, guard, norm_user)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    norm_user,
                    "AGR_USERS_AUTHORIZATION_ERROR",
                    "Sem autorizacao para consultar AGR_USERS.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_user, f"AGR_USERS_{error_type}", message, details=format_exception(exc))

        roles = sorted(assignments.keys())
        warnings: list[str] = []

        if not roles:
            return {
                "ok": True,
                "status": "NAO_ENCONTRADO",
                "username": norm_user,
                "count": 0,
                "roles": [],
                "system": SYSTEM_NAME,
                "client": params["client"],
            }

        descriptions: dict[str, str] = {}
        try:
            descriptions = _fetch_role_descriptions(connection, guard, roles)
        except Exception as exc:
            warnings.append(
                "Sem autorizacao para consultar AGR_TEXTS."
                if is_authorization_error(exc)
                else f"Sem descricoes das funcoes: {classify_rfc_error(exc)[1]}"
            )

        role_rows = [
            {
                "role": role,
                "description": descriptions.get(role),
                "valid_from": assignments[role]["valid_from"],
                "valid_to": assignments[role]["valid_to"],
                "assignment_status": assignments[role]["assignment_status"],
            }
            for role in roles
        ]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "username": norm_user,
            "count": len(role_rows),
            "roles": role_rows,
            "system": SYSTEM_NAME,
            "client": params["client"],
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

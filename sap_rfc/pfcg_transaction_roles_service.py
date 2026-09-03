"""Read-only: dado um codigo de transacao, devolve as funcoes PFCG (roles) a que
essa transacao esta atribuida em SAP PRD.

E o inverso de `pfcg_role_transactions_service` (role -> transacoes). Consulta a
mesma tabela padrao `AGR_TCODES` filtrando por `TCODE`, e ainda:
  - `AGR_TEXTS`  -> descricao das funcoes encontradas;
  - `AGR_AGRS`   -> funcoes compostas (Sammelrolle) que incluem as encontradas.
"""

from __future__ import annotations

import os
import re
from typing import Any

from sap_rfc._rfc_common import (
    SYSTEM_NAME,
    build_connection_params,
    build_connection_params_for,
    resolve_target_env,
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

ALLOWED_TABLES = ("AGR_TCODES", "AGR_TEXTS", "AGR_AGRS", "TSTCT")

_TCODE_RE = re.compile(r"^[A-Za-z0-9_/$+.\-]{1,40}$")


def validate_tcode(raw: str) -> str:
    value = str(raw or "").strip().upper()
    if not value:
        raise ValueError("Indique um codigo de transacao.")
    if not _TCODE_RE.match(value):
        raise ValueError("Codigo de transacao invalido: use apenas letras, numeros, _, /, -, +, . ou $.")
    return value


def _error_result(tcode: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "tcode": tcode,
        "error_type": error_type,
        "message": message,
        "system": resolve_target_env(),
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def _is_custom_role(name: str) -> bool:
    """Apenas funcoes personalizadas (comecam por Z)."""
    return name.strip().upper().startswith("Z")


def _fetch_roles_with_tcode(connection: Any, guard: Any, tcode: str) -> list[str]:
    rows = read_table(
        connection,
        guard,
        table_name="AGR_TCODES",
        fields=["AGR_NAME", "TCODE"],
        options=make_option_eq("TCODE", tcode),
        rowcount=0,
    )
    return sorted(
        {row[0].strip() for row in rows if row[0].strip() and _is_custom_role(row[0])}
    )


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


def _fetch_parent_composites(connection: Any, guard: Any, child_roles: list[str]) -> dict[str, list[str]]:
    """child role -> lista de funcoes compostas que a incluem."""
    if not child_roles:
        return {}
    rows = read_table(
        connection,
        guard,
        table_name="AGR_AGRS",
        fields=["AGR_NAME", "CHILD_AGR"],
        options=make_option_in("CHILD_AGR", child_roles),
        rowcount=0,
    )
    out: dict[str, list[str]] = {}
    for parent, child in rows:
        parent, child = parent.strip(), child.strip()
        if parent and child and _is_custom_role(parent):
            out.setdefault(child, []).append(parent)
    return {k: sorted(set(v)) for k, v in out.items()}


def _fetch_tcode_description(connection: Any, guard: Any, tcode: str) -> str | None:
    rows = read_table(
        connection,
        guard,
        table_name="TSTCT",
        fields=["SPRSL", "TCODE", "TTEXT"],
        options=make_option_eq("TCODE", tcode),
        rowcount=20,
    )
    candidates = [{"SPRSL": s, "TTEXT": t} for s, _tc, t in rows]
    best = choose_best_text(candidates)
    if best and str(best.get("TTEXT", "")).strip():
        return str(best["TTEXT"]).strip()
    return None


def analyze_transaction_roles_prd(tcode: str) -> dict[str, Any]:
    try:
        norm_tcode = validate_tcode(tcode)
    except ValueError as exc:
        return _error_result(str(tcode or "").strip().upper(), "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        target_env = resolve_target_env()
        params = build_connection_params_for(target_env)
    except Exception as exc:
        return _error_result(norm_tcode, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(norm_tcode, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(norm_tcode, error_type, message, details=format_exception(exc))

    try:
        try:
            roles = _fetch_roles_with_tcode(connection, guard, norm_tcode)
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    norm_tcode,
                    "AGR_TCODES_AUTHORIZATION_ERROR",
                    "Sem autorizacao para consultar AGR_TCODES.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_tcode, f"AGR_TCODES_{error_type}", message, details=format_exception(exc))

        tcode_description = None
        warnings: list[str] = []
        try:
            tcode_description = _fetch_tcode_description(connection, guard, norm_tcode)
        except Exception as exc:
            warnings.append(
                "Sem autorizacao para consultar TSTCT."
                if is_authorization_error(exc)
                else f"Sem descricao da transacao: {classify_rfc_error(exc)[1]}"
            )

        if not roles:
            return {
                "ok": True,
                "status": "NAO_ENCONTRADO",
                "tcode": norm_tcode,
                "tcode_description": tcode_description,
                "count": 0,
                "roles": [],
                "system": target_env,
                "client": params["client"],
                **({"warning": " ".join(warnings)} if warnings else {}),
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

        parents: dict[str, list[str]] = {}
        try:
            parents = _fetch_parent_composites(connection, guard, roles)
        except Exception as exc:
            warnings.append(
                "Sem autorizacao para consultar AGR_AGRS."
                if is_authorization_error(exc)
                else f"Sem funcoes compostas: {classify_rfc_error(exc)[1]}"
            )

        role_rows = [
            {
                "role": role,
                "description": descriptions.get(role),
                "composite_parents": parents.get(role, []),
            }
            for role in roles
        ]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "tcode": norm_tcode,
            "tcode_description": tcode_description,
            "count": len(role_rows),
            "roles": role_rows,
            "system": target_env,
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

"""Read-only: pesquisa funcoes/perfis PFCG (AGR_DEFINE) por padrao de nome
(curinga '*') no ambiente pedido. Devolve nome + descricao (AGR_TEXTS).

Uso tipico: utilizador nao lembra o nome exato ("Z_IT_EQUIPA_INTERNA"),
escreve "Z*EQUIPA*" e recebe todas as funcoes cujo AGR_NAME bate com o padrao.
"""
from __future__ import annotations

import os
import re
from typing import Any

from sap_rfc._rfc_common import (
    build_connection_params_for,
    choose_best_text,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    is_authorization_error,
    load_project_env,
    make_read_only_guard,
    read_table,
    resolve_target_env,
)

ALLOWED_TABLES = ("AGR_DEFINE", "AGR_TEXTS")

_MAX_RESULTS = 200
_MAX_PATTERN = 40
_PATTERN_RE = re.compile(r"^[A-Z0-9_\-:/*]+$")


def sanitize_role_pattern(raw: str) -> str:
    value = str(raw or "").strip().upper()
    if not value:
        raise ValueError("Indique um padrão de pesquisa (ex.: Z*EQUIPA*).")
    value = value[:_MAX_PATTERN]
    if not _PATTERN_RE.fullmatch(value):
        raise ValueError("Padrão inválido: use apenas A-Z, 0-9, _, -, /, : e * (curinga).")
    if len(value.replace("*", "")) < 2:
        raise ValueError("Indique pelo menos 2 caracteres fora dos curingas (*).")
    return value


def _like_option(pattern: str) -> list[dict[str, str]]:
    # '_' e caracter valido em nomes de funcao (ex.: Z_IT_EQUIPA_INTERNA), nao e
    # curinga aqui — escapa-o para o LIKE nao o tratar como "um caracter qualquer".
    # So '*' e curinga (-> '%' do LIKE).
    escaped = pattern.replace("_", "@_").replace("*", "%")
    return [{"TEXT": f"AGR_NAME LIKE '{escaped}' ESCAPE '@'"}]


def _error_result(pattern: str, error_type: str, message: str, *, details: str | None = None) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": False,
        "status": "ERRO",
        "pattern": pattern,
        "error_type": error_type,
        "message": message,
        "system": resolve_target_env(),
        "client": os.getenv("SAP_PRD_CLIENT", "").strip() or None,
    }
    if details:
        payload["details"] = details
    return payload


def search_pfcg_roles(pattern: str) -> dict[str, Any]:
    raw_pattern = str(pattern or "").strip().upper()

    try:
        norm_pattern = sanitize_role_pattern(pattern)
    except ValueError as exc:
        return _error_result(raw_pattern, "INVALID_INPUT", str(exc))

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        target_env = resolve_target_env()
        params = build_connection_params_for(target_env)
    except Exception as exc:
        return _error_result(norm_pattern, "CONFIG_ERROR", str(exc), details=format_exception(exc))

    try:
        from pyrfc import Connection  # type: ignore
    except Exception as exc:
        error_type, message = classify_import_error(exc)
        return _error_result(norm_pattern, error_type, message, details=format_exception(exc))

    guard = make_read_only_guard(ALLOWED_TABLES)
    connection = None
    try:
        connection = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        connection.call("RFC_PING")
    except Exception as exc:
        error_type, message = classify_rfc_error(exc)
        return _error_result(norm_pattern, error_type, message, details=format_exception(exc))

    try:
        try:
            define_rows = read_table(
                connection,
                guard,
                table_name="AGR_DEFINE",
                fields=["AGR_NAME"],
                options=_like_option(norm_pattern),
                rowcount=_MAX_RESULTS + 1,
            )
        except Exception as exc:
            if is_authorization_error(exc):
                return _error_result(
                    norm_pattern,
                    "AGR_DEFINE_AUTHORIZATION_ERROR",
                    "Não foi possível pesquisar: sem autorização para consultar AGR_DEFINE.",
                    details=format_exception(exc),
                )
            error_type, message = classify_rfc_error(exc)
            return _error_result(norm_pattern, f"AGR_DEFINE_{error_type}", message, details=format_exception(exc))

        role_names = sorted({row[0].strip() for row in define_rows if row and row[0].strip()})
        truncated = len(role_names) > _MAX_RESULTS
        role_names = role_names[:_MAX_RESULTS]

        if not role_names:
            return {
                "ok": True,
                "status": "NAO_ENCONTRADO",
                "pattern": norm_pattern,
                "count": 0,
                "roles": [],
                "system": target_env,
                "client": params["client"],
            }

        descriptions: dict[str, str] = {}
        warning = None
        try:
            text_rows_raw = read_table(
                connection,
                guard,
                table_name="AGR_TEXTS",
                fields=["AGR_NAME", "SPRAS", "TEXT"],
                options=_like_option(norm_pattern),
                rowcount=0,
            )
            by_role: dict[str, list[dict[str, str]]] = {}
            for agr_name, spras, text in text_rows_raw:
                by_role.setdefault(agr_name.strip(), []).append({"SPRAS": spras, "TEXT": text})
            for role in role_names:
                best = choose_best_text(by_role.get(role, []))
                if best and str(best.get("TEXT", "")).strip():
                    descriptions[role] = str(best["TEXT"]).strip()
        except Exception as exc:
            if is_authorization_error(exc):
                warning = "Sem autorização para consultar AGR_TEXTS (descrições omitidas)."
            else:
                warning = classify_rfc_error(exc)[1]

        roles = [{"role": role, "description": descriptions.get(role)} for role in role_names]

        payload: dict[str, Any] = {
            "ok": True,
            "status": "OK",
            "pattern": norm_pattern,
            "count": len(roles),
            "roles": roles,
            "system": target_env,
            "client": params["client"],
        }
        if truncated:
            payload["warning"] = (
                f"Mais de {_MAX_RESULTS} resultados; mostro apenas os primeiros {_MAX_RESULTS}. Refine o padrão."
            )
        if warning:
            payload["warning"] = ((payload.get("warning", "") + " ") + warning).strip()
        return payload
    finally:
        try:
            if connection is not None:
                connection.close()
        except Exception:
            pass

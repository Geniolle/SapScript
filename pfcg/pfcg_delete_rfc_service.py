from __future__ import annotations

from pfcg.pfcg_create_rfc_service import _run_bridge_cli


def preview_pfcg_role_delete_rfc(
    environment: str,
    role_name: str,
    transport_mode: str = "LOCAL",
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    args = [
        "--environment",
        str(environment or "").strip().upper(),
        "--role-name",
        str(role_name or "").strip(),
        "--transport-mode",
        str(transport_mode or "LOCAL").strip().upper(),
        "--request",
        str(request_number or "").strip(),
        "--request-description",
        str(request_description or ""),
    ]
    return _run_bridge_cli("sap_rfc.pfcg_role_delete_preview_cli", args)


def delete_pfcg_role_rfc(
    environment: str,
    role_name: str,
    transport_mode: str = "LOCAL",
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    args = [
        "--environment",
        str(environment or "").strip().upper(),
        "--role-name",
        str(role_name or "").strip(),
        "--transport-mode",
        str(transport_mode or "LOCAL").strip().upper(),
        "--request",
        str(request_number or "").strip(),
        "--request-description",
        str(request_description or ""),
        "--confirm",
    ]
    return _run_bridge_cli("sap_rfc.pfcg_role_delete_cli", args)


def _bulk_role_name_args(role_names: list[str]) -> list[str]:
    args: list[str] = []
    for role_name in role_names or []:
        args.extend(["--role-name", str(role_name or "").strip()])
    return args


def preview_pfcg_bulk_role_delete_rfc(
    environment: str,
    role_names: list[str],
    transport_mode: str = "LOCAL",
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    args = [
        "--environment",
        str(environment or "").strip().upper(),
        *_bulk_role_name_args(role_names),
        "--transport-mode",
        str(transport_mode or "LOCAL").strip().upper(),
        "--request",
        str(request_number or "").strip(),
        "--request-description",
        str(request_description or ""),
    ]
    timeout = max(120, 10 * len(role_names or []) + 60)
    return _run_bridge_cli("sap_rfc.pfcg_role_bulk_delete_preview_cli", args, timeout=timeout)


def bulk_delete_pfcg_roles_rfc(
    environment: str,
    role_names: list[str],
    transport_mode: str = "LOCAL",
    request_number: str = "",
    request_description: str = "",
) -> dict[str, Any]:
    args = [
        "--environment",
        str(environment or "").strip().upper(),
        *_bulk_role_name_args(role_names),
        "--transport-mode",
        str(transport_mode or "LOCAL").strip().upper(),
        "--request",
        str(request_number or "").strip(),
        "--request-description",
        str(request_description or ""),
        "--confirm",
    ]
    timeout = max(120, 10 * len(role_names or []) + 60)
    return _run_bridge_cli("sap_rfc.pfcg_role_bulk_delete_cli", args, timeout=timeout)

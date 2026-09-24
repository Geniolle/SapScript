"""Read-only: apoio à investigação de item de liquidez / fluxo de caixa (New Cash
Management, S/4HANA) para um documento financeiro.

Contexto: o campo clássico SKB1-FIPOS não está em uso neste sistema. A
classificação real ("Item de Liquidez") vem do New Cash Management (FQM) e é
derivada em runtime (views HANA como FQM_BSEG_FLOW_TYPE / FQMS_LIQUPOS_DERIV),
não é lida por RFC_READ_TABLE. Este serviço não calcula o item de liquidez —
apoia a investigação manual trazendo o documento, a respetiva cadeia de
compensação (AUGBL) e a descrição de códigos de item de liquidez já
conhecidos (tabela ALIQITEMTXT).

Ver docs/CONHECIMENTO_FLUXO_DE_CAIXA_LIQUIDEZ.md para o contexto completo.

Estritamente READ-ONLY: apenas RFC_PING e RFC_READ_TABLE. Nunca escreve,
nunca compensa, nunca reverte documentos.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    _PYRFC_IMPORT_ERROR = exc
else:
    _PYRFC_IMPORT_ERROR = None

from sap_rfc._rfc_common import (
    build_connection_params_for_env,
    classify_import_error,
    classify_rfc_error,
    find_project_root,
    format_exception,
    load_project_env,
    make_read_only_guard,
    resolve_target_env,
)

CASH_FLOW_ALLOWED_TABLES = ("BKPF", "BSEG", "SKAT", "SKB1", "ALIQITEMTXT")

DOCUMENT_HEADER_FIELDS = ("BUKRS", "BELNR", "GJAHR", "BLART", "BLDAT", "BUDAT", "BKTXT", "TCODE", "USNAM")
DOCUMENT_LINE_FIELDS = (
    "BUKRS", "BELNR", "GJAHR", "BUZEI", "HKONT", "SHKZG", "DMBTR", "BSCHL",
    "SGTXT", "KOART", "AUGBL", "AUGDT",
)


def _is_windows_runtime() -> bool:
    return os.name == "nt" or sys.platform.startswith("win")


def _bridge_python_executable() -> str | None:
    candidates = [os.getenv("SAP_FI_BRIDGE_PYTHON", "").strip()]
    if _is_windows_runtime():
        candidates.append(str((Path(__file__).resolve().parents[1] / ".venv-rfc" / "Scripts" / "python.exe").resolve()))
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None


def _is_no_data_error(exc: BaseException) -> bool:
    text = f"{exc.__class__.__name__} {getattr(exc, 'key', '')} {getattr(exc, 'message', '')} {exc}".upper()
    return "TABLE_WITHOUT_DATA" in text or "RFC_TABLE_WITHOUT_DATA" in text


def _read_table_core(environment: str, table: str, options: list[str], fields: list[str], rowcount: int) -> list[dict[str, Any]]:
    project_root = find_project_root()
    load_project_env(project_root)

    from sap_agent.safety import SafetyGuard

    guard = make_read_only_guard(CASH_FLOW_ALLOWED_TABLES)
    guard.assert_function_allowed("RFC_READ_TABLE")
    guard.assert_table_allowed(table)

    params = build_connection_params_for_env(environment)
    conn = None
    try:
        conn = Connection(**params)
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": f} for f in fields],
            OPTIONS=[{"TEXT": o} for o in options],
            ROWCOUNT=rowcount,
        )
        rows: list[dict[str, Any]] = []
        for item in result.get("DATA", []) or []:
            parts = [p.strip() for p in str(item.get("WA", "")).split("|")]
            rows.append({field: (parts[i] if i < len(parts) else "") for i, field in enumerate(fields)})
        return rows
    except Exception as exc:
        if _is_no_data_error(exc):
            return []
        code, msg = classify_rfc_error(exc)
        raise RuntimeError(f"{msg} ({code}): {format_exception(exc)}") from exc
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def _read_table_via_bridge(environment: str, table: str, options: list[str], fields: list[str], rowcount: int) -> list[dict[str, Any]]:
    python_exe = _bridge_python_executable()
    if not python_exe:
        raise RuntimeError(
            "Leitura via bridge indisponível neste runtime. "
            "Configure SAP_FI_BRIDGE_PYTHON com um Python que tenha PyRFC instalado."
        )

    payload = {"environment": environment, "table": table, "options": options, "fields": fields, "rowcount": rowcount}
    bridge_code = (
        "import json, os, sys\n"
        "from pathlib import Path\n"
        "project_dir = os.environ.get('SAP_SCRIPT_PROJECT_DIR', '').strip()\n"
        "if project_dir and project_dir not in sys.path:\n"
        "    sys.path.insert(0, project_dir)\n"
        "payload = json.loads(sys.stdin.read() or '{}')\n"
        "from sap_rfc.cash_flow_validation_service import _read_table_core\n"
        "try:\n"
        "    rows = _read_table_core(payload['environment'], payload['table'], payload['options'], payload['fields'], payload['rowcount'])\n"
        "    print(json.dumps({'ok': True, 'rows': rows}, ensure_ascii=False))\n"
        "except Exception as exc:\n"
        "    print(json.dumps({'ok': False, 'message': str(exc)}, ensure_ascii=False))\n"
        "    raise SystemExit(1)\n"
    )
    proc = subprocess.run(
        [python_exe, "-c", bridge_code],
        input=json.dumps(payload, ensure_ascii=False),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
        cwd=str(Path(__file__).resolve().parent.parent),
    )
    stdout = (proc.stdout or "").strip()
    if not stdout:
        raise RuntimeError(f"Bridge devolveu saída vazia.\nSTDERR: {proc.stderr}")
    try:
        data = json.loads(stdout.splitlines()[-1])
    except Exception as exc:
        raise RuntimeError(f"Bridge devolveu JSON inválido.\nSTDOUT: {proc.stdout}\nSTDERR: {proc.stderr}") from exc
    if not data.get("ok"):
        raise RuntimeError(str(data.get("message") or "Falha no bridge."))
    rows = data.get("rows")
    if not isinstance(rows, list):
        raise RuntimeError("Bridge devolveu linhas inválidas.")
    return [row for row in rows if isinstance(row, dict)]


def read_cash_flow_table(environment: str, table: str, options: list[str], fields: list[str], rowcount: int = 200) -> list[dict[str, Any]]:
    try:
        return _read_table_core(environment, table, options, fields, rowcount)
    except Exception as exc:
        if Connection is None:
            return _read_table_via_bridge(environment, table, options, fields, rowcount)
        raise RuntimeError(str(exc)) from exc


def _account_description(environment: str, saknr: str) -> str:
    if not saknr:
        return ""
    rows = read_cash_flow_table(
        environment, "SKAT", [f"SAKNR = '{saknr}' AND SPRAS = 'P'"], ["SAKNR", "TXT50"], rowcount=1
    )
    return rows[0]["TXT50"] if rows else ""


def fetch_document(environment: str, bukrs: str, belnr: str, gjahr: str) -> dict[str, Any]:
    """Cabeçalho + itens do documento, com a descrição da conta de cada item."""
    bukrs, belnr, gjahr = bukrs.strip().upper(), belnr.strip().zfill(10), gjahr.strip()

    header_rows = read_cash_flow_table(
        environment,
        "BKPF",
        [f"BUKRS = '{bukrs}' AND BELNR = '{belnr}' AND GJAHR = '{gjahr}'"],
        list(DOCUMENT_HEADER_FIELDS),
        rowcount=1,
    )
    if not header_rows:
        return {"found": False, "bukrs": bukrs, "belnr": belnr, "gjahr": gjahr}

    line_rows = read_cash_flow_table(
        environment,
        "BSEG",
        [f"BUKRS = '{bukrs}' AND BELNR = '{belnr}' AND GJAHR = '{gjahr}'"],
        list(DOCUMENT_LINE_FIELDS),
        rowcount=200,
    )
    for line in line_rows:
        line["HKONT_DESC"] = _account_description(environment, line.get("HKONT", ""))

    return {"found": True, "header": header_rows[0], "lines": line_rows}


def fetch_clearing_group(environment: str, bukrs: str, augbl: str, gjahr: str) -> list[dict[str, Any]]:
    """Todos os itens BSEG (qualquer documento) compensados pelo mesmo AUGBL.

    Útil para "subir até a origem": mostra, na mesma conta ou noutras, todos
    os documentos que a compensação técnica juntou/fechou.
    """
    bukrs, augbl, gjahr = bukrs.strip().upper(), augbl.strip().zfill(10), gjahr.strip()
    rows = read_cash_flow_table(
        environment,
        "BSEG",
        [f"BUKRS = '{bukrs}' AND AUGBL = '{augbl}' AND GJAHR = '{gjahr}'"],
        list(DOCUMENT_LINE_FIELDS),
        rowcount=200,
    )
    for row in rows:
        row["HKONT_DESC"] = _account_description(environment, row.get("HKONT", ""))
    return rows


def lookup_liquidity_item(environment: str, code: str) -> dict[str, Any] | None:
    """Descrição de um código de Item de Liquidez (New Cash Management), ex.: LQAO0141."""
    code = code.strip().upper()
    if not code:
        return None
    rows = read_cash_flow_table(
        environment,
        "ALIQITEMTXT",
        [f"LANGUAGE = 'P' AND LIQUIDITYITEM = '{code}'"],
        ["LIQUIDITYITEM", "LIQUIDITYITEMNAME", "LANGUAGE"],
        rowcount=1,
    )
    return rows[0] if rows else None

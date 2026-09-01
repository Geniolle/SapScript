from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params_for_env, find_project_root, load_project_env
from sap_rfc.fi_document_service import _bridge_python_executable
from tests.manual.abap_source_rfc import collect_source_hits, resolve_abap_roots

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    PYRFC_IMPORT_ERROR = exc
else:
    PYRFC_IMPORT_ERROR = None


TARGET_OBJECT_NAME = os.getenv("SAP_ABAP_TARGET_NAME", "ZFI_EXTERNAL_PAYMENTS").strip() or "ZFI_EXTERNAL_PAYMENTS"
TARGET_OBJECT_KIND = os.getenv("SAP_ABAP_TARGET_KIND", "PROG").strip().upper() or "PROG"
MARKERS = (
    "BAPI_ACC_DOCUMENT_POST",
    "BAPI_DOCUMENT_POST",
    "BAPI_ACC_DOCUMENT_CHECK",
    "ACCOUNTTAX",
    "BAPIACTX09",
    "BAPIACCR09",
    "ACCOUNTWT",
    "CALCULATE_TAX_FROM_NET_AMOUNT",
    "CALCULATE_TAX_FROM_GROSSAMOUNT",
    "CALL FUNCTION",
)


def _scan_roots(connection: object, roots: list[str]) -> list[dict[str, object]]:
    hits: list[dict[str, object]] = []
    visited: set[str] = set()
    for root in roots:
        for hit in collect_source_hits(
            connection,
            root,
            MARKERS,
            language=os.getenv("SAP_QAD_LANG", "PT").strip() or "PT",
            visited=visited,
        ):
            hits.append(
                {
                    "root_name": hit.root_name,
                    "program": hit.program_name,
                    "include": hit.include_name,
                    "marker": hit.marker,
                    "line_number": hit.line_number,
                    "line_text": hit.line_text,
                }
            )
    return hits


def main() -> int:
    if Connection is None:
        bridge_python = _bridge_python_executable()
        if bridge_python is None:
            raise RuntimeError(f"PyRFC indisponível: {PYRFC_IMPORT_ERROR}")
        proc = subprocess.run(
            [str(bridge_python), str(Path(__file__).resolve())],
            cwd=str(REPO_ROOT),
            capture_output=True,
            text=True,
        )
        if proc.stdout:
            print(proc.stdout, end="")
        if proc.stderr:
            print(proc.stderr, end="", file=sys.stderr)
        return proc.returncode

    project_root = find_project_root()
    load_project_env(project_root)

    connection = Connection(**build_connection_params_for_env("QAD"))  # type: ignore[misc]
    try:
        roots, metadata_probe = resolve_abap_roots(
            connection,
            TARGET_OBJECT_NAME,
            object_kind=TARGET_OBJECT_KIND,
            language=os.getenv("SAP_QAD_LANG", "PT").strip() or "PT",
        )
        hits = _scan_roots(connection, roots)
        print(
            json.dumps(
                {
                    "target_name": TARGET_OBJECT_NAME,
                    "target_kind": TARGET_OBJECT_KIND,
                    "resolved_roots": roots,
                    "class_metadata": None
                    if metadata_probe is None
                    else {
                        "class_name": metadata_probe.class_name,
                        "function_name": metadata_probe.function_name,
                        "request_parameters": metadata_probe.request_parameters,
                        "errors": metadata_probe.errors,
                        "response_keys": sorted(metadata_probe.response.keys()),
                    },
                    "hit_count": len(hits),
                    "hits": hits,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    finally:
        try:
            connection.close()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())

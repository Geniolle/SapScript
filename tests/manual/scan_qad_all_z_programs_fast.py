from __future__ import annotations

import json
import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params_for_env, find_project_root, load_project_env
from sap_rfc.fi_document_service import _bridge_python_executable

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    PYRFC_IMPORT_ERROR = exc
else:
    PYRFC_IMPORT_ERROR = None


MARKERS = ("ACCOUNTTAX", "BAPI_ACC_DOCUMENT_POST", "BAPI_DOCUMENT_POST")


def _programs(connection: object) -> list[str]:
    response = connection.call(  # type: ignore[attr-defined]
        "RFC_READ_TABLE",
        QUERY_TABLE="TADIR",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "OBJ_NAME"}],
        OPTIONS=[
            {"TEXT": "PGMID = 'R3TR'"},
            {"TEXT": "AND OBJECT = 'PROG'"},
            {"TEXT": "AND OBJ_NAME LIKE 'Z%'"},
        ],
        ROWCOUNT=0,
    )
    rows = response.get("DATA") or []
    programs: list[str] = []
    seen: set[str] = set()
    for row in rows:
        wa = str(row.get("WA") or "").strip()
        if not wa:
            continue
        name = wa.split("|", 1)[0].strip().upper()
        if name and name not in seen:
            seen.add(name)
            programs.append(name)
    return programs


def _source_lines(connection: object, program: str) -> list[str]:
    response = connection.call(  # type: ignore[attr-defined]
        "RPY_PROGRAM_READ",
        PROGRAM_NAME=program,
        LANGUAGE=os.getenv("SAP_QAD_LANG", "PT").strip() or "PT",
        ONLY_SOURCE="X",
        READ_LATEST_VERSION="X",
        WITH_LOWERCASE="X",
    )
    source = (response or {}).get("SOURCE") or (response or {}).get("SOURCE_TAB") or []
    if isinstance(source, str):
        return source.splitlines()
    lines: list[str] = []
    for item in source:
        if isinstance(item, dict):
            for key in ("LINE", "TEXT", "SOURCE_LINE", "ABAP"):
                value = item.get(key)
                if value is not None:
                    lines.append(str(value))
                    break
            else:
                lines.append(str(item))
        else:
            lines.append(str(item))
    return lines


def main() -> int:
    if Connection is None:
        bridge_python = _bridge_python_executable()
        if bridge_python is None:
            raise RuntimeError(f"PyRFC indisponível: {PYRFC_IMPORT_ERROR}")
        import subprocess

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
        programs = _programs(connection)
        hits: list[dict[str, object]] = []
        for index, program in enumerate(programs, start=1):
            try:
                lines = _source_lines(connection, program)
            except Exception:
                continue
            for line_number, line_text in enumerate(lines, start=1):
                upper = line_text.upper()
                if any(marker in upper for marker in MARKERS):
                    hits.append(
                        {
                            "program": program,
                            "line_number": line_number,
                            "line_text": line_text.strip(),
                        }
                    )
                    break
        print(json.dumps({"program_count": len(programs), "hit_count": len(hits), "hits": hits}, ensure_ascii=False, indent=2))
        return 0
    finally:
        try:
            connection.close()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())

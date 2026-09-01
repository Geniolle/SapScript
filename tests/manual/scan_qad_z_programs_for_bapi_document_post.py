from __future__ import annotations

import json
import os
import sys
import subprocess
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

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


TARGET_MARKERS = ("BAPI_DOCUMENT_POST", "BAPI_ACC_DOCUMENT_POST")


@dataclass
class Hit:
    program: str
    include: str
    marker: str
    line_number: int
    line_text: str


def _safe_lines(source: Any) -> list[str]:
    if isinstance(source, str):
        return source.splitlines()
    if isinstance(source, list):
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
    return []


def _safe_includes(response: dict[str, Any]) -> list[str]:
    includes: list[str] = []
    for key in ("INCLUDE_TAB", "INCLUDETAB", "INCLUDES"):
        rows = response.get(key) or []
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            if isinstance(row, dict):
                for field in ("INCLUDE", "NAME", "PROGRAM", "PROGNAME"):
                    value = str(row.get(field) or "").strip()
                    if value:
                        includes.append(value)
                        break
            else:
                value = str(row).strip()
                if value:
                    includes.append(value)
    deduped: list[str] = []
    seen: set[str] = set()
    for include in includes:
        if include not in seen:
            seen.add(include)
            deduped.append(include)
    return deduped


def _read_program(connection: Any, program_name: str) -> dict[str, Any]:
    return dict(
        connection.call(
            "RPY_PROGRAM_READ",
            PROGRAM_NAME=program_name,
            LANGUAGE=os.getenv("SAP_QAD_LANG", "PT").strip() or "PT",
            WITH_INCLUDELIST="X",
            ONLY_SOURCE="X",
            READ_LATEST_VERSION="X",
            WITH_LOWERCASE="X",
        )
        or {}
    )


def _collect_hits(
    connection: Any,
    program_name: str,
    *,
    include_name: str | None = None,
    visited: set[str] | None = None,
) -> list[Hit]:
    visited = visited or set()
    normalized = program_name.strip().upper()
    if not normalized or normalized in visited:
        return []
    visited.add(normalized)

    try:
        response = _read_program(connection, normalized)
    except Exception as exc:
        print(f"[WARN] Não foi possível ler {normalized}: {exc}", file=sys.stderr)
        return []

    hits: list[Hit] = []
    source_lines = _safe_lines(response.get("SOURCE") or response.get("SOURCE_TAB"))
    for line_number, line_text in enumerate(source_lines, start=1):
        upper = line_text.upper()
        for marker in TARGET_MARKERS:
            if marker in upper:
                hits.append(
                    Hit(
                        program=normalized,
                        include=include_name or normalized,
                        marker=marker,
                        line_number=line_number,
                        line_text=line_text.strip(),
                    )
                )

    for include in _safe_includes(response):
        hits.extend(_collect_hits(connection, include, include_name=include, visited=visited))
    return hits


def _list_z_programs(connection: Any) -> list[str]:
    rows = connection.call(
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
    data = rows.get("DATA") or []
    programs: list[str] = []
    for row in data:
        wa = str(row.get("WA") or "").strip()
        if not wa:
            continue
        name = wa.split("|", 1)[0].strip().upper()
        if name:
            programs.append(name)

    deduped: list[str] = []
    seen: set[str] = set()
    for program in programs:
        if program not in seen:
            seen.add(program)
            deduped.append(program)
    return deduped


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
        programs = _list_z_programs(connection)
        hits: list[Hit] = []
        for program in programs:
            hits.extend(_collect_hits(connection, program))

        unique_hits = [asdict(hit) for hit in hits]
        print(
            json.dumps(
                {
                    "program_count": len(programs),
                    "hit_count": len(unique_hits),
                    "hits": unique_hits,
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

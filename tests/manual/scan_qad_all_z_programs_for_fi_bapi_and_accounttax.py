from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params_for_env, find_project_root, load_project_env
from sap_rfc.fi_document_service import _bridge_python_executable
from tests.manual.scan_qad_z_programs_for_bapi_document_post import Hit, asdict, _collect_hits

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    PYRFC_IMPORT_ERROR = exc
else:
    PYRFC_IMPORT_ERROR = None


def _programs_from_tadir(connection: object) -> list[str]:
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


def _run(connection: object) -> dict[str, object]:
    programs = _programs_from_tadir(connection)
    hits: list[Hit] = []
    for program in programs:
        hits.extend(_collect_hits(connection, program))
    return {
        "program_count": len(programs),
        "hit_count": len(hits),
        "hits": [asdict(hit) for hit in hits],
    }


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
        print(
            json.dumps(
                _run(connection),
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

from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from sap_rfc._rfc_common import build_connection_params_for, find_project_root, load_project_env  # noqa: E402


def fields_for(conn: Any, table: str) -> list[dict[str, Any]]:
    response = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE="DD03L",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": name} for name in ("FIELDNAME", "POSITION", "LENG", "DATATYPE")],
        OPTIONS=[{"TEXT": f"TABNAME = '{table}' AND AS4LOCAL = 'A'"}],
        ROWCOUNT=0,
    )
    rows = []
    for item in response.get("DATA", []) or []:
        parts = str(item.get("WA", "")).split("|")
        parts += [""] * max(0, 4 - len(parts))
        rows.append(dict(zip(("FIELDNAME", "POSITION", "LENG", "DATATYPE"), (part.strip() for part in parts))))
    return sorted(rows, key=lambda row: int(row.get("POSITION") or 0))


def read_rows(conn: Any, table: str, fields: list[str]) -> list[dict[str, str]]:
    response = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE=table,
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": field} for field in fields],
        OPTIONS=[],
        ROWCOUNT=0,
    )
    rows = []
    for item in response.get("DATA", []) or []:
        parts = str(item.get("WA", "")).split("|")
        parts += [""] * max(0, len(fields) - len(parts))
        rows.append({field: parts[i].strip() for i, field in enumerate(fields)})
    return rows


def main() -> int:
    load_project_env(find_project_root())
    from pyrfc import Connection

    out_dir = ROOT / "output" / f"ZCKP_runtime_config_PRD_{datetime.now():%Y%m%d_%H%M%S}"
    out_dir.mkdir(parents=True, exist_ok=True)
    tables = ["/SBXC/ZCKP_TAB00", "/SBXC/ZCKP_TAB03"]
    report: dict[str, Any] = {"read_only": True, "environment": "PRD", "timestamp": datetime.now().isoformat(timespec="seconds"), "tables": {}, "errors": []}
    conn = Connection(**build_connection_params_for("PRD"))
    try:
        conn.call("RFC_PING")
        for table in tables:
            try:
                # These fields are confirmed by the extracted active source. Avoiding a broad
                # DDIC lookup also prevents a costly full DD03L scan in PRD.
                if table == "/SBXC/ZCKP_TAB00":
                    field_names = ["PROCESSO", "EST_CAB", "EST_LIN", "FM_CAB_DISP", "FM_LIN_DISP", "FM_DBLCLK", "FM_ALTERACAO"]
                else:
                    field_names = ["EST_BOTAO", "ALVHL", "NUM", "FUNCTION", "ICON", "BUTN_TYPE", "DISABLED", "FM", "NUM_PAI"]
                metadata = [{"FIELDNAME": name} for name in field_names]
                rows = read_rows(conn, table, field_names)
                report["tables"][table] = {"fields": metadata, "rows": rows}
            except Exception as exc:
                report["errors"].append({"table": table, "error": repr(exc)})
        (out_dir / "runtime_configuration.json").write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        summary = {table: len(data["rows"]) for table, data in report["tables"].items()}
        print(json.dumps({"run_dir": str(out_dir), "row_counts": summary, "errors": report["errors"]}, ensure_ascii=False, indent=2))
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

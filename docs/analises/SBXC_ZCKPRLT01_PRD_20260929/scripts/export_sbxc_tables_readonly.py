from __future__ import annotations

import csv
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from sap_rfc._rfc_common import build_connection_params_for, find_project_root, load_project_env  # noqa: E402


DELIMITER = "|"
SAMPLE_SIZE = 20
MAX_CHUNK_WIDTH = 450


def safe_name(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value).strip("_") or "EMPTY"


def parse(result: dict[str, Any], fields: list[str]) -> list[dict[str, str]]:
    rows = []
    for item in result.get("DATA", []) or []:
        parts = str(item.get("WA", "")).split(DELIMITER)
        parts += [""] * max(0, len(fields) - len(parts))
        rows.append({field: parts[index].strip() for index, field in enumerate(fields)})
    return rows


def read(conn: Any, table: str, fields: list[str], *, options: list[str] | None = None, rowskips: int = 0, rowcount: int = 0) -> list[dict[str, str]]:
    payload = {
        "QUERY_TABLE": table,
        "DELIMITER": DELIMITER,
        "FIELDS": [{"FIELDNAME": field} for field in fields],
        "OPTIONS": [{"TEXT": value} for value in (options or [])],
        "ROWSKIPS": rowskips,
        "ROWCOUNT": rowcount,
    }
    if rowskips:
        payload["GET_SORTED"] = "X"
    response = conn.call("RFC_READ_TABLE", **payload)
    return parse(dict(response or {}), fields)


def field_metadata(conn: Any, table: str) -> list[dict[str, str]]:
    fields = ["FIELDNAME", "POSITION", "LENG", "DATATYPE", "KEYFLAG"]
    rows = read(conn, "DD03L", fields, options=[f"TABNAME = '{table}'", "AND AS4LOCAL = 'A'"])
    usable = [row for row in rows if row["FIELDNAME"] and row["FIELDNAME"] != ".INCLUDE"]
    return sorted(usable, key=lambda row: int(row["POSITION"] or 0))


def chunks(metadata: list[dict[str, str]]) -> list[list[str]]:
    result: list[list[str]] = []
    current: list[str] = []
    width = 0
    for item in metadata:
        name = item["FIELDNAME"]
        length = max(1, int(item.get("LENG") or 1))
        if current and width + length + 1 > MAX_CHUNK_WIDTH:
            result.append(current)
            current, width = [], 0
        current.append(name)
        width += length + 1
    if current:
        result.append(current)
    return result


def export_chunk(conn: Any, table: str, fields: list[str], path: Path) -> int:
    rows = read(conn, table, fields, rowskips=0, rowcount=SAMPLE_SIZE)
    with path.open("w", encoding="utf-8") as handle:
        for row in rows:
            handle.write(json.dumps(row, ensure_ascii=False) + "\n")
    return len(rows)


def main() -> int:
    load_project_env(find_project_root())
    from pyrfc import Connection

    started = datetime.now()
    out = ROOT / "output" / f"SBXC_tables_PRD_{started:%Y%m%d_%H%M%S}"
    data_dir = out / "tables"
    data_dir.mkdir(parents=True, exist_ok=True)
    manifest: dict[str, Any] = {
        "environment": "PRD",
        "read_only": True,
        "started_at": started.isoformat(timespec="seconds"),
        "namespace": "/SBXC/%",
        "sample_size_per_table": SAMPLE_SIZE,
        "tables": {},
        "errors": [],
    }
    conn = Connection(**build_connection_params_for("PRD"))
    try:
        conn.call("RFC_PING")
        inventory = read(
            conn,
            "DD02L",
            ["TABNAME", "TABCLASS", "AS4LOCAL", "AS4VERS"],
            options=["TABNAME LIKE '/SBXC/%'", "AND AS4LOCAL = 'A'", "AND TABCLASS = 'TRANSP'"],
        )
        inventory = sorted(
            {row["TABNAME"]: row for row in inventory if row["TABNAME"]}.values(),
            key=lambda row: (row["TABNAME"] in {"/SBXC/ENVIA_GM", "/SBXC/ENVIA_PED"}, row["TABNAME"]),
        )
        with (out / "inventory.csv").open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.DictWriter(handle, fieldnames=["TABNAME", "TABCLASS", "AS4LOCAL", "AS4VERS"])
            writer.writeheader()
            writer.writerows(inventory)

        for index, table_info in enumerate(inventory, start=1):
            table = table_info["TABNAME"]
            entry: dict[str, Any] = {"catalog": table_info, "status": "pending", "chunks": []}
            manifest["tables"][table] = entry
            try:
                metadata = field_metadata(conn, table)
                entry["fields"] = metadata
                groups = chunks(metadata)
                if not groups:
                    raise RuntimeError("No active elementary fields returned from DD03L")
                row_counts = []
                for chunk_index, field_group in enumerate(groups, start=1):
                    filename = f"{safe_name(table)}__part{chunk_index:03d}.jsonl"
                    row_count = export_chunk(conn, table, field_group, data_dir / filename)
                    row_counts.append(row_count)
                    entry["chunks"].append({"file": f"tables/{filename}", "fields": field_group, "rows": row_count})
                entry["row_count"] = row_counts[0]
                entry["consistent_chunk_counts"] = len(set(row_counts)) == 1
                entry["status"] = "exported"
                print(f"[{index}/{len(inventory)}] {table}: {entry['row_count']} sample row(s), {len(groups)} part(s)", flush=True)
            except Exception as exc:
                entry["status"] = "error"
                entry["error"] = repr(exc)
                manifest["errors"].append({"table": table, "error": repr(exc)})
                print(f"[{index}/{len(inventory)}] {table}: ERROR {exc!r}", flush=True)
            (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    finally:
        conn.close()

    manifest["finished_at"] = datetime.now().isoformat(timespec="seconds")
    exported = sum(1 for item in manifest["tables"].values() if item["status"] == "exported")
    total_rows = sum(int(item.get("row_count") or 0) for item in manifest["tables"].values())
    manifest["summary"] = {"catalog_objects": len(manifest["tables"]), "exported": exported, "errors": len(manifest["errors"]), "total_rows": total_rows}
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (out / "README.md").write_text(
        "# Exportação read-only de tabelas /SBXC/* em PRD\n\n"
        f"- Objetos encontrados: {len(manifest['tables'])}\n"
        f"- Exportados: {exported}\n"
        f"- Erros: {len(manifest['errors'])}\n"
        f"- Total de linhas amostradas (por tabela): {total_rows}\n"
        f"- Limite por tabela: {SAMPLE_SIZE}\n\n"
        "Os conteúdos são amostras em JSON Lines. Tabelas largas são divididas em partes por grupos de campos; o manifesto descreve a composição.\n",
        encoding="utf-8",
    )
    print(json.dumps({"run_dir": str(out), "summary": manifest["summary"]}, ensure_ascii=False, indent=2), flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

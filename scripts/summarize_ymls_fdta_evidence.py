from __future__ import annotations

import json
from pathlib import Path


path = Path(__file__).resolve().parents[1] / "output" / "ymls_fdta_readonly_evidence.json"
data = json.loads(path.read_text(encoding="utf-8"))
print("REGUT_FIELDS", [row.get("FIELDNAME") for row in data.get("ddic", {}).get("REGUT", [])])


def show(label: str, item: object) -> None:
    print(f"\n### {label}")
    if not isinstance(item, dict):
        print(item)
        return
    if "error" in item:
        print("ERROR:", item["error"])
    rows = item.get("rows", [])
    print("ROWS:", len(rows))
    for row in rows:
        print({key: str(value).strip() for key, value in row.items() if str(value).strip()})


for label, item in data.get("transaction", {}).items():
    show(f"transaction/{label}", item)
for label, item in data.get("menu", {}).items():
    show(f"menu/{label}", item)
show("ymls_package_objects", data.get("ymls_package_objects", {}))

tables = data.get("repository_discovery", {}).get("TFPM_tables", {}).get("rows", [])
print("\n### TFPM TABLES")
print([row.get("TABNAME", "").strip() for row in tables])

for run, table_map in data.get("runs", {}).items():
    for table, item in table_map.items():
        show(f"runs/{run}/{table}", item)

for table, comparisons in data.get("format_search", {}).items():
    for version, item in comparisons.items():
        show(f"format/{table}/{version}", item)
for table, comparisons in data.get("all_tfpm_format_search", {}).items():
    for version, item in comparisons.items():
        if item.get("rows") or item.get("error"):
            show(f"all_format/{table}/{version}", item)
show("filename_search", data.get("filename_search", {}))

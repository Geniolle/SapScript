from __future__ import annotations

import json
from pathlib import Path
from typing import Any


path = Path(__file__).resolve().parents[1] / "output" / "ymls_fdta_readonly_evidence.json"
data = json.loads(path.read_text(encoding="utf-8"))


def clean_rows(item: dict[str, Any]) -> list[dict[str, str]]:
    return [{key: str(value).strip() for key, value in row.items() if str(value).strip()} for row in item.get("rows", [])]


def emit(label: str, item: dict[str, Any]) -> None:
    print(f"\n## {label}")
    if item.get("error"):
        print("ERROR", item["error"])
    for row in clean_rows(item):
        print(row)


for key, item in data.get("transaction", {}).items():
    emit(f"TRANSACTION {key}", item)

objects = clean_rows(data.get("ymls_package_objects", {}))
print("\n## YMLS OBJECT COUNTS")
counts: dict[str, int] = {}
for row in objects:
    counts[row.get("OBJECT", "")] = counts.get(row.get("OBJECT", ""), 0) + 1
print(counts)
for row in objects:
    if "PAYMEDIUM" in row.get("OBJ_NAME", "") or row.get("OBJ_NAME") in {"YMLS_T_BCM_FNAME", "YMLS_T_BCM_FNAMH", "YMLS_T_BCM_FNAMT", "YMLS_T_BCM_PAYM", "YMLS_T_BCM_DIR"}:
        print(row)

for table, versions in data.get("all_tfpm_format_search", {}).items():
    for version, item in versions.items():
        if item.get("rows"):
            emit(f"FORMAT {table} {version}", item)

for key, item in data.get("regut_narrow", {}).items():
    emit(f"REGUT {key}", item)

for name, fm in data.get("function_modules", {}).items():
    for table, item in fm.items():
        emit(f"FM {name} {table}", item)

for table, item in data.get("ymls_payment_config", {}).items():
    print(f"\n## YMLS CONFIG {table}")
    if item.get("error"):
        print("ERROR", item["error"])
    for row in clean_rows(item):
        text = " ".join(row.values()).casefold()
        if any(token in text for token in ("isprptpl", "sct", "z_sepa_ap", "z_pt_cgi_xml_ct_v9", "cgi", "sepa")):
            print(row)

for table, content in data.get("obpm_variant_tables", {}).items():
    print(f"\n## OBPM TABLE {table} FIELDS")
    print(content.get("fields"))
    for version in ("old", "new"):
        emit(f"OBPM {table} {version}", content.get(version, {}))

for key, item in data.get("obpm_standard_modules", {}).items():
    emit(f"STANDARD MODULES {key}", item)

tokens = ["Z_SEPA_AP", "Z_PT_CGI_XML_CT_V9", "ISPRPTPL", "SCT", "FILENAME", "FILE_NAME", "C_FILENAME", "FORMI", "DWNAM", "PATH", "INCLUDE"]
for program, source in data.get("program_sources", {}).items():
    print(f"\n## SOURCE {program}")
    if source.get("error"):
        print("ERROR", source["error"])
        continue
    lines = source.get("SOURCE_EXTENDED") or source.get("SOURCE") or []
    texts = [str(line.get("LINE", line.get("ZEILE", line))) if isinstance(line, dict) else str(line) for line in lines]
    found: set[int] = set()
    for index, text in enumerate(texts):
        if any(token.casefold() in text.casefold() for token in tokens):
            found.update(range(max(0, index - 4), min(len(texts), index + 5)))
    for index in sorted(found):
        print(f"{index + 1}: {texts[index]}")

print("\n## CROSS REFERENCE FIELDS")
for table, fields in data.get("cross_reference_ddic", {}).items():
    print(table, [str(item.get("FIELDNAME", "")).strip() for item in fields])

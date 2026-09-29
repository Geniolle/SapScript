from __future__ import annotations

import json
from pathlib import Path


data = json.loads((Path(__file__).resolve().parents[1] / "output" / "ymls_fdta_readonly_evidence.json").read_text(encoding="utf-8"))

for section in ("obpm_variant_tables", "obpm_standard_modules"):
    print("SECTION", section)
    content = data.get(section, {})
    for key, value in content.items():
        print("TABLE", key)
        if section == "obpm_variant_tables":
            print("FIELDS", value.get("fields"))
            for version in ("old", "new"):
                item = value.get(version, {})
                print(version, "ERROR", item.get("error"), "ROWS", len(item.get("rows", [])))
                for row in item.get("rows", []):
                    print({k: str(v).strip() for k, v in row.items() if str(v).strip()})
        else:
            rows = value.get("rows", [])
            for row in rows:
                form = str(row.get("FORMI", "")).strip()
                if form in {"Z_SEPA_AP", "Z_PT_CGI_XML_CT_V9"}:
                    print({k: str(v).strip() for k, v in row.items() if str(v).strip()})

target = json.loads((Path(__file__).resolve().parents[1] / "output" / "filename_targeted_readonly_evidence.json").read_text(encoding="utf-8"))
for table, block in target.get("custom_tables", {}).items():
    if table == "YMLS_T_BCM_DIR":
        print("TABLE", table, "FIELDS", block.get("fields"))
        for row in block.get("data", {}).get("rows", []):
            print({k: str(v).strip() for k, v in row.items() if str(v).strip()})

from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
path = ROOT / "output" / "dmee_compare_Z_SEPA_CT__Z_PT_CGI_XML_CT_V9.json"

with path.open("r", encoding="utf-8") as handle:
    data = json.load(handle)

for item in data.get("different_fields", []):
    print(item.get("field", ""))

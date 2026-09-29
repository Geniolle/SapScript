from __future__ import annotations

import json
from pathlib import Path

from investigate_ymls_fdta_readonly import ReadOnlySap, normalise


OUT = Path(__file__).resolve().parents[1] / "output" / "filename_targeted_readonly_evidence.json"


def main() -> None:
    sap = ReadOnlySap()
    evidence = {"mode": "SAP PRD READ-ONLY"}
    try:
        evidence["manager_tadir"] = sap.read(
            "TADIR", ["PGMID", "OBJECT", "OBJ_NAME", "DEVCLASS", "AUTHOR", "MASTERLANG"],
            ["OBJ_NAME = 'YMLS_CL_BCM_MANAGER'"], 20
        )
        evidence["get_filename_refs"] = sap.read(
            "WBCROSSGT", ["OTYPE", "NAME", "INCLUDE", "DIRECT", "INDIRECT", "COMPONENT", "UDATE", "UNAME"],
            ["NAME = 'GET_FILENAME'"], 5000
        )
        includes = {
            str(row.get("INCLUDE", "")).strip()
            for row in evidence["get_filename_refs"].get("rows", [])
            if str(row.get("INCLUDE", "")).strip()
        }
        likely = {
            "YMLS_CL_BCM_MANAGER=====CP", "YMLS_CL_BCM_MANAGER=====CM001",
            "YMLS_CL_BCM_MANAGER=====CCIMP", "YMLS_CL_BCM_MANAGER=====CCDEF",
        }
        evidence["sources"] = {name: sap.program(name) for name in sorted(includes | likely)}
        evidence["function_group_tadir"] = sap.read(
            "TADIR", ["PGMID", "OBJECT", "OBJ_NAME", "DEVCLASS", "AUTHOR", "MASTERLANG"],
            ["OBJ_NAME = 'YMLS_BCM_FUNCTIONS' OR OBJ_NAME = 'Z_FI_PAYMEDIUM_21'"], 50
        )
        evidence["custom_tables"] = {}
        for table in ["YMLS_T_BCM_FNAME", "YMLS_T_BCM_FNAMH", "YMLS_T_BCM_FNAMT", "YMLS_T_BCM_PAYM", "YMLS_T_BCM_DIR"]:
            metadata = sap.fields(table)
            fields = [str(item.get("FIELDNAME", "")).strip() for item in metadata]
            chosen = fields[:]
            total = 0
            safe = []
            for item in metadata:
                size = int(item.get("LENG", 0) or 0)
                name = str(item.get("FIELDNAME", "")).strip()
                if size <= 180 and total + size + 1 <= 450:
                    safe.append(name)
                    total += size + 1
            evidence["custom_tables"][table] = {
                "fields": chosen,
                "data": sap.read(table, safe, None, 5000),
            }
    finally:
        sap.conn.close()
    OUT.write_text(json.dumps(normalise(evidence), ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    main()

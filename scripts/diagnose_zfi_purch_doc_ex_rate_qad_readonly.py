from __future__ import annotations

import json
import os
from pathlib import Path

from pyrfc import Connection

from sap_rfc._rfc_common import build_connection_params_for_env, find_project_root, load_project_env


ALLOWED_FUNCTIONS = {
    "RFC_SYSTEM_INFO",
    "RFC_READ_TABLE",
    "DDIF_FIELDINFO_GET",
    "RPY_PROGRAM_READ",
}


def call_readonly(conn: Connection, function_name: str, **kwargs):
    if function_name not in ALLOWED_FUNCTIONS:
        raise RuntimeError(f"RFC bloqueada pelo diagnóstico read-only: {function_name}")
    return conn.call(function_name, **kwargs)


def source_lines(response: dict) -> list[str]:
    rows = response.get("SOURCE") or response.get("SOURCE_EXTENDED") or response.get("SOURCE_TAB") or []
    lines: list[str] = []
    for row in rows:
        if isinstance(row, dict):
            lines.append(str(row.get("LINE") or row.get("TEXT") or row.get("SOURCE_LINE") or row.get("ABAP") or ""))
        else:
            lines.append(str(row))
    return lines


def include_names(response: dict) -> list[str]:
    result: list[str] = []
    for key in ("INCLUDE_TAB", "INCLUDETAB", "INCLUDES"):
        for row in response.get(key) or []:
            if isinstance(row, dict):
                value = row.get("INCLUDE") or row.get("INCL_NAME") or row.get("INCLNAME") or row.get("NAME") or row.get("PROGRAM") or row.get("PROGNAME")
            else:
                value = row
            if value:
                result.append(str(value).strip().upper())
    return list(dict.fromkeys(result))


def read_program(conn: Connection, name: str) -> dict:
    return call_readonly(
        conn,
        "RPY_PROGRAM_READ",
        PROGRAM_NAME=name,
        LANGUAGE=os.getenv("SAP_QAD_LANG", "PT").strip() or "PT",
        WITH_INCLUDELIST="X",
        ONLY_SOURCE="X",
        READ_LATEST_VERSION="X",
        WITH_LOWERCASE="X",
    )


def main() -> None:
    load_project_env(find_project_root())
    conn = Connection(**build_connection_params_for_env("QAD"))
    try:
        system = call_readonly(conn, "RFC_SYSTEM_INFO")
        root = read_program(conn, "ZFI_PURCH_DOC_EX_RATE")
        names = ["ZFI_PURCH_DOC_EX_RATE", *include_names(root)]
        output = {
            "system": system,
            "root_response_shape": {
                key: (len(value) if isinstance(value, (list, dict, str)) else str(type(value).__name__))
                for key, value in root.items()
            },
            "root_include_rows": root.get("INCLUDE_TAB") or root.get("INCLUDETAB") or root.get("INCLUDES") or [],
            "programs": {},
        }
        for name in list(dict.fromkeys(names)):
            response = root if name == "ZFI_PURCH_DOC_EX_RATE" else read_program(conn, name)
            lines = source_lines(response)
            output["programs"][name] = {
                "line_count": len(lines),
                "includes": include_names(response),
                "source": lines,
            }
        output_path = Path("output") / "zfi_purch_doc_ex_rate_qad_source.json"
        output_path.parent.mkdir(exist_ok=True)
        output_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
        print(output_path.resolve())
        print("programs=" + ",".join(output["programs"]))
    finally:
        conn.close()


if __name__ == "__main__":
    main()

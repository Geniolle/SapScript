from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import deque
from datetime import datetime
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from sap_rfc._rfc_common import (  # noqa: E402
    build_connection_params_for,
    find_project_root,
    load_project_env,
)


DELIMITER = "|"
SOURCE_KEYS = ("LINE", "TEXT", "SOURCE_LINE", "ABAP")


def stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.#$@-]+", "_", name)[:180]


def parse_rows(result: dict[str, Any], fields: list[str]) -> list[dict[str, str]]:
    rows = []
    for item in result.get("DATA", []) or []:
        parts = str(item.get("WA", "")).split(DELIMITER)
        parts += [""] * max(0, len(fields) - len(parts))
        rows.append({field: parts[i].strip() for i, field in enumerate(fields)})
    return rows


def read_table(conn: Any, table: str, fields: list[str], condition: str, errors: list[dict[str, str]]) -> list[dict[str, str]]:
    try:
        result = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table,
            DELIMITER=DELIMITER,
            FIELDS=[{"FIELDNAME": field} for field in fields],
            OPTIONS=[{"TEXT": condition}] if condition else [],
            ROWCOUNT=0,
        )
        return parse_rows(dict(result or {}), fields)
    except Exception as exc:  # extraction must preserve partial evidence
        errors.append({"scope": f"RFC_READ_TABLE {table} [{condition}]", "error": repr(exc)})
        return []


def normalize_source(value: Any) -> list[str]:
    if isinstance(value, str):
        return value.splitlines()
    lines = []
    for item in value or []:
        if isinstance(item, dict):
            found = next((item.get(key) for key in SOURCE_KEYS if item.get(key) is not None), "")
            lines.append(str(found))
        else:
            lines.append(str(item))
    return lines


def normalize_includes(response: dict[str, Any]) -> list[str]:
    found: list[str] = []
    for key in ("INCLUDE_TAB", "INCLUDETAB", "INCLUDES"):
        rows = response.get(key) or []
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            if isinstance(row, dict):
                value = next((str(row.get(k) or "").strip() for k in ("INCLNAME", "INCLUDE", "NAME", "PROGRAM") if row.get(k)), "")
            else:
                value = str(row).strip()
            if value:
                found.append(value.upper())
    return list(dict.fromkeys(found))


def read_program(conn: Any, name: str, language: str, errors: list[dict[str, str]]) -> dict[str, Any]:
    try:
        response = dict(conn.call(
            "RPY_PROGRAM_READ",
            PROGRAM_NAME=name,
            LANGUAGE=language,
            WITH_INCLUDELIST="X",
            ONLY_SOURCE="X",
            READ_LATEST_VERSION="X",
            WITH_LOWERCASE="X",
        ) or {})
        source = normalize_source(response.get("SOURCE") or response.get("SOURCE_EXTENDED") or response.get("SOURCE_TAB"))
        return {
            "name": name,
            "ok": True,
            "source": source,
            "includes": normalize_includes(response),
            "program_info": response.get("PROG_INF") or {},
            "response_keys": sorted(response),
        }
    except Exception as exc:
        errors.append({"scope": f"RPY_PROGRAM_READ {name}", "error": repr(exc)})
        return {"name": name, "ok": False, "source": [], "includes": [], "error": repr(exc)}


PATTERNS = {
    "function_modules": r"\bCALL\s+FUNCTION\s+['`]([^'`]+)['`]",
    "bapis": r"\bCALL\s+FUNCTION\s+['`](BAPI_[^'`]+)['`]",
    "classes": r"\b(?:NEW|TYPE\s+REF\s+TO|CREATE\s+OBJECT(?:\s+\w+)?\s+TYPE|CLASS)\s+([A-Z/][A-Z0-9_/$#]+)|\b([A-Z/][A-Z0-9_/$#]+)=>",
    "badis": r"\b(?:GET\s+BADI|CALL\s+BADI)\s+([A-Z/][A-Z0-9_/$#]+)",
    "enhancements": r"\b(?:ENHANCEMENT|ENHANCEMENT-POINT|ENHANCEMENT-SECTION)\s+([A-Z/][A-Z0-9_/$#]+)",
    "tables": r"\b(?:FROM|JOIN|UPDATE|INSERT\s+INTO|MODIFY|DELETE\s+FROM)\s+([A-Z/][A-Z0-9_/$#]+)",
    "transactions": r"\bCALL\s+TRANSACTION\s+['`]([^'`]+)['`]",
    "submits": r"\bSUBMIT\s+([A-Z/][A-Z0-9_/$#]+)",
}


def static_refs(lines: list[str]) -> dict[str, list[str]]:
    text = "\n".join(lines).upper()
    result: dict[str, list[str]] = {}
    for key, pattern in PATTERNS.items():
        values = set()
        for match in re.finditer(pattern, text, re.MULTILINE):
            values.add(next(group for group in match.groups() if group))
        result[key] = sorted(values)
    return result


def class_pool_name(class_name: str) -> str:
    return class_name.upper().ljust(30, "=") + "CP"


def main() -> int:
    parser = argparse.ArgumentParser(description="Read-only SAP ABAP program and dependency extraction")
    parser.add_argument("program")
    parser.add_argument("--env", default="PRD")
    parser.add_argument("--language", default="PT")
    parser.add_argument("--out", default=str(PROJECT_ROOT / "output"))
    args = parser.parse_args()

    project_root = find_project_root()
    load_project_env(project_root)
    from pyrfc import Connection

    program = args.program.strip().upper()
    env = args.env.strip().upper()
    run_dir = Path(args.out).resolve() / f"{safe_name(program)}_{env}_{stamp()}"
    source_dir = run_dir / "abap"
    source_dir.mkdir(parents=True, exist_ok=True)
    errors: list[dict[str, str]] = []

    params = build_connection_params_for(env)
    conn = Connection(**params)
    try:
        conn.call("RFC_PING")
        repository = {
            "TRDIR": read_table(conn, "TRDIR", ["NAME", "SUBC", "APPL", "RSTAT", "CDAT", "UDAT", "CNAM", "UNAM"], f"NAME = '{program}'", errors),
            "TADIR": read_table(conn, "TADIR", ["PGMID", "OBJECT", "OBJ_NAME", "DEVCLASS", "AUTHOR", "SRCSYSTEM", "CREATED_ON"], f"OBJ_NAME = '{program}'", errors),
            "D010INC": read_table(conn, "D010INC", ["MASTER", "INCLUDE"], f"MASTER = '{program}'", errors),
        }

        queue = deque([program])
        objects: dict[str, dict[str, Any]] = {}
        while queue:
            name = queue.popleft()
            if name in objects:
                continue
            item = read_program(conn, name, args.language, errors)
            item["static_references"] = static_refs(item["source"])
            objects[name] = item
            queue.extend(include for include in item["includes"] if include not in objects)

        combined_refs = {key: set() for key in PATTERNS}
        for item in objects.values():
            for key, values in item["static_references"].items():
                combined_refs[key].update(values)

        class_metadata: dict[str, Any] = {}
        for class_name in sorted(combined_refs["classes"]):
            rows = read_table(conn, "SEOCLASS", ["CLSNAME", "VERSION", "STATE", "CLSCCINCL", "AUTHOR", "CREATEDON", "CHANGEDBY", "CHANGEDON"], f"CLSNAME = '{class_name}' AND VERSION = '1'", errors)
            pool = class_pool_name(class_name)
            class_item = read_program(conn, pool, args.language, errors)
            class_item["static_references"] = static_refs(class_item["source"])
            class_metadata[class_name] = {"SEOCLASS": rows, "class_pool": class_item}

        fm_metadata: dict[str, Any] = {}
        for fm_name in sorted(combined_refs["function_modules"]):
            fm_metadata[fm_name] = {
                "TFDIR": read_table(conn, "TFDIR", ["FUNCNAME", "PNAME", "INCLUDE", "FMODE", "UTASK", "REMOTE"], f"FUNCNAME = '{fm_name}'", errors),
                "ENLFDIR": read_table(conn, "ENLFDIR", ["FUNCNAME", "AREA", "ACTIVE"], f"FUNCNAME = '{fm_name}'", errors),
            }

        badi_metadata: dict[str, Any] = {}
        for badi_name in sorted(combined_refs["badis"]):
            badi_metadata[badi_name] = {
                "SXS_ATTR": read_table(conn, "SXS_ATTR", ["EXIT_NAME", "INTERNAL", "SMOD_EXIT"], f"EXIT_NAME = '{badi_name}'", errors),
                "SXS_INTER": read_table(conn, "SXS_INTER", ["EXIT_NAME", "INTER_NAME"], f"EXIT_NAME = '{badi_name}'", errors),
                "SXC_EXIT": read_table(conn, "SXC_EXIT", ["EXIT_NAME", "IMP_NAME", "ACTIVE"], f"EXIT_NAME = '{badi_name}'", errors),
            }

        for name, item in objects.items():
            (source_dir / f"{safe_name(name)}.abap").write_text("\n".join(item["source"]) + "\n", encoding="utf-8")
        for class_name, data in class_metadata.items():
            pool = data["class_pool"]
            if pool["source"]:
                (source_dir / f"{safe_name(pool['name'])}.abap").write_text("\n".join(pool["source"]) + "\n", encoding="utf-8")

        report = {
            "extraction": {"timestamp": datetime.now().isoformat(timespec="seconds"), "environment": env, "program": program, "read_only": True},
            "summary": {
                "program_objects": len(objects),
                "source_lines": sum(len(item["source"]) for item in objects.values()),
                **{key: len(values) for key, values in combined_refs.items()},
                "errors": len(errors),
            },
            "repository": repository,
            "references": {key: sorted(values) for key, values in combined_refs.items()},
            "function_module_metadata": fm_metadata,
            "class_metadata": class_metadata,
            "badi_metadata": badi_metadata,
            "objects": objects,
            "errors": errors,
        }
        (run_dir / "technical_extraction.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

        readme = [
            f"# Extração técnica RFC — `{program}`", "",
            f"- Ambiente: `{env}`", f"- Data/hora: `{report['extraction']['timestamp']}`", "- Modo: somente leitura", "",
            "## Resumo", "",
        ]
        readme.extend(f"- {key}: {value}" for key, value in report["summary"].items())
        for key, values in report["references"].items():
            readme.extend(["", f"## {key}", ""] + ([f"- `{v}`" for v in values] or ["- Nenhuma referência estática encontrada."]))
        if errors:
            readme.extend(["", "## Limitações/erros RFC", ""])
            readme.extend(f"- `{item['scope']}`: {item['error']}" for item in errors)
        (run_dir / "README.md").write_text("\n".join(readme) + "\n", encoding="utf-8")
        print(json.dumps({"run_dir": str(run_dir), "summary": report["summary"]}, ensure_ascii=False, indent=2))
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

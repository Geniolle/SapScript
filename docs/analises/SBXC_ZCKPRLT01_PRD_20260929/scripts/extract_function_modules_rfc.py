from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from sap_rfc._rfc_common import build_connection_params_for, find_project_root, load_project_env  # noqa: E402


FUNCTIONS = [
    "/SBXC/ZCKP_CAB_FAT",
    "/SBXC/ZCKP_ITEM_FAT",
    "/SBXC/ZCKP_DBL_CLK",
    "/SBXC/ZCKP_GUARDA_ALT_FAT",
    "/SBXC/ZCKP_VALIDA_PROCESSO",
    "/SBXC/ZCKP_BASELINE_DATE",
    "/SBXC/ZCKP_DBL_CLK_LIN",
    "/SBXC/ZCKP_LOAD_DD03L",
    "/SBXC/ZCKP_LOAD_TAB10",
]


def lines(value: Any) -> list[str]:
    result = []
    for item in value or []:
        if isinstance(item, dict):
            result.append(str(item.get("LINE") or item.get("TEXT") or item.get("SOURCE_LINE") or ""))
        else:
            result.append(str(item))
    return result


def safe(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value)


def main() -> int:
    load_project_env(find_project_root())
    from pyrfc import Connection

    out = ROOT / "output" / f"ZCKP_function_modules_PRD_{datetime.now():%Y%m%d_%H%M%S}"
    src = out / "abap"
    src.mkdir(parents=True, exist_ok=True)
    report: dict[str, Any] = {"environment": "PRD", "read_only": True, "timestamp": datetime.now().isoformat(timespec="seconds"), "functions": {}, "errors": []}
    conn = Connection(**build_connection_params_for("PRD"))
    try:
        conn.call("RFC_PING")
        for name in FUNCTIONS:
            try:
                response = dict(conn.call("RPY_FUNCTIONMODULE_READ", FUNCTIONNAME=name) or {})
                source = lines(response.get("SOURCE") or response.get("SOURCE_EXTENDED") or response.get("SOURCE_TAB"))
                report["functions"][name] = {
                    "response_keys": sorted(response),
                    "source": source,
                    "interface": {key: value for key, value in response.items() if key not in {"SOURCE", "SOURCE_EXTENDED", "SOURCE_TAB"}},
                }
                if source:
                    (src / f"{safe(name)}.abap").write_text("\n".join(source) + "\n", encoding="utf-8")
            except Exception as exc:
                report["errors"].append({"function": name, "error": repr(exc)})
        (out / "function_modules.json").write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        print(json.dumps({"run_dir": str(out), "functions": {name: len(item["source"]) for name, item in report["functions"].items()}, "errors": report["errors"]}, ensure_ascii=False, indent=2))
    finally:
        conn.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

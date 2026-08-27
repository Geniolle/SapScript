from __future__ import annotations

import os
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
os.environ.setdefault("SAP_SCRIPT_PROJECT_DIR", str(PROJECT_ROOT))

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from sap_agent.sap_gui_actions import open_transaction


TRANSACTION_CODE = "FB01"
DESCRIPTION = "Abrir FB01 para criar Documento FI"


def main() -> int:
    try:
        result = open_transaction(TRANSACTION_CODE, DESCRIPTION)
    except Exception as exc:
        print(f"ERRO: {exc}")
        return 1

    print(result.result_text)
    if getattr(result, "error", None):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

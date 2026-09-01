from __future__ import annotations

import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import find_project_root, load_project_env


def main() -> int:
    project_root = find_project_root()
    load_project_env(project_root)
    keys = [
        "SAP_FI_TAX_CODE",
        "SAP_FI_TAX_AMOUNT",
        "SAP_FI_TAX_RATE",
        "SAP_FI_TAX_GL_ACCOUNT",
        "SAP_FI_TAX_DIRECTION",
        "SAP_FI_WITHHOLDING_TAX_TYPE",
        "SAP_FI_WITHHOLDING_TAX_CODE",
        "SAP_FI_WITHHOLDING_TAX_BASE_AMOUNT",
        "SAP_FI_WITHHOLDING_TAX_AMOUNT",
        "SAP_FI_FORM_PAGTO_FORNECEDOR",
        "SAP_QAD_USER",
        "SAP_QAD_PASSWD",
    ]
    for key in keys:
        print(f"{key}={'set' if os.getenv(key) else 'missing'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import json
import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ["SAP_TARGET_ENV"] = "PRD"

from sap_rfc.user_data_service import analyze_user_data  # noqa: E402


if __name__ == "__main__":
    username = sys.argv[1] if len(sys.argv) > 1 else "S80001870"
    print(json.dumps(analyze_user_data(username, "master"), ensure_ascii=False, indent=2))

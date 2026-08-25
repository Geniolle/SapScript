# -*- coding: utf-8 -*-
"""
Entry point da simulacao UAT RFF110S.
"""

from __future__ import annotations

import sys
from pathlib import Path
import importlib

BASE_DIR = Path(__file__).resolve().parent
PARENT_DIR = BASE_DIR.parent
for candidate in (PARENT_DIR, BASE_DIR):
    candidate_str = str(candidate)
    if candidate_str not in sys.path:
        sys.path.insert(0, candidate_str)

importlib.invalidate_caches()
sys.modules.pop("rff110s_uat_orchestrator", None)
_module = importlib.import_module("rff110s_uat_orchestrator")
_module = importlib.reload(_module)
def executar(*args, **kwargs):
    return _module.main()


main = executar


if __name__ == "__main__":
    raise SystemExit(main())

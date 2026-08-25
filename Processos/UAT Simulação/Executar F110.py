# -*- coding: utf-8 -*-
"""
Compatibilidade para executar o fluxo completo da F110 a partir do Cockpit.
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
sys.modules.pop("executar_f110", None)
_module = importlib.import_module("executar_f110")
_module = importlib.reload(_module)
def executar(*args, **kwargs):
    return _module.executar(**kwargs)


main = _module.main


if __name__ == "__main__":
    raise SystemExit(main())

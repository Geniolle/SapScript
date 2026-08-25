# -*- coding: utf-8 -*-
"""
Compatibilidade com o nome legado do processo de criacao de documento FI.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PARENT_DIR = BASE_DIR.parent
for candidate in (PARENT_DIR, BASE_DIR):
    candidate_str = str(candidate)
    if candidate_str not in sys.path:
        sys.path.insert(0, candidate_str)

_module_path = BASE_DIR / "criar_documento_teste_f110.py"
_spec = importlib.util.spec_from_file_location(
    "uat_simulacao_criar_documento_teste_f110",
    _module_path,
)
if _spec is None or _spec.loader is None:
    raise ImportError(f"Não foi possível carregar o módulo RFC local: {_module_path}")

_module = importlib.util.module_from_spec(_spec)
sys.modules[_spec.name] = _module
_spec.loader.exec_module(_module)
executar = _module.executar
main = _module.main


if __name__ == "__main__":
    raise SystemExit(main())

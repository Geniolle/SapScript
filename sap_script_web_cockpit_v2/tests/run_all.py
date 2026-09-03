"""
Corre a rede de seguranca do cockpit e devolve exit code combinado.

    python tests/run_all.py        # a partir de sap_script_web_cockpit_v2/

Inclui:
  1. tests/js_smoke.py            -> <script> inline avalia limpo (TDZ, duplicados, codigo morto)
  2. unittest discover em tests/  -> rotas do Agente Salsa IT e afins

Usar o Python do venv do cockpit (tem fastapi/pydantic):
    .venv\\Scripts\\python.exe tests\\run_all.py
"""

from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
COCKPIT_DIR = HERE.parent


def run_js_smoke() -> int:
    print("== js_smoke ==", flush=True)
    proc = subprocess.run([sys.executable, str(HERE / "js_smoke.py")], cwd=str(COCKPIT_DIR))
    return proc.returncode


def run_unittests() -> int:
    print("== unittest (tests/) ==", flush=True)
    loader = unittest.TestLoader()
    suite = loader.discover(start_dir=str(HERE), top_level_dir=str(COCKPIT_DIR))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    return 0 if result.wasSuccessful() else 1


def main() -> int:
    rc = 0
    rc |= run_js_smoke()
    rc |= run_unittests()
    print("\nRESULTADO:", "OK" if rc == 0 else "FALHOU")
    return rc


if __name__ == "__main__":
    sys.exit(main())

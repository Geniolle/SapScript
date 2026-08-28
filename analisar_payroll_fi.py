#!/usr/bin/env python
"""Launcher do diagnóstico READ-ONLY Payroll (RH) -> FI.

Exemplos:

    python analisar_payroll_fi.py --diagnostic
    python analisar_payroll_fi.py
    python analisar_payroll_fi.py --company 2010 --year 2026 --month 6 --account 23120000

O `main.py` da raiz pertence ao cockpit web; este script é o ponto de entrada
do pacote `sap_payroll_analysis`.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from sap_payroll_analysis.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())

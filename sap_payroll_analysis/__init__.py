"""Diagnóstico READ-ONLY Payroll (RH) -> FI via RFC.

Investiga divergências entre o posting do Payroll (PCP0) e o que foi
efectivamente contabilizado em FI numa conta do Razão.

Todo o package é estritamente de leitura. Ver `security.py`.
"""

from .config import AnalysisParams, DEFAULTS

__all__ = ["AnalysisParams", "DEFAULTS"]

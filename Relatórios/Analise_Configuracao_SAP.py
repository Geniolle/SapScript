# -*- coding: utf-8 -*-
"""Launcher de compatibilidade.

A implementação foi refatorada para:
    Relatórios/Analises_Tabelas_SAP/

Este ficheiro continua a funcionar para não quebrar comandos antigos e executa,
por defeito, o processo ``metodos_pagamento_pt``.

Novo comando recomendado:
    .venv\Scripts\python.exe "Relatórios\Analises_Tabelas_SAP\runner.py" metodos_pagamento_pt
"""
from __future__ import annotations

import sys
from pathlib import Path


RELATORIOS_DIR = Path(__file__).resolve().parent
PACKAGE_DIR = RELATORIOS_DIR / "Analises_Tabelas_SAP"

if str(PACKAGE_DIR) not in sys.path:
    sys.path.insert(0, str(PACKAGE_DIR))

from runner import main  # noqa: E402


if __name__ == "__main__":
    raise SystemExit(main(["metodos_pagamento_pt"]))

# -*- coding: utf-8 -*-
"""Runner dos processos de análise de tabelas SAP.

Uso:
    .venv\Scripts\python.exe "Relatórios\Analises_Tabelas_SAP\runner.py" metodos_pagamento_pt

Listar processos disponíveis:
    .venv\Scripts\python.exe "Relatórios\Analises_Tabelas_SAP\runner.py" --listar
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from types import ModuleType


BASE_DIR = Path(__file__).resolve().parent
PROCESSOS_DIR = BASE_DIR / "processos"

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from engine import executar_processo  # noqa: E402


def listar_processos() -> list[str]:
    nomes: list[str] = []
    if not PROCESSOS_DIR.exists():
        return nomes
    for path in sorted(PROCESSOS_DIR.glob("*.py")):
        if path.name.startswith("_"):
            continue
        nomes.append(path.stem)
    return nomes


def carregar_processo(nome: str) -> dict:
    nome_limpo = str(nome or "").strip()
    if not nome_limpo:
        raise ValueError("Informe o nome do processo.")

    path = PROCESSOS_DIR / f"{nome_limpo}.py"
    if not path.exists():
        disponiveis = ", ".join(listar_processos()) or "(nenhum)"
        raise FileNotFoundError(
            f"Processo '{nome_limpo}' não encontrado em {PROCESSOS_DIR}. "
            f"Disponíveis: {disponiveis}"
        )

    spec = importlib.util.spec_from_file_location(f"sap_table_process_{nome_limpo}", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Não foi possível carregar o processo: {path}")

    module: ModuleType = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    config = getattr(module, "PROCESSO", None)
    if not isinstance(config, dict):
        raise RuntimeError(
            f"O ficheiro {path.name} deve exportar um dicionário chamado PROCESSO."
        )
    return config


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Executa processos reutilizáveis de consulta SAP via SE16H/SE16N."
    )
    parser.add_argument(
        "processo",
        nargs="?",
        default="metodos_pagamento_pt",
        help="Nome do ficheiro em processos/ sem .py",
    )
    parser.add_argument(
        "--listar",
        action="store_true",
        help="Lista os processos disponíveis e termina.",
    )
    args = parser.parse_args(argv)

    if args.listar:
        print("Processos disponíveis:")
        for nome in listar_processos():
            print(f"  - {nome}")
        return 0

    try:
        config = carregar_processo(args.processo)
        return executar_processo(config)
    except Exception as exc:
        print(f"❌ {exc}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

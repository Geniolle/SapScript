# -*- coding: utf-8 -*-
"""
sap_rfc/projeto_perfil_cli.py
======================================================================
CLI JSON para execução subprocessual isolada das 6 operações do
Projeto Perfil no ambiente virtual .venv-rfc.
======================================================================
"""

from __future__ import annotations

import argparse
import json
import sys

from sap_rfc.projeto_perfil_service import (
    executar_processo_completo,
    fluxo_departamento,
    auditoria_utilizador,
    pesquisa_atribuir_transacao,
    correcao_sincronizacao_posterior,
    diagnostico_su53,
)


def parse_args(args_list: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI JSON para operações do Projeto Perfil de Autorização."
    )
    parser.add_argument(
        "--action",
        dest="action",
        required=True,
        choices=["execucao", "departamento", "utilizador", "pesquisa", "corrigir", "su53"],
        help="Ação a executar.",
    )
    parser.add_argument("--departamento", "-d", dest="departamento", default="", help="Nome ou linha do departamento.")
    parser.add_argument("--subacao", dest="subacao", default="analisar", help="Subação para departamento (listar, analisar, global, validar_users, cruzar_fontes).")
    parser.add_argument("--user", "--username", "-u", dest="username", default="", help="ID do utilizador SAP.")
    parser.add_argument("--tcode", "-t", dest="tcode", default="", help="Código da transação SAP.")
    parser.add_argument("--env", "--environment", dest="environment", default="PRD", help="Ambiente SAP (PRD / QAD).")
    parser.add_argument("--tcode-atribuir", dest="tcode_atribuir", default="", help="Transação a atribuir na SU53.")
    parser.add_argument("--simular", dest="simular", action="store_true", help="Modo de simulação sem gravação.")
    parser.add_argument("--excel", dest="caminho_excel", default=None, help="Caminho do ficheiro Excel.")
    return parser.parse_args(args_list)


def run_cli(args_list: list[str] | None = None) -> int:
    args = parse_args(args_list)

    try:
        if args.action == "execucao":
            result = executar_processo_completo(caminho_excel=args.caminho_excel, assumir_sim=True)
        elif args.action == "departamento":
            result = fluxo_departamento(
                departamento=args.departamento,
                subacao=args.subacao,
                caminho_excel=args.caminho_excel,
            )
        elif args.action == "utilizador":
            result = auditoria_utilizador(
                username=args.username,
                caminho_excel=args.caminho_excel,
            )
        elif args.action == "pesquisa":
            result = pesquisa_atribuir_transacao(
                tcode=args.tcode,
                username=args.username if args.username else None,
                simular=args.simular,
                caminho_excel=args.caminho_excel,
            )
        elif args.action == "corrigir":
            result = correcao_sincronizacao_posterior(caminho_excel=args.caminho_excel)
        elif args.action == "su53":
            result = diagnostico_su53(
                username=args.username,
                target_env=args.environment,
                tcode_atribuir=args.tcode_atribuir if args.tcode_atribuir else None,
                caminho_excel=args.caminho_excel,
            )
        else:
            result = {"ok": False, "message": f"Ação desconhecida: {args.action}"}

        print(json.dumps(result, ensure_ascii=False))
        return 0 if result.get("ok") else 1

    except Exception as exc:
        err_res = {"ok": False, "status": "ERRO", "message": str(exc)}
        print(json.dumps(err_res, ensure_ascii=False))
        return 1


def main() -> int:
    return run_cli()


if __name__ == "__main__":
    sys.exit(main())

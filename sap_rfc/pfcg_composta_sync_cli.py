# -*- coding: utf-8 -*-
"""
sap_rfc/pfcg_composta_sync_cli.py
======================================================================
CLI de Execução da Fase 4: Sincronização da Sheet PFCG_COMPOSTA
======================================================================
"""

import sys
import argparse
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from sap_rfc.pfcg_composta_sync_service import sincronizar_pfcg_composta

CAMINHO_EXCEL_PADRAO = str(PROJECT_ROOT / "S4H_Perfis de autorização_v1.xlsx")


def main():
    parser = argparse.ArgumentParser(description="Sincronizar linhas pendentes da sheet PFCG_COMPOSTA com SAP QAD e PRD.")
    parser.add_argument("--excel", type=str, default=str(PROJECT_ROOT / "S4H_Perfis de autorização_v1.xlsx"), help="Caminho do ficheiro Excel mestre.")
    parser.add_argument("--real", action="store_true", help="Executar gravação real no Excel (SAP PRD permanece READ-ONLY).")
    parser.add_argument("--only-prd", action="store_true", help="Executar sincronização apenas com SAP PRD (bypassa QAD indisponível e marca STATUS='Pendente QAD').")
    parser.add_argument("--ambiente-qad", type=str, default="QAD", help="Ambiente SAP QAD.")
    parser.add_argument("--ambiente-prd", type=str, default="PRD", help="Ambiente SAP PRD.")
    args = parser.parse_args()

    dry_run = not args.real

    print(f"\n==============================================================================")
    print(f"  FASE 4: SINCRONIZAÇÃO DE COMPOSITE ROLES (PFCG_COMPOSTA -> QAD / PRD)")
    print(f"  Ficheiro: {args.excel}")
    print(f"  Modo PRD-Only: {'SIM' if args.only_prd else 'NÃO'}")
    print(f"  Modo Execução: {'SIMULAÇÃO pure dry-run (Sem alteração no Excel)' if dry_run else 'EXECUÇÃO REAL NO EXCEL'}")
    print(f"==============================================================================\n")

    res = sincronizar_pfcg_composta(
        caminho_excel=args.excel,
        dry_run=dry_run,
        only_prd=args.only_prd,
        ambiente_qad=args.ambiente_qad,
        ambiente_prd=args.ambiente_prd
    )

    print(f"Status: {res.get('status')}")
    print(f"Mensagem: {res.get('mensagem', 'OK')}")
    print(f"Total Pendências Encontradas: {res.get('total_linhas_pendentes', 0)}")
    print(f"  - NOVO_LOTE (IDs 742-849): {res.get('total_novo_lote', 0)}")
    print(f"  - PENDENCIAS_HISTORICAS (IDs < 742): {res.get('total_historicas', 0)}")
    print(f"Total Composites Processadas: {res.get('total_composites_processadas', 0)}\n")

    if res.get("detalhes"):
        print("Detalhamento por Composite Role:")
        for det in res["detalhes"]:
            print(f"  - {det['composite']}: Status={det['status_final']} | Motivo={det.get('motivo')}")

    print(f"\n==============================================================================\n")


if __name__ == "__main__":
    main()

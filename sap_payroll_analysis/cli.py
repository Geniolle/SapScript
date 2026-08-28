"""Ponto de entrada CLI do diagnóstico Payroll -> FI (READ-ONLY).

    python -m sap_payroll_analysis --diagnostic
    python -m sap_payroll_analysis
    python -m sap_payroll_analysis --company 2010 --year 2026 --month 6 --account 23120000
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from dataclasses import replace
from pathlib import Path

from .config import DEFAULTS, AnalysisParams
from .sap_connection import (
    SapConnectionError,
    load_env,
    resolve_env_prefix,
    safe_connection_summary,
    sap_connection,
    build_connection_params,
)

logger = logging.getLogger("sap_payroll_analysis")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="sap_payroll_analysis",
        description="Diagnóstico READ-ONLY das divergências Payroll (RH) -> FI via RFC.",
    )
    p.add_argument("--diagnostic", action="store_true",
                   help="Só testa ligação, tabelas, campos e amostras. Sem análise pesada.")
    p.add_argument("--company", default=DEFAULTS.empresa, help="Código da empresa (BUKRS).")
    p.add_argument("--year", type=int, default=DEFAULTS.ano, help="Exercício.")
    p.add_argument("--month", type=int, default=DEFAULTS.mes, help="Mês / período contabilístico.")
    p.add_argument("--account", default=DEFAULTS.conta, help="Conta do Razão.")
    p.add_argument("--run", default=DEFAULTS.primary_run,
                   help="Posting run analisado nas fases 2/3 (composição por rubrica).")
    p.add_argument("--explain-cluster", action="store_true",
                   help="Fase 3: RGDIR + tentativa de leitura da RT + análise de retro. "
                        "(Já corre na análise completa; esta flag imprime só a fase 3.)")
    p.add_argument("--rt-link-diagnostic", action="store_true",
                   help="Fase 3: imprime o diagnóstico RT<->PPOIX por PERNR.")
    p.add_argument("--payroll-timeline", action="store_true",
                   help="Fase 3: imprime a timeline de resultados de Payroll (+ pares).")
    p.add_argument("--pernr", default=None,
                   help="Fase 3: restringe o diagnóstico/timeline a este PERNR.")
    p.add_argument("--manual-rt-request", action="store_true",
                   help="Gera a shortlist mínima de PERNR/SEQNR para consulta manual no PC00_M99_CWTR.")
    p.add_argument("--trace-wagetype", action="store_true",
                   help="Fase 4.1: rastreia uma rubrica (--lgart) de um PERNR (--pernr) "
                        "de PPOIX -> PPDIX -> PPDIT.")
    p.add_argument("--lgart", default=None, help="Fase 4.1: rubrica a rastrear (ex.: 0029, /559).")
    p.add_argument("--compare-lgart", default=None,
                   help="Fase 4.1: rubrica a comparar lado a lado (por omissão a 1ª de "
                        "WAGE_TYPES_REFERENCIA diferente de --lgart).")
    p.add_argument("--trace-posting-delta", action="store_true",
                   help="Fase 4.2: localiza a origem técnica do delta SUM(PPOIX) vs "
                        "PPDIT.WRBTR de uma linha de posting (exige --docnum e --doclin).")
    p.add_argument("--analyze-zero-tslin", action="store_true",
                   help="Fase 4.2: agrega todos os PPOIX do run com TSLIN = 0.")
    p.add_argument("--trace-seqno-history", action="store_true",
                   help="Fase 4.2: mapa de uso de PERNR+SEQNO do run vs outros runs.")
    p.add_argument("--docnum", default=None, help="Fase 4.2: nº do documento de posting HR (PPDIT).")
    p.add_argument("--doclin", default=None, help="Fase 4.2: nº da linha do documento (PPDIT.DOCLIN).")
    p.add_argument("--reconcile-payroll-regu", action="store_true",
                   help="Fase 5.0: reconcilia RH (/559 corrente) x programa de pagamentos "
                        "REGU* no R/3 (exige SAP_R3_* completo no .env — sem fallback).")
    p.add_argument("--period", default=None, help="Fase 5.0: período de pagamento AAAAMM (ex.: 202606).")
    p.add_argument("--payment-run-date", default=None, help="Fase 5.0: LAUFD explícito (opcional).")
    p.add_argument("--payment-run-id", default=None, help="Fase 5.0: LAUFI explícito (opcional).")
    p.add_argument("--no-cluster", action="store_true", help="Salta a fase 3 (cluster).")
    p.add_argument("--no-files", action="store_true", help="Não escrever JSON/CSV.")
    p.add_argument("--output-dir", default=None, help="Directório de output (por omissão ./output).")
    p.add_argument("-v", "--verbose", action="store_true", help="Log ao nível DEBUG.")
    return p


def configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def _slug(text: str) -> str:
    return "".join(ch if ch.isalnum() else "" for ch in str(text)) or "x"


def _norm_lgart(value: str | None) -> str | None:
    """Recupera rubricas '/NNN' que o Git-Bash converteu em caminho Windows
    (ex.: 'C:/.../Git/559' -> '/559'). Use MSYS_NO_PATHCONV=1 para evitar."""
    if not value:
        return value
    v = str(value).strip()
    if ("/" in v or "\\" in v or ":" in v) and v.replace("/", "\\").split("\\")[-1].isdigit():
        return "/" + v.replace("/", "\\").split("\\")[-1]
    return v


def params_from_args(args: argparse.Namespace) -> AnalysisParams:
    return replace(
        DEFAULTS,
        empresa=str(args.company).strip(),
        ano=int(args.year),
        mes=int(args.month),
        conta=str(args.account).strip(),
        primary_run=str(args.run).strip(),
    )


def _run_reconcile_payroll_regu(args, params, output_dir: Path) -> int:
    """Fase 5.0 — reconciliação Payroll x REGU*. Só R/3 (`SAP_R3_*`), sem fallback."""
    from .sap_connection import require_prefix

    try:
        load_env()
    except SapConnectionError as exc:
        print(f"ERROR:\n{exc}")
        return 2
    try:
        prefix = require_prefix("SAP_R3_", purpose="payroll/payment reconciliation")
    except SapConnectionError as exc:
        print("ERROR:")
        print("SAP_R3_* connection parameters required for payroll/payment")
        print("reconciliation.")
        logger.error("%s", exc)
        return 2

    period = str(args.period or f"{params.ano:04d}{params.mes:02d}").strip()
    run10 = params.primary_run_10
    try:
        conn_summary = safe_connection_summary(build_connection_params(prefix))
        logger.info("Ligação RFC (R/3): %s", conn_summary)
        with sap_connection(prefix) as conn:
            from .payment_reconciliation import (
                print_reconciliation_report,
                reconcile_payroll_payments,
                write_reconciliation_csv,
                write_reconciliation_json,
            )

            recon = reconcile_payroll_payments(
                conn, params, run=run10, company=str(args.company).strip(),
                period=period, payment_run_date=args.payment_run_date,
                payment_run_id=args.payment_run_id)
            print_reconciliation_report(recon)
            if not args.no_files:
                stem = run10
                write_reconciliation_json(
                    recon, output_dir / f"payroll_regu_reconciliation_{stem}.json")
                write_reconciliation_csv(
                    recon, output_dir / f"payroll_regu_reconciliation_{stem}.csv")
            return 0
    except SapConnectionError as exc:
        print(f"ERROR:\n{exc}")
        return 2
    except Exception as exc:  # noqa: BLE001
        logger.exception("Falha na reconciliação Payroll x REGU: %s", exc)
        return 1


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    configure_logging(args.verbose)

    params = params_from_args(args)
    output_dir = Path(args.output_dir) if args.output_dir else Path("output")

    if args.manual_rt_request:
        from .manual_request import build_manual_rt_shortlist, print_manual_rt_report

        try:
            res = build_manual_rt_shortlist(params, output_dir=output_dir, write_files=not args.no_files)
            print_manual_rt_report(res)
            return 0
        except Exception as local_exc:
            logger.debug("Tentando com ligação SAP RFC: %s", local_exc)

    if args.reconcile_payroll_regu:
        return _run_reconcile_payroll_regu(args, params, output_dir)

    try:
        load_env()
        prefix = resolve_env_prefix()
        conn_summary = safe_connection_summary(build_connection_params(prefix))
    except SapConnectionError as exc:
        if args.manual_rt_request:
            logger.error("Não foi possível gerar manual-rt-request (sem dados locais e sem RFC): %s", exc)
            return 1
        logger.error("Configuração RFC: %s", exc)
        return 2

    logger.info("Ligação RFC: %s", conn_summary)

    try:
        with sap_connection(prefix) as conn:
            if args.manual_rt_request:
                from .manual_request import build_manual_rt_shortlist, print_manual_rt_report

                res = build_manual_rt_shortlist(params, output_dir=output_dir, connection=conn, write_files=not args.no_files)
                print_manual_rt_report(res)
                return 0

            if args.trace_posting_delta or args.analyze_zero_tslin or args.trace_seqno_history:
                from .posting_delta_trace import (
                    analyze_zero_tslin_standalone,
                    print_posting_delta_report,
                    trace_posting_delta,
                    trace_seqno_history_standalone,
                    write_posting_delta_csvs,
                    write_posting_delta_json,
                )

                run10 = params.primary_run_10
                if args.analyze_zero_tslin and not args.trace_posting_delta:
                    res = analyze_zero_tslin_standalone(conn, params, run10)
                    print(json.dumps(res, indent=2, ensure_ascii=False))
                    if not args.no_files:
                        (output_dir).mkdir(parents=True, exist_ok=True)
                        (output_dir / f"zero_tslin_{run10}.json").write_text(
                            json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
                    return 0
                if args.trace_seqno_history and not args.trace_posting_delta:
                    res = trace_seqno_history_standalone(conn, params, run10)
                    print(json.dumps(res, indent=2, ensure_ascii=False))
                    if not args.no_files:
                        (output_dir).mkdir(parents=True, exist_ok=True)
                        (output_dir / f"seqno_history_{run10}.json").write_text(
                            json.dumps(res, indent=2, ensure_ascii=False), encoding="utf-8")
                    return 0
                if not (args.docnum and args.doclin):
                    logger.error("--trace-posting-delta exige --docnum e --doclin.")
                    return 2
                tr = trace_posting_delta(conn, params, docnum=args.docnum,
                                         doclin=args.doclin, run=run10)
                print_posting_delta_report(tr)
                if not args.no_files:
                    stem = f"{run10}_{int(str(args.docnum).strip())}_{int(str(args.doclin).strip())}"
                    write_posting_delta_json(tr, output_dir / f"posting_delta_{stem}.json")
                    write_posting_delta_csvs(tr, output_dir, run10,
                                             str(args.docnum).strip().zfill(10),
                                             str(args.doclin).strip().zfill(10))
                return 0

            if args.trace_wagetype:
                from .config import WAGE_TYPES_REFERENCIA
                from .wagetype_trace import (
                    print_trace_report, trace_wagetype, write_trace_csv, write_trace_json,
                )

                lg = _norm_lgart(args.lgart)
                if not (args.pernr and lg):
                    logger.error("--trace-wagetype exige --pernr e --lgart.")
                    return 2
                cmp = _norm_lgart(args.compare_lgart)
                if cmp is None:
                    prefer = ["/559", *WAGE_TYPES_REFERENCIA]
                    cmp = next((w for w in prefer if w != lg), None)
                tr = trace_wagetype(conn, params, pernr=str(args.pernr).strip(),
                                    lgart=lg, compare_lgart=cmp)
                print_trace_report(tr)
                if not args.no_files:
                    slug = f"{params.primary_run_10}_{str(args.pernr).strip().zfill(8)}_{_slug(lg)}"
                    write_trace_json(tr, output_dir / f"trace_{slug}.json")
                    write_trace_csv(tr, output_dir / f"trace_{slug}.csv")
                return 0

            if args.diagnostic:
                from .diagnostics import run_diagnostic

                run_diagnostic(conn, params)
                return 0

            if args.explain_cluster or args.rt_link_diagnostic or args.payroll_timeline:
                from .cluster_cli import run_cluster_only

                run_cluster_only(conn, params, pernr=args.pernr,
                                 rt_link_diagnostic=args.rt_link_diagnostic,
                                 payroll_timeline=args.payroll_timeline)
                return 0

            from .analysis import run_analysis

            output_dir = Path(args.output_dir) if args.output_dir else None
            run_analysis(
                conn,
                params,
                connection_info=conn_summary,
                output_dir=output_dir,
                write_files=not args.no_files,
                with_cluster=not args.no_cluster,
            )
            return 0
    except SapConnectionError as exc:
        logger.error("Ligação RFC falhou: %s", exc)
        return 2
    except KeyboardInterrupt:  # pragma: no cover
        logger.warning("Interrompido pelo utilizador.")
        return 130
    except Exception as exc:  # noqa: BLE001
        logger.exception("Falha inesperada: %s", exc)
        return 1


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())

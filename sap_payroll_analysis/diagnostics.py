"""Modo `--diagnostic`: liga, verifica tabelas, lista campos, lê amostras.

Não executa análise pesada. Serve para validar acesso e descobrir os campos
reais antes da reconciliação.
"""

from __future__ import annotations

import logging
from typing import Any

from .config import AnalysisParams
from .ddic import CONCEPT_KEYWORDS, describe_table, guess_fields
from .models import TableDiag
from .payroll_posting import PAYROLL_TABLES
from .sap_reader import RfcReadError, opt_in, read_table

logger = logging.getLogger(__name__)

DIAG_TABLES: tuple[str, ...] = (
    *PAYROLL_TABLES,
    "ACDOCA",
    "BKPF",
    "BSEG",
    "BSIS",
    "BSAS",
)

RELEVANT_CONCEPTS: tuple[str, ...] = (
    "empresa", "conta", "valor", "moeda", "posting_run",
    "documento", "item", "exercicio", "periodo", "debito_credito",
)


def run_diagnostic(connection: Any, params: AnalysisParams) -> dict[str, Any]:
    report: dict[str, Any] = {"tables": {}, "field_guesses": {}, "posting_run_probe": {}}

    print("=" * 60)
    print(" DIAGNÓSTICO SAP  (READ-ONLY)")
    print("=" * 60)
    print()
    print(f"{'Tabela':<10} {'Existe':<7} {'Autoriz.':<9} {'Nº campos':<10} Observação")
    print("-" * 60)

    diags: dict[str, TableDiag] = {}
    for table in DIAG_TABLES:
        try:
            diag = describe_table(connection, table, sample=2)
        except RfcReadError as exc:
            diag = TableDiag(table=table, note=f"{exc.kind}: {exc}")
        diags[table] = diag
        report["tables"][table] = diag.as_dict()
        print(
            f"{table:<10} {_yn(diag.exists):<7} {_yn(diag.authorized):<9} "
            f"{diag.field_count:<10} {diag.note or 'OK'}"
        )

    print()
    print("-" * 60)
    print("CAMPOS CANDIDATOS POR CONCEITO")
    print("-" * 60)
    for table, diag in diags.items():
        if not (diag.exists and diag.fields):
            continue
        guesses = guess_fields(diag, RELEVANT_CONCEPTS)
        report["field_guesses"][table] = {c: g.as_dict() for c, g in guesses.items()}
        print(f"\n[{table}]  ({diag.field_count} campos)")
        for concept in RELEVANT_CONCEPTS:
            g = guesses.get(concept)
            cands = ", ".join(g.candidates[:6]) if g and g.candidates else "—"
            print(f"  {concept:<16}: {cands}")

    # Sonda dos posting runs: tenta localizar os números nas tabelas de header.
    print()
    print("-" * 60)
    print("SONDA POSTING RUNS PCP0")
    print("-" * 60)
    probe = _probe_posting_runs(connection, params, diags)
    report["posting_run_probe"] = probe
    for table, info in probe.items():
        print(f"  {table}: {info}")

    print()
    print("Diagnóstico concluído. Se o acesso estiver OK, executar sem --diagnostic.")
    return report


def _probe_posting_runs(
    connection: Any, params: AnalysisParams, diags: dict[str, TableDiag]
) -> dict[str, Any]:
    runs = [r.zfill(10) for r in params.posting_runs]
    out: dict[str, Any] = {}
    for table in ("PPDHD", "PPDIT", "PPDIX", "PPOIX", "PEVST"):
        diag = diags.get(table)
        if not diag or not (diag.exists and diag.authorized and diag.fields):
            out[table] = "indisponível"
            continue
        guesses = guess_fields(diag, ["posting_run"])
        field = guesses["posting_run"].chosen
        if not field:
            out[table] = (
                "sem campo de run; ligar por DOCNUM via PPDHD/PPDIX. "
                f"campos: {diag.field_names()[:10]}"
            )
            continue
        try:
            rows = read_table(
                connection, table, fields=[field], options=opt_in(field, runs), max_rows=50
            ).rows
            found = sorted({r.get(field, "").strip() for r in rows})
            out[table] = {"campo": field, "linhas": len(rows), "runs_vistos": found}
        except RfcReadError as exc:
            out[table] = f"{field}: {exc.kind}"
    return out


def _yn(value: bool) -> str:
    return "sim" if value else "não"

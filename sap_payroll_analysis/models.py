"""Estruturas de dados do diagnóstico (dataclasses)."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from typing import Any


@dataclass
class FieldInfo:
    """Metadados de um campo de tabela (via RFC_READ_TABLE / DDIC)."""

    name: str
    datatype: str = ""
    length: int = 0
    offset: int = 0
    description: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "datatype": self.datatype,
            "length": self.length,
            "offset": self.offset,
            "description": self.description,
        }


@dataclass
class TableDiag:
    """Resultado do diagnóstico de uma tabela."""

    table: str
    exists: bool = False
    authorized: bool = False
    fields: list[FieldInfo] = field(default_factory=list)
    sample_rows: list[dict[str, str]] = field(default_factory=list)
    note: str = ""

    @property
    def field_count(self) -> int:
        return len(self.fields)

    def field_names(self) -> list[str]:
        return [f.name for f in self.fields]

    def as_dict(self) -> dict[str, Any]:
        return {
            "table": self.table,
            "exists": self.exists,
            "authorized": self.authorized,
            "field_count": self.field_count,
            "fields": [f.as_dict() for f in self.fields],
            "sample_rows": self.sample_rows,
            "note": self.note,
        }


@dataclass
class FieldGuess:
    """Candidatos de campo para um conceito semântico (ex.: 'conta')."""

    concept: str
    candidates: list[str] = field(default_factory=list)
    chosen: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return {"concept": self.concept, "candidates": self.candidates, "chosen": self.chosen}


@dataclass
class PostingItem:
    """Item de posting do Payroll (linha de PPDIT / equivalente)."""

    run_id: str
    doc_number: str = ""
    line: str = ""
    account: str = ""
    company: str = ""
    currency: str = ""
    debit_credit: str = ""
    amount_raw: str = ""
    amount: Decimal = Decimal("0")
    signed_amount: Decimal = Decimal("0")
    raw: dict[str, str] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "doc_number": self.doc_number,
            "line": self.line,
            "account": self.account,
            "company": self.company,
            "currency": self.currency,
            "debit_credit": self.debit_credit,
            "amount": str(self.amount),
            "signed_amount": str(self.signed_amount),
        }


@dataclass
class FIItem:
    """Linha contabilística de FI (ACDOCA ou BKPF/BSEG)."""

    source: str  # "ACDOCA" | "BSEG"
    document: str = ""
    fiscal_year: str = ""
    line: str = ""
    posting_date: str = ""
    period: str = ""
    account: str = ""
    company: str = ""
    currency: str = ""
    debit_credit: str = ""
    amount_raw: str = ""
    amount: Decimal = Decimal("0")
    signed_amount: Decimal = Decimal("0")
    doc_type: str = ""
    reference: str = ""
    text: str = ""
    raw: dict[str, str] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "document": self.document,
            "fiscal_year": self.fiscal_year,
            "line": self.line,
            "posting_date": self.posting_date,
            "period": self.period,
            "account": self.account,
            "company": self.company,
            "currency": self.currency,
            "debit_credit": self.debit_credit,
            "amount": str(self.amount),
            "signed_amount": str(self.signed_amount),
            "doc_type": self.doc_type,
            "reference": self.reference,
            "text": self.text,
        }


@dataclass
class ReconLine:
    label: str
    left: Decimal | None
    right: Decimal | None
    diff: Decimal | None
    status: str  # OK | DIVERGENTE | INDETERMINADO

    def as_dict(self) -> dict[str, Any]:
        return {
            "label": self.label,
            "left": None if self.left is None else str(self.left),
            "right": None if self.right is None else str(self.right),
            "diff": None if self.diff is None else str(self.diff),
            "status": self.status,
        }


@dataclass
class AnalysisResult:
    """Contentor de todo o resultado do diagnóstico (serializável para JSON)."""

    parameters: dict[str, Any] = field(default_factory=dict)
    reference_values: dict[str, Any] = field(default_factory=dict)
    connection: dict[str, Any] = field(default_factory=dict)
    tables: dict[str, Any] = field(default_factory=dict)
    field_guesses: dict[str, Any] = field(default_factory=dict)
    posting_runs: list[dict[str, Any]] = field(default_factory=list)
    payroll_posting: dict[str, Any] = field(default_factory=dict)
    fi: dict[str, Any] = field(default_factory=dict)
    wage_link: dict[str, Any] = field(default_factory=dict)
    payroll_cluster: dict[str, Any] = field(default_factory=dict)
    reconciliation: list[dict[str, Any]] = field(default_factory=list)
    next_steps: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def add_warning(self, message: str) -> None:
        if message not in self.warnings:
            self.warnings.append(message)

    def add_error(self, message: str) -> None:
        if message not in self.errors:
            self.errors.append(message)

    def as_dict(self) -> dict[str, Any]:
        return {
            "parameters": self.parameters,
            "reference_values": self.reference_values,
            "connection": self.connection,
            "tables": self.tables,
            "field_guesses": self.field_guesses,
            "posting_runs": self.posting_runs,
            "payroll_posting": self.payroll_posting,
            "fi": self.fi,
            "wage_link": self.wage_link,
            "payroll_cluster": self.payroll_cluster,
            "reconciliation": self.reconciliation,
            "next_steps": self.next_steps,
            "warnings": self.warnings,
            "errors": self.errors,
        }

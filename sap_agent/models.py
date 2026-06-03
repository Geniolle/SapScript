from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class TicketContext:
    key: str
    summary: str
    description: str
    comments: list[str] = field(default_factory=list)
    labels: list[str] = field(default_factory=list)
    components: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)
    attachment_texts: list[str] = field(default_factory=list)
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def full_text(self) -> str:
        return "\n".join([self.summary, self.description, *self.comments])


@dataclass
class SapErrorSignal:
    transaction: str | None = None
    transaction_description: str | None = None
    program: str | None = None
    program_description: str | None = None
    message_id: str | None = None
    message_number: str | None = None
    company_code: str | None = None
    document_number: str | None = None
    fiscal_year: str | None = None
    job_name: str | None = None
    user: str | None = None
    iban: str | None = None
    keyno: str | None = None
    module: str | None = None


@dataclass
class ValidationEvidence:
    name: str
    status: str
    details: str
    data: list[dict[str, Any]] = field(default_factory=list)


@dataclass
class DiagnosisResult:
    ticket_key: str
    signal: SapErrorSignal
    evidences: list[ValidationEvidence]
    probable_cause: str
    proposed_solution: str
    tests_to_execute: list[str]
    confidence: str = "baixa"
    ticket_attachments: list[str] = field(default_factory=list)
    ticket_attachment_texts: list[str] = field(default_factory=list)

    def to_jira_comment(self, prefix: str = "Pré-análise automática SAP") -> str:
        evidence_text = "\n".join(
            f"- *{evidence.name}* [{evidence.status}]: {evidence.details}" for evidence in self.evidences
        ) or "- Nenhuma evidência técnica recolhida."
        tests = "\n".join(f"- {test}" for test in self.tests_to_execute) or "- Não definido."
        signal = self.signal
        t_desc = f" ({signal.transaction_description})" if signal.transaction_description else ""
        p_desc = f" ({signal.program_description})" if signal.program_description else ""
        return f"""h3. {prefix}

*Ticket:* {self.ticket_key}
*Confiança:* {self.confidence}

h4. Sinais identificados
- Transação: {signal.transaction or 'não identificada'}{t_desc}
- Programa/Classe: {signal.program or 'não identificado'}{p_desc}
- Mensagem SAP: {(signal.message_id or '')} {(signal.message_number or '')}
- Empresa: {signal.company_code or 'não identificada'}
- Documento: {signal.document_number or 'não identificado'}
- Exercício: {signal.fiscal_year or 'não identificado'}
- Job: {signal.job_name or 'não identificado'}
- Utilizador: {signal.user or 'não identificado'}

h4. Evidências de leitura SAP
{evidence_text}

h4. Possível causa
{self.probable_cause}

h4. Prévia de solução
{self.proposed_solution}

h4. Testes/validações sugeridos
{tests}

_Observação: análise gerada em modo somente leitura. Nenhuma alteração foi executada no SAP._
"""

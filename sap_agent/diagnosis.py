from __future__ import annotations

from .extractor import extract_signal
from .models import DiagnosisResult, SapErrorSignal, TicketContext, ValidationEvidence
from .validators import SapReadOnlyValidator


class SapDiagnosisEngine:
    def __init__(self, validator: SapReadOnlyValidator):
        self.validator = validator

    def diagnose(self, ticket: TicketContext) -> DiagnosisResult:
        signal = extract_signal(ticket)
        
        if signal.transaction and not signal.program:
            program = self.validator.get_program_for_transaction(signal.transaction)
            if program:
                signal.program = program

        evidences = self.validator.validate(ticket, signal)
        probable_cause = self._probable_cause(signal, evidences)
        proposed_solution = self._proposed_solution(signal, evidences)
        tests = self._tests(signal)
        confidence = self._confidence(signal, evidences)
        return DiagnosisResult(
            ticket_key=ticket.key,
            signal=signal,
            evidences=evidences,
            probable_cause=probable_cause,
            proposed_solution=proposed_solution,
            tests_to_execute=tests,
            confidence=confidence,
            ticket_attachments=ticket.attachments,
            ticket_attachment_texts=ticket.attachment_texts,
        )

    def _probable_cause(self, signal: SapErrorSignal, evidences: list[ValidationEvidence]) -> str:
        evidence_text = " ".join(e.details for e in evidences).upper()
        if signal.module == "FI" and (signal.transaction or "").startswith("FIBL"):
            return (
                "O erro parece relacionado com processo FI de ordem de pagamento, banco empresa, método de pagamento, "
                "IBAN ou conta de compensação. Validar PAYRQ, banco empresa/conta bancária e configuração do método de pagamento."
            )
        if "DOCUMENTO NÃO ENCONTRADO" in evidence_text or "DOCUMENTO NAO ENCONTRADO" in evidence_text:
            return "O documento informado no ticket não foi encontrado com os critérios disponíveis. Pode haver erro no número, empresa, exercício ou ambiente."
        if "MENSAGEM NÃO ENCONTRADA" in evidence_text or "MENSAGEM NAO ENCONTRADA" in evidence_text:
            return "A mensagem SAP indicada não foi localizada na T100. Pode ser mensagem dinâmica, texto incompleto ou erro de extração do ticket."
        if signal.job_name:
            return "O erro pode estar relacionado com processamento em background. Validar status do job, spool e logs associados."
        return "Causa preliminar não conclusiva. Foram recolhidas evidências em modo leitura, mas o ticket precisa de mais detalhe técnico para diagnóstico preciso."

    def _proposed_solution(self, signal: SapErrorSignal, evidences: list[ValidationEvidence]) -> str:
        actions = [
            "Confirmar se o erro ocorre no mesmo ambiente informado no ticket.",
            "Validar os dados de entrada usados pelo utilizador e comparar com as evidências recolhidas em SAP.",
        ]
        if signal.module == "FI":
            actions.extend([
                "Validar documento FI na BKPF/BSEG quando houver empresa, documento e exercício.",
                "Validar ordem de pagamento na PAYRQ quando houver empresa, chave ou IBAN.",
                "Confirmar banco empresa, conta bancária e método de pagamento antes de sugerir customizing.",
            ])
        if signal.message_id and signal.message_number:
            actions.append("Usar o texto da mensagem SAP T100 para pesquisar SAP Notes/KBA e histórico interno de incidentes.")
        if signal.job_name:
            actions.append("Validar SM37, spool e logs do job em background. O agente apenas consulta metadados por tabela; o consultor deve confirmar o detalhe funcional.")
        actions.append("Não executar correção automática. Abrir tarefa técnica/funcional para validação humana se a causa for confirmada.")
        return "\n".join(f"{index}. {action}" for index, action in enumerate(actions, start=1))

    def _tests(self, signal: SapErrorSignal) -> list[str]:
        tests = [
            "Reproduzir o cenário em QAS com o mesmo utilizador/perfil ou perfil equivalente.",
            "Executar novamente com dados mínimos válidos e registar mensagem completa.",
        ]
        if signal.module == "FI":
            tests.extend([
                "Validar cenário com documento FI existente e inexistente.",
                "Validar dados de banco empresa, método de pagamento e IBAN.",
                "Validar se existe duplicidade de ordem de pagamento quando o cenário envolver PAYRQ.",
            ])
        if signal.job_name:
            tests.append("Executar teste controlado do job em QAS e comparar status/log antes e depois.")
        return tests

    def _confidence(self, signal: SapErrorSignal, evidences: list[ValidationEvidence]) -> str:
        ok_count = sum(1 for evidence in evidences if evidence.status == "ok")
        extracted = sum(
            bool(value)
            for value in (
                signal.transaction,
                signal.message_id,
                signal.message_number,
                signal.company_code,
                signal.document_number,
                signal.fiscal_year,
                signal.job_name,
                signal.iban,
            )
        )
        if ok_count >= 3 and extracted >= 4:
            return "alta"
        if ok_count >= 2 or extracted >= 3:
            return "média"
        return "baixa"

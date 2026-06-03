from __future__ import annotations

from .models import SapErrorSignal, TicketContext, ValidationEvidence
from .sap_rfc_client import SapRfcClient


class SapReadOnlyValidator:
    def __init__(self, sap_client: SapRfcClient):
        self.sap = sap_client

    def get_program_for_transaction(self, tcode: str) -> str | None:
        try:
            rows = self.sap.read_table(
                "TSTC",
                fields=["TCODE", "PGMNA"],
                options=[f"TCODE = '{tcode}'"],
                rowcount=1,
            )
            if rows:
                return rows[0].get("PGMNA")
        except Exception:
            pass
        return None

    def enrich_signal(self, signal: SapErrorSignal) -> None:
        if signal.transaction and not signal.transaction_description:
            try:
                rows = self.sap.read_table(
                    "TSTCT",
                    fields=["TTEXT"],
                    options=[f"TCODE = '{signal.transaction}'"],
                    rowcount=1,
                )
                if rows:
                    signal.transaction_description = rows[0].get("TTEXT")
            except Exception:
                pass

        if signal.program and not signal.program_description:
            try:
                rows = self.sap.read_table(
                    "TRDIRT",
                    fields=["TEXT"],
                    options=[f"NAME = '{signal.program}'"],
                    rowcount=1,
                )
                if rows:
                    signal.program_description = rows[0].get("TEXT")
            except Exception:
                pass

    def validate(self, ticket: TicketContext, signal: SapErrorSignal) -> list[ValidationEvidence]:
        self.enrich_signal(signal)
        evidences: list[ValidationEvidence] = []
        evidences.append(self._validate_connection())

        if signal.message_id and signal.message_number:
            evidences.append(self._validate_message(signal))
        if signal.company_code and signal.document_number and signal.fiscal_year:
            evidences.append(self._validate_fi_document(signal))
        if signal.company_code or signal.keyno or signal.iban:
            evidences.append(self._validate_payment_request(signal))
        if signal.job_name:
            evidences.append(self._validate_background_job(signal))

        if len(evidences) == 1:
            evidences.append(
                ValidationEvidence(
                    name="Extração de sinais SAP",
                    status="aviso",
                    details="O ticket não possui dados técnicos suficientes para executar validações específicas em SAP.",
                )
            )
        return evidences

    def _validate_connection(self) -> ValidationEvidence:
        try:
            self.sap.ping()
            return ValidationEvidence("Conexão RFC", "ok", "Conexão SAP RFC respondendo em modo leitura.")
        except Exception as exc:
            return ValidationEvidence("Conexão RFC", "erro", f"Não foi possível validar conexão SAP: {exc}")

    def _validate_message(self, signal: SapErrorSignal) -> ValidationEvidence:
        try:
            rows = self.sap.get_message_text(signal.message_id or "", signal.message_number or "")
            if rows:
                text = rows[0].get("TEXT", "")
                return ValidationEvidence("Mensagem SAP T100", "ok", f"Mensagem encontrada: {text}", rows)
            return ValidationEvidence("Mensagem SAP T100", "aviso", "Mensagem não encontrada na T100.")
        except Exception as exc:
            return ValidationEvidence("Mensagem SAP T100", "erro", str(exc))

    def _validate_fi_document(self, signal: SapErrorSignal) -> ValidationEvidence:
        try:
            rows = self.sap.get_fi_document_header(
                signal.company_code or "",
                signal.document_number or "",
                signal.fiscal_year or "",
            )
            if rows:
                row = rows[0]
                return ValidationEvidence(
                    "Documento FI BKPF",
                    "ok",
                    f"Documento localizado. Tipo {row.get('BLART')}, transação {row.get('TCODE')}, utilizador {row.get('USNAM')}.",
                    rows,
                )
            return ValidationEvidence("Documento FI BKPF", "aviso", "Documento não encontrado na BKPF.")
        except Exception as exc:
            return ValidationEvidence("Documento FI BKPF", "erro", str(exc))

    def _validate_payment_request(self, signal: SapErrorSignal) -> ValidationEvidence:
        try:
            rows = self.sap.get_payment_request(signal.company_code, signal.keyno, signal.iban)
            if rows:
                return ValidationEvidence(
                    "Ordem de pagamento PAYRQ",
                    "ok",
                    f"Foram encontradas {len(rows)} ordem(ns) de pagamento compatíveis.",
                    rows,
                )
            return ValidationEvidence("Ordem de pagamento PAYRQ", "aviso", "Nenhuma ordem de pagamento encontrada na PAYRQ para os critérios disponíveis.")
        except Exception as exc:
            return ValidationEvidence("Ordem de pagamento PAYRQ", "erro", str(exc))

    def _validate_background_job(self, signal: SapErrorSignal) -> ValidationEvidence:
        try:
            rows = self.sap.get_background_jobs(signal.job_name or "", signal.user)
            if rows:
                return ValidationEvidence(
                    "Job background TBTCO",
                    "ok",
                    f"Foram encontrados {len(rows)} job(s) compatíveis.",
                    rows,
                )
            return ValidationEvidence("Job background TBTCO", "aviso", "Nenhum job encontrado na TBTCO para os critérios disponíveis.")
        except Exception as exc:
            return ValidationEvidence("Job background TBTCO", "erro", str(exc))

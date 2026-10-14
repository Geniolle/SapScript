"""
Salsa IT Agent - Phase 1 Orchestrator.
Coordinates:
- Ticket eligibility verification (Jira project=IZ, assignee=Clayton Lopes, category=CA Acessos, open status).
- Idempotent analysis using content hash comparison.
- Level 1 deterministic execution.
- Level 2 AI execution when needed and authorized.
- Safe persistence via store.py.
- STRICT SAFETY: ZERO SAP calls, ZERO Jira writes.
"""
from __future__ import annotations

import os
from typing import Any, Dict, Optional, Tuple

from web_api.jira_client import fetch_ticket_details
from web_api.store import (
    get_connection,
    get_salsa_ticket_analysis,
    get_salsa_ticket_analysis_history,
    save_salsa_ticket_analysis,
)

from .ai_provider import BaseAIProvider, get_ai_provider
from .classifier import RULES_VERSION, SalsaLevel1Classifier
from .models import (
    AnalysisLevel,
    AnalysisResult,
    AnalysisStatus,
    ClassificationCode,
    OperationalStatus,
)

# Eligible criteria defaults (configurable via environment variables)
ELIGIBLE_PROJECT_DEFAULT = "IZ"
ELIGIBLE_ASSIGNEE_DEFAULT = "Clayton Lopes"
ELIGIBLE_ASSIGNEE_ID_DEFAULT = "601a7387332cbe007015180a"
ELIGIBLE_PROCESS_DEFAULT = "CA Acessos"
CLOSED_STATUSES = {"done", "closed", "concluído", "resolvido", "fechado", "fechada", "cancelled"}


class SalsaAgentOrchestrator:
    """Orchestrator for Salsa IT Agent ticket analysis."""

    def __init__(self, ai_provider: Optional[BaseAIProvider] = None):
        self.classifier = SalsaLevel1Classifier()
        self.ai_provider = ai_provider or get_ai_provider()

    @staticmethod
    def get_eligible_config() -> dict[str, str]:
        return {
            "project": os.getenv("SALSA_AGENT_PROJECT", ELIGIBLE_PROJECT_DEFAULT).strip(),
            "assignee": os.getenv("SALSA_AGENT_ASSIGNEE", ELIGIBLE_ASSIGNEE_DEFAULT).strip(),
            "assignee_id": os.getenv("SALSA_AGENT_ASSIGNEE_ID", ELIGIBLE_ASSIGNEE_ID_DEFAULT).strip(),
            "process": os.getenv("SALSA_AGENT_PROCESS", ELIGIBLE_PROCESS_DEFAULT).strip(),
        }

    def check_eligibility(self, ticket: dict[str, Any]) -> Tuple[bool, str]:
        """
        Determines whether a ticket belongs to the Salsa IT Agent scope.
        Validates project, assignee, process/category and status.
        """
        cfg = self.get_eligible_config()

        # 1. Project check
        ticket_project = str(ticket.get("project") or "").strip().upper()
        ticket_key = str(ticket.get("key") or "").strip().upper()
        # In Jira, ticket key starts with PROJECT- (e.g. IZ-12345)
        is_iz_project = (
            ticket_project == cfg["project"].upper()
            or ticket_project.startswith("IT - SALSA")
            or ticket_key.startswith(f"{cfg['project'].upper()}-")
        )
        if not is_iz_project:
            return False, f"Projeto ({ticket_project or ticket_key}) fora do âmbito do agente ({cfg['project']})"

        # 2. Assignee check (supports display name or account ID)
        assignee = str(ticket.get("assignee") or "").strip()
        assignee_id = str(ticket.get("assignee_id") or ticket.get("accountId") or "").strip()
        is_clayton = (
            cfg["assignee"].lower() in assignee.lower()
            or (assignee_id and assignee_id == cfg["assignee_id"])
        )
        if not is_clayton:
            return False, f"Responsável ({assignee or 'vazio'}) não corresponde a {cfg['assignee']}"

        # 3. Process / Category check
        process = str(ticket.get("process") or ticket.get("categoria_sap") or "").strip()
        # Customfield value or presented name
        is_ca_acessos = (
            cfg["process"].lower() in process.lower()
            or "ca acessos" in process.lower()
            or process == "28790"
        )
        if not is_ca_acessos:
            return False, f"Categoria/Processo ({process}) não corresponde a {cfg['process']}"

        # 4. Status check
        status = str(ticket.get("status") or "").strip().lower()
        if status in CLOSED_STATUSES:
            return False, f"Ticket em estado concluído ({status})"

        return True, "Elegível para o Agente Salsa IT"

    def analyze_ticket_by_key(self, ticket_key: str, force_reanalyze: bool = False) -> AnalysisResult:
        """
        Fetches ticket details, checks eligibility and runs analysis with idempotency.
        """
        clean_key = ticket_key.strip().upper()

        # 1. Retrieve ticket from local DB first to check metadata
        ticket_record = None
        with get_connection() as conn:
            cur = conn.execute("SELECT * FROM jira_tickets WHERE key = ?", (clean_key,))
            row = cur.fetchone()
            if row:
                ticket_record = dict(row)

        # 2. Fetch full details (summary, description, comments, categoria_sap) via jira_client
        try:
            details = fetch_ticket_details(clean_key)
        except Exception as exc:
            # If Jira fetch fails, but we have local record, use local details
            if ticket_record:
                details = {
                    "summary": ticket_record.get("summary") or "",
                    "description": "",
                    "comments": [],
                    "categoria_sap": ticket_record.get("process") or "",
                }
            else:
                return AnalysisResult(
                    ticket_key=clean_key,
                    classification=ClassificationCode.INSUFFICIENT_INFORMATION,
                    level=AnalysisLevel.LEVEL_1,
                    status=AnalysisStatus.ERROR,
                    intent="Erro ao consultar detalhes no Jira",
                    justification=f"Falha de comunicação com o Jira: {str(exc)}",
                    error_message=str(exc),
                    requires_human_review=True,
                )

        summary = str(details.get("summary") or (ticket_record.get("summary") if ticket_record else "") or "")
        description = str(details.get("description") or "")
        process_category = str(details.get("categoria_sap") or (ticket_record.get("process") if ticket_record else "") or "")
        assignee = str(ticket_record.get("assignee") if ticket_record else ELIGIBLE_ASSIGNEE_DEFAULT)
        project = str(ticket_record.get("project") if ticket_record else ELIGIBLE_PROJECT_DEFAULT)
        status = str(details.get("status") or (ticket_record.get("status") if ticket_record else "Open"))
        requester = str(details.get("reporter") or (ticket_record.get("creator") if ticket_record else "") or "")
        comments = list(details.get("comments") or details.get("structured_comments") or [])

        ticket_context = {
            "key": clean_key,
            "summary": summary,
            "description": description,
            "project": project,
            "assignee": assignee,
            "process": process_category,
            "status": status,
        }

        # 3. Backend eligibility validation
        is_eligible, reason = self.check_eligibility(ticket_context)
        if not is_eligible:
            # Not eligible
            return AnalysisResult(
                ticket_key=clean_key,
                classification=ClassificationCode.INSUFFICIENT_INFORMATION,
                level=AnalysisLevel.LEVEL_1,
                status=AnalysisStatus.REVISION_REQUIRED,
                intent="Ticket fora do âmbito de elegibilidade",
                justification=reason,
                suggested_next_action="Rever filtros de atribuição ou categoria no Jira.",
                requires_human_review=True,
                content_hash=self.classifier.compute_content_hash(summary, description, process_category),
            )

        # 4. Check idempotency (unless forced)
        content_hash = self.classifier.compute_content_hash(summary, description, process_category)
        if not force_reanalyze:
            cached = get_salsa_ticket_analysis(clean_key)
            if cached:
                cached_hash = cached.get("content_hash")
                cached_rules = cached.get("rules_version")
                # If content hash and rules version are identical, return cached analysis
                if cached_hash == content_hash and cached_rules == RULES_VERSION:
                    return AnalysisResult.from_dict(cached)

        # 5. Run Level 1 deterministic analysis
        l1_result = self.classifier.analyze(
            ticket_key=clean_key,
            summary=summary,
            description=description,
            process_category=process_category,
            requester=requester or None,
        )

        final_result = l1_result

        # 6. Route to Level 2 (AI) if needed
        # Criteria: OTHER_ACCESS or AMBIGUOUS or explicitly WAITING_AI
        needs_level_2 = (
            l1_result.classification in (ClassificationCode.OTHER_ACCESS, ClassificationCode.AMBIGUOUS)
            or l1_result.status == AnalysisStatus.WAITING_AI
        )

        if needs_level_2:
            l2_result = self.ai_provider.analyze_ticket(
                ticket_key=clean_key,
                summary=summary,
                description=description,
                category=process_category,
                level_1_result=l1_result,
            )
            final_result = l2_result

        # 7. Analyze comments & status for operational dependencies
        op_status, dependency, dependency_action = self._detect_operational_status(
            ticket_status=status,
            comments=comments,
            missing_info=final_result.missing_info,
            extracted_data=final_result.extracted_data,
        )
        final_result.operational_status = op_status
        if dependency:
            final_result.identified_dependency = dependency
        if dependency_action:
            final_result.suggested_next_action = dependency_action

        # 8. Persist result in DB
        save_salsa_ticket_analysis(final_result.to_dict())

        return final_result

    def _detect_operational_status(
        self,
        ticket_status: str,
        comments: list[Any],
        missing_info: list[str],
        extracted_data: Any,
    ) -> Tuple[OperationalStatus, Optional[str], Optional[str]]:
        """
        Determines operational dependency from Jira status and comment history.
        Detects pending feedback, missing information or approvals.
        """
        import re

        status_lower = (ticket_status or "").lower()
        all_comment_texts = []
        for c in comments:
            if isinstance(c, dict):
                body = c.get("body", "")
                author = c.get("author", "")
                all_comment_texts.append(f"{author}: {body}")
            elif isinstance(c, str):
                all_comment_texts.append(c)

        combined_comments = "\n".join(all_comment_texts)

        # 1. Check for waiting for feedback in comments
        # Pattern: aguarda feedback de/da/do [Nome], aguardando resposta de [Nome], à espera de [Nome], aguardar feedback de [Nome]
        feedback_patterns = [
            re.compile(r"(?:aguard(?:a|ar|ando)?|esper(?:a|ar|ando)?|à espera|a aguardar)\s+(?:(?:de\s+)?(?:feedback|resposta|validação|informação|confirmação)\s+)?(?:d[eoa]|da|do|de|por)\s+([A-ZÀ-Ú][a-zà-ú]+(?:\s+[A-ZÀ-Ú][a-zà-ú]+)?)", re.IGNORECASE),
        ]

        for pat in feedback_patterns:
            m = pat.search(combined_comments)
            if m:
                person = m.group(1).strip()
                dep_desc = f"resposta de {person}"
                action = f"Aguardar a confirmação de {person} e validar as autorizações solicitadas na PFCG."
                return OperationalStatus.WAITING_FOR_FEEDBACK, dep_desc, action

        # 2. Check ticket status itself
        if "waiting for customer" in status_lower or "aguarda cliente" in status_lower or "aguarda solicitante" in status_lower:
            return (
                OperationalStatus.WAITING_FOR_FEEDBACK,
                "resposta do solicitante",
                "Aguardar resposta do solicitante aos esclarecimentos solicitados.",
            )
        if "approval" in status_lower or "aprovação" in status_lower:
            return (
                OperationalStatus.REQUIRES_APPROVAL,
                "aprovação formal da chefia",
                "Aguardar aprovação hierárquica no fluxo do Jira.",
            )

        # 3. Check for missing information
        if missing_info and any("utilizador" in m.lower() for m in missing_info):
            return (
                OperationalStatus.MISSING_INFORMATION,
                "dados do utilizador ausentes",
                "Solicitar os dados completos do utilizador ao requerente.",
            )

        return OperationalStatus.READY_FOR_REVIEW, None, None

    def get_analysis_for_ticket(self, ticket_key: str) -> Optional[dict[str, Any]]:
        """Returns existing stored analysis."""
        return get_salsa_ticket_analysis(ticket_key)

    def get_history_for_ticket(self, ticket_key: str) -> list[dict[str, Any]]:
        """Returns history of past analyses."""
        return get_salsa_ticket_analysis_history(ticket_key)

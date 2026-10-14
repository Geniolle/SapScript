"""
Salsa IT Agent - Level 2 AI Provider Abstraction.
Vendor-agnostic AI interface (compatible with OpenAI, Anthropic, Gemini, Ollama, etc.).

Safety rules:
- External provider disabled by default.
- Returns AI_NOT_CONFIGURED when no authorized provider is enabled.
- Never sends ticket data outside without explicit authorization.
- Provides MockAIProvider for test validation.
- Output validated by schema.
"""
from __future__ import annotations

import json
import os
from abc import ABC, abstractmethod
from typing import Any, Dict, Optional

from .models import (
    AnalysisLevel,
    AnalysisResult,
    AnalysisStatus,
    ClassificationCode,
    ExtractedEntities,
)


class BaseAIProvider(ABC):
    """Abstract interface for AI classification providers."""

    @abstractmethod
    def is_configured(self) -> bool:
        """Checks if the provider is fully configured and authorized."""
        pass

    @abstractmethod
    def analyze_ticket(
        self,
        ticket_key: str,
        summary: str,
        description: str,
        category: str,
        level_1_result: AnalysisResult,
    ) -> AnalysisResult:
        """Executes Level 2 AI assisted classification."""
        pass


class DisabledAIProvider(BaseAIProvider):
    """
    Default safe provider when AI is not configured or not authorized.
    Returns AI_NOT_CONFIGURED without sending any network request.
    """

    def is_configured(self) -> bool:
        return False

    def analyze_ticket(
        self,
        ticket_key: str,
        summary: str,
        description: str,
        category: str,
        level_1_result: AnalysisResult,
    ) -> AnalysisResult:
        return AnalysisResult(
            ticket_key=ticket_key,
            classification=level_1_result.classification,
            level=AnalysisLevel.LEVEL_2,
            status=AnalysisStatus.AI_NOT_CONFIGURED,
            intent=level_1_result.intent or "IA não configurada",
            justification=(
                "Fornecedor de IA não está ativo ou configurado no ambiente. "
                f"Resultado mantido do Nível 1: {level_1_result.justification}"
            ),
            evidences=level_1_result.evidences,
            extracted_data=level_1_result.extracted_data,
            missing_info=level_1_result.missing_info,
            suggested_next_action=(
                "Configurar fornecedor de IA ou proceder à validação técnica manual."
            ),
            requires_human_review=True,
            confidence=level_1_result.confidence,
            rules_version=level_1_result.rules_version,
            ai_model_version=None,
            content_hash=level_1_result.content_hash,
        )


class MockAIProvider(BaseAIProvider):
    """
    Mock AI provider for automated testing and deterministic validation.
    Allows simulating valid outputs, invalid outputs, schema errors, timeouts and contradictory outputs.
    """

    def __init__(
        self,
        mock_response: Optional[Dict[str, Any]] = None,
        simulate_error: Optional[str] = None,
        simulate_invalid_schema: bool = False,
    ):
        self._mock_response = mock_response or {}
        self._simulate_error = simulate_error
        self._simulate_invalid_schema = simulate_invalid_schema

    def is_configured(self) -> bool:
        return True

    def analyze_ticket(
        self,
        ticket_key: str,
        summary: str,
        description: str,
        category: str,
        level_1_result: AnalysisResult,
    ) -> AnalysisResult:
        if self._simulate_error:
            return AnalysisResult(
                ticket_key=ticket_key,
                classification=level_1_result.classification,
                level=AnalysisLevel.LEVEL_2,
                status=AnalysisStatus.ERROR,
                intent="Falha no fornecedor de IA",
                justification=f"Erro simulado do fornecedor: {self._simulate_error}",
                error_message=self._simulate_error,
                requires_human_review=True,
                content_hash=level_1_result.content_hash,
            )

        if self._simulate_invalid_schema:
            return AnalysisResult(
                ticket_key=ticket_key,
                classification=ClassificationCode.AMBIGUOUS,
                level=AnalysisLevel.LEVEL_2,
                status=AnalysisStatus.ERROR,
                intent="Resposta de IA fora do formato esperado",
                justification="O fornecedor devolveu uma estrutura inválida que foi recusada pelo validador.",
                error_message="Schema validation error",
                requires_human_review=True,
                content_hash=level_1_result.content_hash,
            )

        # Build structured response from mock dict
        classification_raw = self._mock_response.get(
            "classification", level_1_result.classification.value
        )
        try:
            classification = ClassificationCode(classification_raw)
        except ValueError:
            classification = ClassificationCode.AMBIGUOUS

        extracted_raw = self._mock_response.get("extracted_data") or {}
        extracted = ExtractedEntities.from_dict(extracted_raw)

        return AnalysisResult(
            ticket_key=ticket_key,
            classification=classification,
            level=AnalysisLevel.LEVEL_2,
            status=AnalysisStatus.LEVEL_2_COMPLETED,
            intent=self._mock_response.get("intent", "Análise de IA concluída"),
            justification=self._mock_response.get(
                "justification", "Classificação refinada via modelo generativo."
            ),
            evidences=list(self._mock_response.get("evidences", ["Evidência contextual de IA"])),
            extracted_data=extracted,
            missing_info=list(self._mock_response.get("missing_info", [])),
            suggested_next_action=self._mock_response.get(
                "suggested_next_action", "Revisar proposta e autorizar execução."
            ),
            requires_human_review=bool(self._mock_response.get("requires_human_review", False)),
            confidence=float(self._mock_response.get("confidence", 0.9)),
            ai_model_version=self._mock_response.get("ai_model_version", "mock-model-v1"),
            content_hash=level_1_result.content_hash,
        )


def get_ai_provider() -> BaseAIProvider:
    """
    Factory function for the active AI Provider.
    In Phase 1, defaults to DisabledAIProvider unless explicitly enabled in environment.
    """
    ai_enabled = os.getenv("SALSA_AI_ENABLED", "false").lower() in ("1", "true", "yes")
    provider_type = os.getenv("SALSA_AI_PROVIDER", "disabled").lower()

    if not ai_enabled or provider_type == "disabled":
        return DisabledAIProvider()

    # If mock provider is requested
    if provider_type == "mock":
        return MockAIProvider()

    # For safety in Phase 1, any other unconfigured provider falls back to DisabledAIProvider
    return DisabledAIProvider()

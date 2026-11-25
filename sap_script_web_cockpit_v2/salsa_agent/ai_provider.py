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
import urllib.error
import urllib.request
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


class GeminiAIProvider(BaseAIProvider):
    """
    Google Gemini via REST generateContent (sem dependências novas).
    Dados enviados: categoria, resumo e descrição do ticket. Nunca comentários nem credenciais.
    Resultado sempre com requires_human_review=True.
    """

    ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"

    def __init__(self, api_key: str, model: str = "gemini-2.5-flash", timeout: int = 60):
        self._api_key = api_key
        self._model = model
        self._timeout = timeout

    def is_configured(self) -> bool:
        return bool(self._api_key)

    def _build_prompt(
        self, summary: str, description: str, category: str, level_1_result: AnalysisResult
    ) -> str:
        allowed = ", ".join(c.value for c in ClassificationCode)
        return (
            "Analisa este pedido de IT SAP e responde apenas com JSON com as chaves: "
            f"classification (um de: {allowed}), intent, justification, evidences (lista), "
            "missing_info (lista), suggested_next_action, confidence (número de 0 a 1). "
            "Não inventes dados. Se faltar informação, usa INSUFFICIENT_INFORMATION.\n\n"
            f"Categoria: {category}\n"
            f"Resumo: {summary}\n"
            f"Descrição: {description}\n"
            f"Classificação do Nível 1: {level_1_result.classification.value}"
        )

    def analyze_ticket(
        self,
        ticket_key: str,
        summary: str,
        description: str,
        category: str,
        level_1_result: AnalysisResult,
    ) -> AnalysisResult:
        def _fail(message: str) -> AnalysisResult:
            return AnalysisResult(
                ticket_key=ticket_key,
                classification=level_1_result.classification,
                level=AnalysisLevel.LEVEL_2,
                status=AnalysisStatus.ERROR,
                intent="Falha no fornecedor de IA",
                justification=message,
                error_message=message,
                requires_human_review=True,
                content_hash=level_1_result.content_hash,
            )

        body = {
            "contents": [
                {"parts": [{"text": self._build_prompt(summary, description, category, level_1_result)}]}
            ],
            "generationConfig": {"responseMimeType": "application/json", "temperature": 0},
        }
        request = urllib.request.Request(
            self.ENDPOINT.format(model=self._model),
            data=json.dumps(body).encode("utf-8"),
            headers={"Content-Type": "application/json", "x-goog-api-key": self._api_key},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self._timeout) as response:
                payload = json.loads(response.read().decode("utf-8"))
            text = payload["candidates"][0]["content"]["parts"][0]["text"]
            out = json.loads(text)
        except urllib.error.HTTPError as exc:
            return _fail(f"Gemini devolveu HTTP {exc.code}")
        except OSError:
            return _fail("Gemini indisponível: falha de rede ou timeout")
        except (KeyError, IndexError, TypeError, ValueError):
            return _fail("Resposta do Gemini fora do formato esperado")

        if not isinstance(out, dict):
            return _fail("Resposta do Gemini fora do formato esperado")

        try:
            classification = ClassificationCode(out.get("classification"))
        except ValueError:
            classification = ClassificationCode.AMBIGUOUS

        try:
            confidence = min(max(float(out.get("confidence", 0.5)), 0.0), 1.0)
        except (TypeError, ValueError):
            confidence = 0.5

        def _str_list(value: Any) -> list[str]:
            return [str(item) for item in value] if isinstance(value, list) else []

        return AnalysisResult(
            ticket_key=ticket_key,
            classification=classification,
            level=AnalysisLevel.LEVEL_2,
            status=AnalysisStatus.LEVEL_2_COMPLETED,
            intent=str(out.get("intent") or "Análise de IA concluída"),
            justification=str(out.get("justification") or ""),
            evidences=_str_list(out.get("evidences")),
            extracted_data=level_1_result.extracted_data,
            missing_info=_str_list(out.get("missing_info")),
            suggested_next_action=str(out.get("suggested_next_action") or "Revisar proposta."),
            requires_human_review=True,
            confidence=confidence,
            ai_model_version=self._model,
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

    if provider_type == "gemini":
        api_key = os.getenv("GEMINI_API_KEY", "").strip()
        if api_key:
            return GeminiAIProvider(api_key, model=os.getenv("SALSA_GEMINI_MODEL", "gemini-2.5-flash"))
        return DisabledAIProvider()

    # For safety in Phase 1, any other unconfigured provider falls back to DisabledAIProvider
    return DisabledAIProvider()

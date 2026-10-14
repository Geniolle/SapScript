"""
Salsa IT Agent - Extractor of Entities for Phase 1.
Extracts SAP user IDs, names, reference users, SAP environments, company codes, etc.
Deterministic and strictly evidence-based: NO hallucinated or guessed values.
"""
from __future__ import annotations

import re
import unicodedata
from typing import List, Optional, Tuple
from .models import ExtractedEntities


def normalize_text_for_search(text: str) -> str:
    """Normalizes text by removing diacritics and converting to lowercase."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


class SalsaEntityExtractor:
    """Deterministic extractor for ticket information."""

    # Patterns for SAP User IDs (e.g., S80001977, USR12345, C12345, etc.)
    SAP_USER_PATTERNS = [
        re.compile(r"\b([S|C]\d{5,9})\b", re.IGNORECASE),
        re.compile(r"\butilizador\s+(?:SAP\s+)?([A-Z0-9_]{3,12})\b", re.IGNORECASE),
        re.compile(r"\buser\s+(?:SAP\s+)?([A-Z0-9_]{3,12})\b", re.IGNORECASE),
        re.compile(r"\bconta\s+(?:SAP\s+)?([A-Z0-9_]{3,12})\b", re.IGNORECASE),
        re.compile(r"\blogin\s+(?:SAP\s+)?([A-Z0-9_]{3,12})\b", re.IGNORECASE),
        re.compile(r"\busername\s+(?:SAP\s+)?([A-Z0-9_]{3,12})\b", re.IGNORECASE),
    ]

    # Reference user patterns (e.g., "mesmas permissões que Marta Santiago", "que a Marta Santiago", "como o user X", "cópia de Y")
    REFERENCE_USER_PATTERNS = [
        re.compile(r"(?:mesmas\s+permiss(?:õ|o)es|mesmo\s+acesso|perfil\s+igual|igual|c(?:ó|o)pia)\s+(?:que|de|ao|a)?\s+(?:o|a|ao|à)?\s+([A-ZÀ-Ú][a-zà-ú]+(?:\s+[A-ZÀ-Ú][a-zà-ú]+)+)", re.IGNORECASE),
        re.compile(r"(?:mesmas\s+permiss(?:õ|o)es|mesmo\s+acesso|igual|c(?:ó|o)pia)\s+(?:que|de|ao|a)?\s+(?:o\s+utilizador|o\s+user|o\s+colaborador|a\s+utilizadora|a\s+colaboradora)\s+([A-ZÀ-Úa-zà-ú0-9_ ]+?)(?:[\.,\n;]|$)", re.IGNORECASE),
        re.compile(r"(?:como\s+o\s+user|como\s+o\s+utilizador|como\s+a\s+utilizadora|baseado\s+em|cópia\s+de)\s+([A-ZÀ-Úa-zà-ú0-9_]+)", re.IGNORECASE),
    ]

    # Environment detection
    ENV_PATTERNS = [
        (r"\b(?:prd|prod|produ[çc][aã]o|produtivo)\b", "PRD"),
        (r"\b(?:qad|qua|qualidade|testes?|staging)\b", "QAD"),
        (r"\b(?:dev|desenvolvimento)\b", "DEV"),
    ]

    # Company code patterns (e.g., "empresa 2080", "sociedade 1000", "empresa SLS")
    COMPANY_PATTERNS = [
        re.compile(r"\b(?:empresa|sociedade|company\s*code)\s*(?:nº|n\.º|#)?\s*([0-9]{4}|[A-Z0-9]{2,4})\b", re.IGNORECASE),
    ]

    def extract(self, text: str, summary: str = "", requester: Optional[str] = None) -> ExtractedEntities:
        combined = f"{summary}\n{text}".strip()
        if not combined:
            return ExtractedEntities(
                requester=requester,
                missing_information=["Descrição ou detalhes do pedido"]
            )

        sap_user_id, all_users = self._extract_sap_user(combined)
        user_name = self._extract_user_name(combined, sap_user_id)
        reference_user = self._extract_reference_user(combined)
        environments = self._extract_environments(combined)
        company_code = self._extract_company_code(combined)

        # If a reference user was matched but it is the same as the target user, clear reference
        if reference_user and sap_user_id and reference_user.upper() == sap_user_id.upper():
            reference_user = None

        missing = []
        if not sap_user_id and not reference_user and not user_name:
            missing.append("Identificação do utilizador (ID SAP ou nome)")
        if not environments:
            missing.append("Ambiente SAP de destino (PRD / QAD / DEV)")

        return ExtractedEntities(
            sap_user_id=sap_user_id,
            user_name=user_name,
            reference_user=reference_user,
            requester=requester,
            sap_environment=environments,
            company_code=company_code,
            missing_information=missing,
        )


    def _extract_sap_user(self, text: str) -> Tuple[Optional[str], List[str]]:
        found_users: List[str] = []
        
        # Stop words that look like words following 'user' or 'utilizador'
        stopwords = {
            "sap", "novo", "nova", "qualidade", "produtivo", "producao", "dev", "qad", "prd",
            "com", "sem", "para", "bloqueado", "bloqueada", "mesmo", "mesma", "teste",
            "acesso", "senha", "password", "pedimos", "preciso", "solicito", "criar", "reset",
            "desbloquear", "desbloqueio", "autorizacao", "perfil", "transacao", "transacoes"
        }

        # Check explicit S-number or C-number first
        s_matches = re.findall(r"\b([S|C]\d{5,9})\b", text, re.IGNORECASE)
        for m in s_matches:
            upper_m = m.upper()
            if upper_m not in found_users:
                found_users.append(upper_m)

        for pat in self.SAP_USER_PATTERNS[1:]:
            for match in pat.finditer(text):
                val = match.group(1).strip()
                if val.lower() not in stopwords and len(val) >= 3:
                    if val.upper() not in found_users:
                        found_users.append(val.upper())

        primary_user = found_users[0] if found_users else None
        return primary_user, found_users

    def _extract_reference_user(self, text: str) -> Optional[str]:
        for pat in self.REFERENCE_USER_PATTERNS:
            match = pat.search(text)
            if match:
                ref = match.group(1).strip()
                # Clean punctuation at end
                ref = re.sub(r"[\.,;:!\?]+$", "", ref).strip()
                if len(ref) >= 3 and not ref.lower().startswith("sap"):
                    return ref
        return None

    def _extract_user_name(self, text: str, found_sap_id: Optional[str]) -> Optional[str]:
        # Check patterns like "colaborador Marta Santiago", "utilizador João Silva"
        pat = re.compile(r"\b(?:colaborador(?:a)?|funcion[áa]rio(?:a)?|utilizador(?:a)?)\s+([A-ZÀ-Ú][a-zà-ú]+(?:\s+[A-ZÀ-Ú][a-zà-ú]+)+)", re.UNICODE)
        m = pat.search(text)
        if m:
            name = m.group(1).strip()
            if found_sap_id and name.upper() == found_sap_id.upper():
                return None
            return name
        return None

    def _extract_environments(self, text: str) -> List[str]:
        norm = normalize_text_for_search(text)
        detected = []
        for pat, env_code in self.ENV_PATTERNS:
            if re.search(pat, norm):
                if env_code not in detected:
                    detected.append(env_code)
        return detected

    def _extract_company_code(self, text: str) -> Optional[str]:
        for pat in self.COMPANY_PATTERNS:
            m = pat.search(text)
            if m:
                return m.group(1).strip().upper()
        return None

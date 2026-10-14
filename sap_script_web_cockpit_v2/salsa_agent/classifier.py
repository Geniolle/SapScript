"""
Salsa IT Agent - Level 1 Deterministic Classifier.
Implements rule-based, linguistically robust classification for SAP CA Acessos tickets:
- CREATE_USER
- UNLOCK_USER
- RESET_PASSWORD
- OTHER_ACCESS
- POSSIBLE_MISCATEGORIZATION
- AMBIGUOUS
- INSUFFICIENT_INFORMATION

Safety & semantic guidelines:
- Normalized linguistic search supporting PT and EN variations.
- Negation detection ("Não é necessário reset, apenas desbloquear").
- Multi-operation detection -> AMBIGUOUS.
- Avoids misinterpreting authorisations/roles alterations as CREATE_USER.
- Defends against prompt injection (instruction-like phrases).
- Deterministic and auditable with clear reasons and extracted evidences.
"""
from __future__ import annotations

import hashlib
import re
import unicodedata
from typing import List, Optional, Tuple

from .extractor import SalsaEntityExtractor, normalize_text_for_search
from .models import (
    AnalysisLevel,
    AnalysisResult,
    AnalysisStatus,
    ClassificationCode,
    ExtractedEntities,
)

RULES_VERSION = "1.1.0"


class SalsaLevel1Classifier:
    """Deterministic Level 1 classifier."""

    def __init__(self):
        self.extractor = SalsaEntityExtractor()

    # Prompt injection check
    INJECTION_PATTERNS = [
        re.compile(r"ignore\s+(?:all\s+)?previous\s+instructions", re.IGNORECASE),
        re.compile(r"you\s+are\s+now\s+(?:in|a)\s+", re.IGNORECASE),
        re.compile(r"system\s+prompt", re.IGNORECASE),
        re.compile(r"forget\s+(?:all\s+)?prior\s+rules", re.IGNORECASE),
        re.compile(r"execute\s+(?:destructive|arbitrary|sql|sap|sh|bash)", re.IGNORECASE),
        re.compile(r"bypass\s+security", re.IGNORECASE),
        re.compile(r"grant\s+(?:all|full|root|sap_all)\s+access", re.IGNORECASE),
    ]

    # Miscategorization indicators (financial/business logic problems, budget, PR/PO, CJ30, MM, etc.)
    MISCATEGORIZATION_PATTERNS = [
        (r"\b(?:cj30|cj20n|me21n|me22n|me23n|pr|po|orcamento|orcamental|budget|franchisad[ao]s?)\b", "Problema orçamental ou funcional de projetos/compras"),
        (r"\b(?:fator\s*conversao|preco|tabela\s*de\s*precos?|taxa\s*iva)\b", "Parametrização de preços/taxas"),
        (r"\b(?:f110|pagamento\s*automatico|dmee|remessa)\b", "Processo financeiro F110/DMEE"),
        (r"\b(?:transito\s*de\s*stocks?|guia\s*de\s*transporte|migo|mb51)\b", "Movimentação logística de stocks"),
    ]

    # Complex / other access indicators (SU53, PA20, restrições finas, notas de crédito, autorizações de campos)
    OTHER_ACCESS_PATTERNS = [
        (r"\b(?:nota[s]?\s+de\s+cr[eé]dito|emiss[aã]o\s+de\s+notas|equipa\s+cci|restringir\s+a\s+emiss[aã]o)\b", "Restrição/autorização para emissão de documentos/notas de crédito"),
        (r"\b(?:pa20|c&b|dados\s+de\s+colaboradores|acessos?\s+indevidos?|rever\s+acesso\s+indevido)\b", "Revisão de autorizações/transações sensíveis de RH/PA20"),
        (r"\b(?:su53|falta\s+de\s+autoriza[çc][aã]o|perfil\s+de\s+autoriza[çc][aã]o|role|fun[çc][aã]o\s+pfcg|transa[çc][aã]o\s+[a-z0-9_]+)\b", "Pedido de perfil/role/autorização específica ou diagnóstico SU53"),
        (r"\b(?:limita[çc][aã]o\s+acesso|restringir|tirar\s+acesso|remover\s+permiss[aã]o|alterar\s+perfil)\b", "Limitação, ajuste fino ou alteração de permissões existentes"),
    ]

    # Core operation intent verbs & terms
    CREATE_USER_SIGNALS = [
        r"\bcriar(?:[- ](?:me|nos))?\s+(?:o\s+|a\s+|um\s+|uma\s+)?(?:utilizador|user|conta|login|novo\s+colaborador)\b",
        r"\bpod(?:em|e)\s+(?:pf\s+|por\s+favor\s+)?criar(?:[- ](?:me|nos))?\s+(?:um\s+|uma\s+)?(?:utilizador|user|conta|login)\b",
        r"\bcria[çc][aã]o\s+de\s+(?:utilizador|user|conta|login)\b",
        r"\bsolicito\s+(?:a\s+)?cria[çc][aã]o\s+de\s+(?:uma\s+)?(?:conta|utilizador|user)\s+sap\b",
        r"\bnovo\s+colaborador\s+necessita\s+de\s+acesso\b",
        r"\bprecis(?:o|ava)\s+(?:de\s+)?(?:uma\s+conta\s+sap|acesso\s+sap\s+pela\s+primeira\s+vez)\b",
        r"\bsolicito\s+acesso\s+sap\s+para\s+novo\b",
        r"\bcreate\s+(?:new\s+)?user\b",
        r"\bcreate\s+sap\s+account\b",
    ]

    UNLOCK_USER_SIGNALS = [
        r"\bdesbloquear\s+(?:o\s+|a\s+|o\s+utilizador|a\s+conta|o\s+user|o\s+login)?\b",
        r"\bdesbloque(?:ar[- ](?:lo|la|los|las)|ie)\b",
        r"\bdesbloqueio\s+(?:d[eo]\s+)?(?:utilizador|user|conta|login)\b",
        r"\bconta\s+bloqueada\b",
        r"\butilizador\s+bloqueado\b",
        r"\buser\s+bloqueado\b",
        r"\bbloqueado\s+ap[oó]s\s+tentativas\b",
        r"\bunlock\s+user\b",
        r"\bunlock\s+account\b",
    ]


    RESET_PASSWORD_SIGNALS = [
        r"\breset\s+(?:d[ae]\s+)?(?:password|senha|palavra[- ]passe)\b",
        r"\breposi[çc][aã]o\s+d[ae]\s+(?:password|senha|palavra[- ]passe)\b",
        r"\bredefinir\s+(?:a\s+)?(?:password|senha|palavra[- ]passe)\b",
        r"\besqueci\s+(?:a\s+)?(?:password|senha|palavra[- ]passe)\b",
        r"\baltera[çc][aã]o\s+de\s+password\b",
        r"\bnova\s+password\b",
        r"\bpassword\s+reset\b",
    ]

    def analyze(
        self,
        ticket_key: str,
        summary: str,
        description: str,
        process_category: str = "",
        requester: Optional[str] = None,
    ) -> AnalysisResult:
        """Executes Level 1 deterministic analysis."""
        combined_text = f"{summary or ''}\n{description or ''}".strip()
        content_hash = self.compute_content_hash(summary or "", description or "", process_category or "")

        # 1. Check for empty or insufficient description
        clean_desc = (description or "").strip()
        clean_sum = (summary or "").strip()
        if not clean_desc and not clean_sum:
            return AnalysisResult(
                ticket_key=ticket_key,
                classification=ClassificationCode.INSUFFICIENT_INFORMATION,
                level=AnalysisLevel.LEVEL_1,
                status=AnalysisStatus.REVISION_REQUIRED,
                intent="Conteúdo vazio ou sem dados",
                justification="O ticket não possui título nem descrição preenchidos.",
                missing_info=["Resumo e descrição detalhada do pedido"],
                suggested_next_action="Solicitar detalhe do pedido ao requerente.",
                requires_human_review=True,
                confidence=1.0,
                content_hash=content_hash,
            )

        # 2. Check for security/prompt injection in text
        for pat in self.INJECTION_PATTERNS:
            if pat.search(combined_text):
                return AnalysisResult(
                    ticket_key=ticket_key,
                    classification=ClassificationCode.AMBIGUOUS,
                    level=AnalysisLevel.LEVEL_1,
                    status=AnalysisStatus.REVISION_REQUIRED,
                    intent="Instruções anómalas detetadas no texto",
                    justification="O texto contém padrões suspeitos semelhantes a injeção de instruções ou comandos restritos.",
                    evidences=["Texto com instrução anómala identificada"],
                    missing_info=["Validação de segurança"],
                    suggested_next_action="Revisão manual prioritária por especialista de segurança.",
                    requires_human_review=True,
                    confidence=0.5,
                    content_hash=content_hash,
                )

        norm_summary = normalize_text_for_search(clean_sum)
        norm_desc = normalize_text_for_search(clean_desc)
        norm_combined = f"{norm_summary}\n{norm_desc}"

        # 3. Check for obvious miscategorization (e.g. IZ-55652: budget, CJ30, franchisadas)
        for pat, reason in self.MISCATEGORIZATION_PATTERNS:
            if re.search(pat, norm_combined):
                evidences = [f"Encontrado padrão: {reason}"]
                return AnalysisResult(
                    ticket_key=ticket_key,
                    classification=ClassificationCode.POSSIBLE_MISCATEGORIZATION,
                    level=AnalysisLevel.LEVEL_1,
                    status=AnalysisStatus.REVISION_REQUIRED,
                    intent="Possível erro de categorização do ticket",
                    justification=f"O ticket refere conceitos funcionais/orçamentais ({reason}) que não aparentam ser gestão de acessos.",
                    evidences=evidences,
                    missing_info=["Confirmação da categoria e equipa responsável"],
                    suggested_next_action="Encaminhar para a equipa funcional correspondente ou Nível 2 para revisão.",
                    requires_human_review=True,
                    confidence=0.85,
                    content_hash=content_hash,
                )

        # 4. Check for OTHER_ACCESS (e.g. IZ-59453 notas crédito / IZ-52936 PA20 C&B / SU53 / roles)
        for pat, reason in self.OTHER_ACCESS_PATTERNS:
            if re.search(pat, norm_combined):
                evidences = [f"Encontrado padrão de autorizações: {reason}"]
                extracted = self.extractor.extract(clean_desc, clean_sum, requester=requester)
                return AnalysisResult(
                    ticket_key=ticket_key,
                    classification=ClassificationCode.OTHER_ACCESS,
                    level=AnalysisLevel.LEVEL_1,
                    status=AnalysisStatus.WAITING_AI,
                    intent=reason,
                    justification=f"O pedido solicita alteração ou limitação específica de autorizações ({reason}) e requer análise detalhada.",
                    evidences=evidences,
                    extracted_data=extracted,
                    missing_info=extracted.missing_information,
                    suggested_next_action="Encaminhar para Nível 2 (IA) para diagnóstico semântico de autorizações PFCG.",
                    requires_human_review=True,
                    confidence=0.9,
                    content_hash=content_hash,
                )

        # 5. Extract core operations while respecting negation & multi-operation
        has_create, create_evidences = self._detect_intent(
            norm_desc, norm_summary, self.CREATE_USER_SIGNALS, "CREATE_USER"
        )
        has_unlock, unlock_evidences = self._detect_intent(
            norm_desc, norm_summary, self.UNLOCK_USER_SIGNALS, "UNLOCK_USER"
        )
        has_reset, reset_evidences = self._detect_intent(
            norm_desc, norm_summary, self.RESET_PASSWORD_SIGNALS, "RESET_PASSWORD"
        )

        detected_operations = []
        if has_create:
            detected_operations.append((ClassificationCode.CREATE_USER, create_evidences))
        if has_unlock:
            detected_operations.append((ClassificationCode.UNLOCK_USER, unlock_evidences))
        if has_reset:
            detected_operations.append((ClassificationCode.RESET_PASSWORD, reset_evidences))

        extracted_entities = self.extractor.extract(clean_desc, clean_sum, requester=requester)

        # Multi-operation detection -> AMBIGUOUS
        if len(detected_operations) > 1:
            op_names = [op[0].value for op in detected_operations]
            all_evidences = [ev for op in detected_operations for ev in op[1]]
            return AnalysisResult(
                ticket_key=ticket_key,
                classification=ClassificationCode.AMBIGUOUS,
                level=AnalysisLevel.LEVEL_1,
                status=AnalysisStatus.REVISION_REQUIRED,
                intent=f"Múltiplas operações solicitadas ({', '.join(op_names)})",
                justification=f"O pedido menciona mais de uma operação de acessos no mesmo ticket: {', '.join(op_names)}.",
                evidences=all_evidences,
                extracted_data=extracted_entities,
                missing_info=extracted_entities.missing_information,
                suggested_next_action="Revisão manual ou decomposição do pedido em tarefas individuais.",
                requires_human_review=True,
                confidence=0.7,
                content_hash=content_hash,
            )

        # Exactly 1 operation detected
        if len(detected_operations) == 1:
            op_code, op_evidences = detected_operations[0]

            if op_code == ClassificationCode.RESET_PASSWORD:
                intent = "Redefinição de Password de Utilizador SAP"
                justification = "Identificada intenção inequívoca de reset/redefinição de palavra-passe."
                suggested_action = "Pré-validação de logon e ambiente antes de qualquer intervenção."
            elif op_code == ClassificationCode.UNLOCK_USER:
                intent = "Desbloqueio de Utilizador SAP"
                justification = "Identificada intenção inequívoca de desbloqueio de conta SAP."
                suggested_action = "Verificar estado do utilizador (USR02) no ambiente de destino."
            else:  # CREATE_USER
                intent = "Criação de Utilizador SAP"
                justification = "Identificada intenção inequívoca de criação de novo utilizador SAP."
                suggested_action = "Validar dados cadastrais e função de referência na PFCG."

            # Update requested operation in extracted data
            extracted_entities.requested_operation = op_code.value

            return AnalysisResult(
                ticket_key=ticket_key,
                classification=op_code,
                level=AnalysisLevel.LEVEL_1,
                status=AnalysisStatus.LEVEL_1_COMPLETED,
                intent=intent,
                justification=justification,
                evidences=op_evidences,
                extracted_data=extracted_entities,
                missing_info=extracted_entities.missing_information,
                suggested_next_action=suggested_action,
                requires_human_review=len(extracted_entities.missing_information) > 0,
                confidence=0.95 if not extracted_entities.missing_information else 0.85,
                content_hash=content_hash,
            )

        # 6. Fallback: No clear single operation found
        # Check if description mentions vague access
        if "acesso" in norm_combined or "permiss" in norm_combined or "login" in norm_combined:
            return AnalysisResult(
                ticket_key=ticket_key,
                classification=ClassificationCode.OTHER_ACCESS,
                level=AnalysisLevel.LEVEL_1,
                status=AnalysisStatus.WAITING_AI,
                intent="Pedido genérico de acessos",
                justification="O ticket menciona termos de acesso mas sem operação de utilizador identificável.",
                evidences=["Termos genéricos de acesso identificados"],
                extracted_data=extracted_entities,
                missing_info=extracted_entities.missing_information,
                suggested_next_action="Encaminhar para Nível 2 (IA) para diagnóstico contextual.",
                requires_human_review=True,
                confidence=0.6,
                content_hash=content_hash,
            )

        return AnalysisResult(
            ticket_key=ticket_key,
            classification=ClassificationCode.INSUFFICIENT_INFORMATION,
            level=AnalysisLevel.LEVEL_1,
            status=AnalysisStatus.REVISION_REQUIRED,
            intent="Informação insuficiente para determinar a operação",
            justification="A descrição não apresenta detalhes suficientes para classificar uma ação de acessos SAP.",
            missing_info=["Detalhes da transação, erro ou utilizador pretendido"],
            suggested_next_action="Solicitar clarificação do pedido ao criador do ticket.",
            requires_human_review=True,
            confidence=0.5,
            content_hash=content_hash,
        )

    def _detect_intent(
        self, norm_desc: str, norm_summary: str, signals: List[str], op_type: str
    ) -> Tuple[bool, List[str]]:
        """
        Detects an operation intent while taking negations into account.
        Checks description first (higher priority), then summary.
        """
        evidences: List[str] = []
        matched = False

        # Check in sentences to inspect negation scope
        sentences = re.split(r"[\.\n;!\?]+", f"{norm_desc}\n{norm_summary}")
        for s in sentences:
            s_clean = s.strip()
            if not s_clean:
                continue

            for pat in signals:
                if re.search(pat, s_clean):
                    # Check negation in the same sentence or clause
                    if self._is_negated(s_clean, pat):
                        # Negated, do not count this match!
                        continue
                    matched = True
                    evidences.append(f"Padrão '{op_type}' encontrado: '{s_clean[:80]}'")
                    break

        return matched, evidences

    def _is_negated(self, sentence: str, pattern: str) -> bool:
        """
        Checks if the pattern in the sentence is preceded by a negation
        e.g., "Não é necessário reset de password", "sem necessidade de criar utilizador".
        """
        negation_prefixes = [
            r"\bn[aã]o\s+(?:[eé]|ser[aá]|foi|preciso|necess[aá]rio|h[aá])\b",
            r"\bn[aã]o\s+fazer\b",
            r"\bsem\s+(?:necessidade|fazer|reset|criar)\b",
            r"\bdispensa\b",
            r"\bnunca\b",
        ]
        for neg in negation_prefixes:
            m_neg = re.search(neg, sentence)
            m_pat = re.search(pattern, sentence)
            if m_neg and m_pat and m_neg.start() < m_pat.start():
                # Make sure there is not an adversative conjunction ("mas", "apenas") between them
                between = sentence[m_neg.end():m_pat.start()]
                if not re.search(r"\b(?:mas|por[eé]m|contudo|apenas|s[oó])\b", between):
                    return True
        return False

    @staticmethod
    def compute_content_hash(summary: str, description: str, process_category: str) -> str:
        """Computes a SHA-256 hash of the relevant ticket content."""
        payload = f"{summary.strip()}|{description.strip()}|{process_category.strip()}".encode("utf-8")
        return hashlib.sha256(payload).hexdigest()

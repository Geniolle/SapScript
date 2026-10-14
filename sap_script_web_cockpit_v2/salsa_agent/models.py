"""
Salsa IT Agent - Models and Data Structures for Phase 1.
"""
from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, List, Optional


class ClassificationCode(str, Enum):
    CREATE_USER = "CREATE_USER"
    UNLOCK_USER = "UNLOCK_USER"
    RESET_PASSWORD = "RESET_PASSWORD"
    OTHER_ACCESS = "OTHER_ACCESS"
    POSSIBLE_MISCATEGORIZATION = "POSSIBLE_MISCATEGORIZATION"
    AMBIGUOUS = "AMBIGUOUS"
    INSUFFICIENT_INFORMATION = "INSUFFICIENT_INFORMATION"


class AnalysisLevel(str, Enum):
    LEVEL_1 = "LEVEL_1"
    LEVEL_2 = "LEVEL_2"


class AnalysisStatus(str, Enum):
    PENDING = "PENDING"
    ANALYZING = "ANALYZING"
    LEVEL_1_COMPLETED = "LEVEL_1_COMPLETED"
    WAITING_AI = "WAITING_AI"
    LEVEL_2_COMPLETED = "LEVEL_2_COMPLETED"
    REVISION_REQUIRED = "REVISION_REQUIRED"
    ERROR = "ERROR"
    AI_NOT_CONFIGURED = "AI_NOT_CONFIGURED"


class OperationalStatus(str, Enum):
    READY_FOR_REVIEW = "READY_FOR_REVIEW"
    WAITING_FOR_FEEDBACK = "WAITING_FOR_FEEDBACK"
    MISSING_INFORMATION = "MISSING_INFORMATION"
    REQUIRES_APPROVAL = "REQUIRES_APPROVAL"
    UNKNOWN = "UNKNOWN"


@dataclass
class ExtractedEntities:
    sap_user_id: Optional[str] = None
    user_name: Optional[str] = None
    reference_user: Optional[str] = None
    requester: Optional[str] = None
    sap_environment: List[str] = field(default_factory=list)
    company_code: Optional[str] = None
    request_reason: Optional[str] = None
    requested_operation: Optional[str] = None
    missing_information: List[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return dataclasses.asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ExtractedEntities:
        if not data:
            return cls()
        return cls(
            sap_user_id=data.get("sap_user_id"),
            user_name=data.get("user_name"),
            reference_user=data.get("reference_user"),
            requester=data.get("requester"),
            sap_environment=list(data.get("sap_environment") or []),
            company_code=data.get("company_code"),
            request_reason=data.get("request_reason"),
            requested_operation=data.get("requested_operation"),
            missing_information=list(data.get("missing_information") or []),
        )


@dataclass
class AnalysisResult:
    ticket_key: str
    classification: ClassificationCode
    level: AnalysisLevel
    status: AnalysisStatus
    intent: str
    justification: str
    operational_status: OperationalStatus = OperationalStatus.READY_FOR_REVIEW
    identified_dependency: Optional[str] = None
    evidences: List[str] = field(default_factory=list)
    extracted_data: ExtractedEntities = field(default_factory=ExtractedEntities)
    missing_info: List[str] = field(default_factory=list)
    suggested_next_action: str = ""
    requires_human_review: bool = False
    confidence: float = 1.0
    rules_version: str = "1.1.0"
    ai_model_version: Optional[str] = None
    content_hash: str = ""
    error_message: Optional[str] = None
    analyzed_at: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "ticket_key": self.ticket_key,
            "classification": self.classification.value if isinstance(self.classification, ClassificationCode) else str(self.classification),
            "level": self.level.value if isinstance(self.level, AnalysisLevel) else str(self.level),
            "status": self.status.value if isinstance(self.status, AnalysisStatus) else str(self.status),
            "operational_status": self.operational_status.value if isinstance(self.operational_status, OperationalStatus) else str(self.operational_status),
            "identified_dependency": self.identified_dependency,
            "intent": self.intent,
            "justification": self.justification,
            "evidences": self.evidences,
            "extracted_data": self.extracted_data.to_dict() if isinstance(self.extracted_data, ExtractedEntities) else (self.extracted_data or {}),
            "missing_info": self.missing_info,
            "suggested_next_action": self.suggested_next_action,
            "requires_human_review": self.requires_human_review,
            "confidence": self.confidence,
            "rules_version": self.rules_version,
            "ai_model_version": self.ai_model_version,
            "content_hash": self.content_hash,
            "error_message": self.error_message,
            "analyzed_at": self.analyzed_at,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> AnalysisResult:
        raw_classification = data.get("classification", ClassificationCode.AMBIGUOUS.value)
        raw_level = data.get("level", AnalysisLevel.LEVEL_1.value)
        raw_status = data.get("status", AnalysisStatus.LEVEL_1_COMPLETED.value)
        raw_op_status = data.get("operational_status", OperationalStatus.READY_FOR_REVIEW.value)

        try:
            classification = ClassificationCode(raw_classification)
        except ValueError:
            classification = ClassificationCode.AMBIGUOUS

        try:
            level = AnalysisLevel(raw_level)
        except ValueError:
            level = AnalysisLevel.LEVEL_1

        try:
            status = AnalysisStatus(raw_status)
        except ValueError:
            status = AnalysisStatus.LEVEL_1_COMPLETED

        try:
            operational_status = OperationalStatus(raw_op_status)
        except ValueError:
            operational_status = OperationalStatus.UNKNOWN

        extracted_raw = data.get("extracted_data") or {}
        extracted_data = ExtractedEntities.from_dict(extracted_raw)

        return cls(
            ticket_key=data.get("ticket_key", ""),
            classification=classification,
            level=level,
            status=status,
            operational_status=operational_status,
            identified_dependency=data.get("identified_dependency"),
            intent=data.get("intent", ""),
            justification=data.get("justification", ""),
            evidences=list(data.get("evidences") or []),
            extracted_data=extracted_data,
            missing_info=list(data.get("missing_info") or []),
            suggested_next_action=data.get("suggested_next_action", ""),
            requires_human_review=bool(data.get("requires_human_review", False)),
            confidence=float(data.get("confidence", 1.0)),
            rules_version=data.get("rules_version", "1.1.0"),
            ai_model_version=data.get("ai_model_version"),
            content_hash=data.get("content_hash", ""),
            error_message=data.get("error_message"),
            analyzed_at=data.get("analyzed_at", ""),
        )


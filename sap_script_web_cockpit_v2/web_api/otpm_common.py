"""
otpm_common.py - Modelos Pydantic, validação de entradas e armazenamento em memória para o fluxo OTPM.
Reutilizado pelas rotas /api/salsa-it-agent/otpm/* em web_api/main.py.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import HTTPException
from pydantic import BaseModel, Field

# Armazenamento em memória das preparações e pré-visualizações validadas
OTPM_EXCEL_SELECTIONS: Dict[str, Dict[str, Any]] = {}
OTPM_DRY_RUN_PREVIEWS: Dict[str, Dict[str, Any]] = {}

OTPM_ALLOWED_ENVIRONMENTS = ("DEV",)


class OtpmValidateRequest(BaseModel):
    environment: str = Field(default="DEV")
    selection_id: Optional[str] = Field(default=None)
    excel_path: Optional[str] = Field(default=None)


class OtpmDryRunRequest(BaseModel):
    environment: str = Field(default="DEV")
    selection_id: Optional[str] = Field(default=None)
    excel_path: Optional[str] = Field(default=None)


class OtpmConfirmRequest(BaseModel):
    dry_run_job_id: str = Field(...)
    request_number: Optional[str] = Field(default="")
    request_description: Optional[str] = Field(default="")


def validate_otpm_environment_or_400(env: str) -> str:
    normalized = str(env or "DEV").strip().upper()
    if normalized not in OTPM_ALLOWED_ENVIRONMENTS:
        raise HTTPException(
            status_code=400,
            detail=f"Ambiente '{normalized}' não permitido para criação/atribuição OTPM. No momento apenas 'DEV' é suportado.",
        )
    return normalized


def safe_otpm_failed_message() -> str:
    return "Não foi possível concluir a operação das Cadeias de Pesquisa OTPM."

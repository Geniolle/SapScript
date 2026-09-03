"""Helpers HTTP genericos, partilhados por todas as rotas."""
from __future__ import annotations

from typing import Any

from fastapi.responses import JSONResponse


def _json_no_store(payload: dict[str, Any], status_code: int = 200) -> JSONResponse:
    response = JSONResponse(content=payload, status_code=status_code)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

from __future__ import annotations

import tempfile
from pathlib import Path

from taxas_cambiais_process import (
    _dedupe_path,
    _extract_folder_segments,
    _message_folder_name,
    sanitize_filename,
)


def test_sanitize_filename_removes_forbidden_characters() -> None:
    assert sanitize_filename("Taxas: Cambiais / Janeiro.xlsx") == "Taxas Cambiais Janeiro.xlsx"


def test_extract_folder_segments_supports_nested_paths() -> None:
    assert _extract_folder_segments(r"Root\Diarios/Taxas Cambiais") == ["Root", "Diarios", "Taxas Cambiais"]


def test_message_folder_name_contains_date_subject_and_id() -> None:
    message = {
        "receivedDateTime": "2026-08-25T10:11:12Z",
        "subject": "Taxas Cambiais Agosto",
        "id": "msg-123",
    }
    folder_name = _message_folder_name(message)
    assert folder_name.startswith("20260825_101112_")
    assert "Taxas_Cambiais_Agosto" in folder_name
    assert "msg-123" in folder_name


def test_dedupe_path_creates_unique_candidate() -> None:
    with tempfile.TemporaryDirectory() as temp_dir:
        base = Path(temp_dir)
        first = base / "anexo.xlsx"
        first.write_text("x", encoding="utf-8")

        candidate = _dedupe_path(first)
        assert candidate.name == "anexo (2).xlsx"

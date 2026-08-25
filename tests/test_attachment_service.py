"""
Unit tests for sap_agent/attachment_service.py

Covers:
  - sanitize_filename
  - validate_extension
  - validate_size
  - get_best_xlsx (selection by Jira metadata)
  - load_manifest / save_manifest (roundtrip)
  - process_ticket_attachments compatibility
    (download_ticket_attachments_to_dir is an integration test — covered separately)
"""

import json
from pathlib import Path

import pytest

from sap_agent.attachment_service import (
    AttachmentMeta,
    AttachmentResult,
    _cache_key,
    get_best_xlsx,
    load_manifest,
    save_manifest,
    sanitize_filename,
    validate_extension,
    validate_size,
)


# ---------------------------------------------------------------------------
# sanitize_filename
# ---------------------------------------------------------------------------

class TestSanitizeFilename:
    def test_simple_name_unchanged(self):
        assert sanitize_filename("relatorio.xlsx") == "relatorio.xlsx"

    def test_unicode_normalized(self):
        result = sanitize_filename("relatório_março.xlsx")
        assert result.endswith(".xlsx")
        assert "/" not in result
        assert "\\" not in result
        assert ".." not in result

    def test_forbidden_windows_chars_replaced(self):
        result = sanitize_filename('file<>:*?|"name.txt')
        for ch in '<>:*?|"':
            assert ch not in result
        assert result.endswith(".txt")

    def test_path_traversal_blocked(self):
        result = sanitize_filename("../../../etc/passwd")
        assert ".." not in result
        assert "/" not in result

    def test_empty_returns_default(self):
        assert sanitize_filename("") == "anexo_sem_nome"
        assert sanitize_filename("   ") == "anexo_sem_nome"

    def test_only_dots_and_spaces(self):
        result = sanitize_filename("... ...")
        assert result  # must not be empty
        assert result != "..."

    def test_stem_limited_to_80_chars(self):
        long_name = "a" * 200 + ".xlsx"
        result = sanitize_filename(long_name)
        stem = Path(result).stem
        assert len(stem) <= 80
        assert result.endswith(".xlsx")

    def test_extension_preserved(self):
        assert sanitize_filename("dados.PDF").endswith(".PDF") or \
               sanitize_filename("dados.PDF").endswith(".pdf")

    def test_no_leading_trailing_dots(self):
        result = sanitize_filename(".hidden.txt")
        assert not result.startswith(".")


# ---------------------------------------------------------------------------
# validate_extension
# ---------------------------------------------------------------------------

class TestValidateExtension:
    _ALLOWED = {".xlsx", ".xlsm", ".pdf", ".png", ".txt"}

    def test_allowed_extension(self):
        assert validate_extension("report.xlsx", self._ALLOWED) is True
        assert validate_extension("image.png", self._ALLOWED) is True

    def test_case_insensitive(self):
        assert validate_extension("REPORT.XLSX", self._ALLOWED) is True
        assert validate_extension("image.PNG", self._ALLOWED) is True

    def test_disallowed_extension(self):
        assert validate_extension("virus.exe", self._ALLOWED) is False
        assert validate_extension("script.py", self._ALLOWED) is False
        assert validate_extension("archive.zip", self._ALLOWED) is False

    def test_no_extension(self):
        assert validate_extension("noextension", self._ALLOWED) is False

    def test_empty_allowed_set(self):
        assert validate_extension("report.xlsx", set()) is False

    def test_uses_module_default_when_no_set_given(self):
        # Should not raise; result depends on env default
        result = validate_extension("report.xlsx")
        assert isinstance(result, bool)


# ---------------------------------------------------------------------------
# validate_size
# ---------------------------------------------------------------------------

class TestValidateSize:
    _MAX = 10 * 1024 * 1024  # 10 MB

    def test_within_limit(self):
        assert validate_size(5 * 1024 * 1024, self._MAX) is True

    def test_exactly_at_limit(self):
        assert validate_size(self._MAX, self._MAX) is True

    def test_exceeds_limit(self):
        assert validate_size(self._MAX + 1, self._MAX) is False

    def test_zero_bytes(self):
        assert validate_size(0, self._MAX) is True

    def test_uses_module_default_when_no_max_given(self):
        result = validate_size(1024)
        assert isinstance(result, bool)


# ---------------------------------------------------------------------------
# get_best_xlsx
# ---------------------------------------------------------------------------

class TestGetBestXlsx:
    def test_selects_most_recent_by_created(self):
        attachments = [
            {"id": "1", "filename": "old.xlsx", "created": "2024-01-01T10:00:00.000Z", "size": 1000},
            {"id": "2", "filename": "new.xlsx", "created": "2024-06-01T10:00:00.000Z", "size": 2000},
            {"id": "3", "filename": "report.pdf", "created": "2024-07-01T10:00:00.000Z", "size": 500},
        ]
        result = get_best_xlsx(attachments)
        assert result is not None
        assert result["filename"] == "new.xlsx"

    def test_tiebreak_uses_id_desc(self):
        attachments = [
            {"id": "100", "filename": "a.xlsx", "created": "2024-06-01T10:00:00.000Z", "size": 1000},
            {"id": "200", "filename": "b.xlsx", "created": "2024-06-01T10:00:00.000Z", "size": 2000},
        ]
        result = get_best_xlsx(attachments)
        assert result is not None
        assert result["id"] == "200"

    def test_no_xlsx_returns_none(self):
        attachments = [
            {"id": "1", "filename": "report.pdf", "created": "2024-06-01T10:00:00.000Z", "size": 500},
            {"id": "2", "filename": "image.png", "created": "2024-06-01T10:00:00.000Z", "size": 100},
        ]
        assert get_best_xlsx(attachments) is None

    def test_empty_list_returns_none(self):
        assert get_best_xlsx([]) is None

    def test_includes_xlsm(self):
        attachments = [
            {"id": "1", "filename": "macro.xlsm", "created": "2024-07-01T10:00:00.000Z", "size": 3000},
            {"id": "2", "filename": "report.xlsx", "created": "2024-06-01T10:00:00.000Z", "size": 2000},
        ]
        result = get_best_xlsx(attachments)
        assert result is not None
        assert result["filename"] == "macro.xlsm"

    def test_prefers_newer_over_larger(self):
        attachments = [
            {"id": "1", "filename": "big.xlsx", "created": "2024-01-01T00:00:00.000Z", "size": 9999},
            {"id": "2", "filename": "small.xlsx", "created": "2024-12-31T00:00:00.000Z", "size": 1},
        ]
        result = get_best_xlsx(attachments)
        assert result is not None
        assert result["filename"] == "small.xlsx"

    def test_ignores_pdf_and_images_with_same_date(self):
        attachments = [
            {"id": "10", "filename": "late.pdf", "created": "2025-01-01T00:00:00.000Z", "size": 100},
            {"id": "5", "filename": "data.xlsx", "created": "2024-06-01T00:00:00.000Z", "size": 200},
        ]
        result = get_best_xlsx(attachments)
        assert result is not None
        assert result["filename"] == "data.xlsx"

    def test_filename_with_unicode_is_sanitized(self):
        attachments = [
            {
                "id": "1",
                "filename": "relatório_mêrcs.xlsx",
                "created": "2024-06-01T00:00:00.000Z",
                "size": 500,
            }
        ]
        result = get_best_xlsx(attachments)
        assert result is not None  # Should be found, not raise


# ---------------------------------------------------------------------------
# Manifest (load / save)
# ---------------------------------------------------------------------------

class TestManifest:
    def test_roundtrip(self, tmp_path):
        ticket_dir = tmp_path / "IZ-99999"
        ticket_dir.mkdir()

        manifest = {
            "ticket_key": "IZ-99999",
            "attachments": {
                "123|file.xlsx|1000|2024-01-01": {
                    "filename": "file.xlsx",
                    "text": "some text",
                    "text_truncated": False,
                }
            },
        }

        save_manifest(ticket_dir, manifest)
        loaded = load_manifest(ticket_dir)

        assert loaded["ticket_key"] == "IZ-99999"
        assert "123|file.xlsx|1000|2024-01-01" in loaded["attachments"]
        assert loaded["attachments"]["123|file.xlsx|1000|2024-01-01"]["text"] == "some text"

    def test_load_missing_dir_returns_empty(self, tmp_path):
        result = load_manifest(tmp_path / "nonexistent")
        assert result == {}

    def test_load_corrupt_json_returns_empty(self, tmp_path):
        ticket_dir = tmp_path / "IZ-BAD"
        ticket_dir.mkdir()
        (ticket_dir / "manifest.json").write_text("NOT_JSON", encoding="utf-8")
        result = load_manifest(ticket_dir)
        assert result == {}

    def test_save_creates_file(self, tmp_path):
        ticket_dir = tmp_path / "IZ-NEW"
        ticket_dir.mkdir()
        save_manifest(ticket_dir, {"ticket_key": "IZ-NEW", "attachments": {}})
        manifest_file = ticket_dir / "manifest.json"
        assert manifest_file.exists()
        data = json.loads(manifest_file.read_text(encoding="utf-8"))
        assert data["ticket_key"] == "IZ-NEW"


# ---------------------------------------------------------------------------
# _cache_key
# ---------------------------------------------------------------------------

class TestCacheKey:
    def test_format(self):
        att = AttachmentMeta(
            attachment_id="789",
            filename="data.xlsx",
            size=4096,
            created="2024-06-01T10:00:00.000Z",
            mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            content_url="https://example.com/att/789",
        )
        key = _cache_key(att)
        assert "789" in key
        assert "data.xlsx" in key
        assert "4096" in key
        assert "2024-06-01" in key

    def test_different_attachments_produce_different_keys(self):
        att1 = AttachmentMeta("1", "a.xlsx", 100, "2024-01-01", "", "")
        att2 = AttachmentMeta("2", "b.xlsx", 200, "2024-02-01", "", "")
        assert _cache_key(att1) != _cache_key(att2)

    def test_same_attachment_produces_same_key(self):
        att = AttachmentMeta("1", "a.xlsx", 100, "2024-01-01", "", "")
        assert _cache_key(att) == _cache_key(att)


# ---------------------------------------------------------------------------
# process_ticket_attachments — without real Jira (mock download)
# ---------------------------------------------------------------------------

class TestProcessTicketAttachments:
    """
    Tests the full pipeline with a fake download function patched in,
    verifying that cache, validation, extraction, and manifest logic work.
    """

    def test_skips_disallowed_extension(self, tmp_path, monkeypatch):
        from sap_agent import attachment_service

        monkeypatch.setattr(attachment_service, "CACHE_BASE_DIR", str(tmp_path))

        attachments = [
            {"id": "1", "filename": "virus.exe", "size": 100, "created": "2024-01-01", "content": "http://x"}
        ]
        results = attachment_service.process_ticket_attachments(
            ticket_key="IZ-TEST",
            attachments_meta=attachments,
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )
        assert len(results) == 1
        assert results[0].skipped is True
        assert "extensão" in results[0].skip_reason

    def test_skips_oversized_attachment(self, tmp_path, monkeypatch):
        from sap_agent import attachment_service

        monkeypatch.setattr(attachment_service, "MAX_SIZE_BYTES", 1024)  # 1 KB max

        attachments = [
            {
                "id": "2",
                "filename": "big.xlsx",
                "size": 999999,
                "created": "2024-01-01",
                "content": "http://x",
            }
        ]
        results = attachment_service.process_ticket_attachments(
            ticket_key="IZ-TEST",
            attachments_meta=attachments,
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )
        assert results[0].skipped is True
        assert "tamanho" in results[0].skip_reason

    def test_text_truncation(self, tmp_path, monkeypatch):
        from sap_agent import attachment_service

        # Simulate a pre-existing cached file with lots of text
        ticket_dir = tmp_path / "IZ-TRUNC"
        original_dir = ticket_dir / "original"
        original_dir.mkdir(parents=True)

        big_text = "x" * 10_000
        txt_file = original_dir / "big.txt"
        txt_file.write_bytes(big_text.encode("utf-8"))

        # Pre-populate manifest so we skip the network download
        att_meta = {
            "id": "99",
            "filename": "big.txt",
            "size": len(big_text),
            "created": "2024-01-01",
            "mimeType": "text/plain",
            "content": "",
        }
        ck = "99|big.txt|10000|2024-01-01"
        manifest = {"ticket_key": "IZ-TRUNC", "attachments": {ck: {}}}
        save_manifest(ticket_dir, manifest)

        results = attachment_service.process_ticket_attachments(
            ticket_key="IZ-TRUNC",
            attachments_meta=[att_meta],
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )

        assert len(results) == 1
        r = results[0]
        assert r.text_truncated is True
        assert len(r.text) == attachment_service.ATTACHMENT_TEXT_MAX_CHARS

    def test_per_attachment_error_does_not_abort_others(self, tmp_path, monkeypatch):
        from sap_agent import attachment_service

        calls = []

        def mock_download(url, auth, target, max_bytes):
            calls.append(url)
            if "fail" in url:
                raise ConnectionError("network error")
            target.write_bytes(b"good content")

        monkeypatch.setattr(attachment_service, "_download_streaming", mock_download)

        attachments = [
            {"id": "1", "filename": "fail.txt", "size": 100, "created": "2024-01-01", "content": "http://fail"},
            {"id": "2", "filename": "ok.txt", "size": 100, "created": "2024-01-02", "content": "http://ok"},
        ]
        results = attachment_service.process_ticket_attachments(
            ticket_key="IZ-ERR",
            attachments_meta=attachments,
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )

        assert len(results) == 2
        assert results[0].error != ""      # fail.txt recorded error
        assert results[1].error == ""      # ok.txt processed fine
        assert results[1].text == "good content"

    def test_manifest_created_after_processing(self, tmp_path, monkeypatch):
        from sap_agent import attachment_service

        def mock_download(url, auth, target, max_bytes):
            target.write_bytes(b"hello world")

        monkeypatch.setattr(attachment_service, "_download_streaming", mock_download)

        attachments = [
            {"id": "5", "filename": "note.txt", "size": 11, "created": "2024-03-01", "content": "http://x"}
        ]
        attachment_service.process_ticket_attachments(
            ticket_key="IZ-MANIFEST",
            attachments_meta=attachments,
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )

        manifest = load_manifest(tmp_path / "IZ-MANIFEST")
        assert manifest.get("ticket_key") == "IZ-MANIFEST"
        assert len(manifest.get("attachments", {})) == 1

    def test_cache_hit_avoids_redownload(self, tmp_path, monkeypatch):
        from sap_agent import attachment_service

        download_count = {"n": 0}

        def mock_download(url, auth, target, max_bytes):
            download_count["n"] += 1
            target.write_bytes(b"cached text")

        monkeypatch.setattr(attachment_service, "_download_streaming", mock_download)

        att_meta = [
            {"id": "7", "filename": "cached.txt", "size": 11, "created": "2024-01-01", "content": "http://x"}
        ]

        # First call — downloads
        attachment_service.process_ticket_attachments(
            ticket_key="IZ-CACHE",
            attachments_meta=att_meta,
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )

        # Second call — should use cache
        attachment_service.process_ticket_attachments(
            ticket_key="IZ-CACHE",
            attachments_meta=att_meta,
            auth=("u", "p"),
            cache_base_dir=str(tmp_path),
        )

        assert download_count["n"] == 1  # downloaded only once

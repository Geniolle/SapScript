from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import requests

from .attachment_service import process_ticket_attachments
from .config import JiraConfig
from .models import TicketContext


@dataclass
class JiraClient:
    config: JiraConfig

    @property
    def auth(self) -> tuple[str, str]:
        return self.config.email, self.config.api_token

    def search_tickets(self) -> list[TicketContext]:
        url = f"{self.config.base_url}/rest/api/3/search/jql"
        payload = {
            "jql": self.config.jql,
            "maxResults": self.config.max_results,
            "fields": ["summary", "description", "comment", "labels", "components", "attachment"],
        }
        response = requests.post(url, json=payload, auth=self.auth, timeout=30)
        response.raise_for_status()
        data = response.json()
        return [self._to_ticket_context(issue) for issue in data.get("issues", [])]

    def add_comment(self, issue_key: str, body: str) -> None:
        url = f"{self.config.base_url}/rest/api/3/issue/{issue_key}/comment"
        response = requests.post(url, json={"body": self._to_adf(body)}, auth=self.auth, timeout=30)
        response.raise_for_status()

    def _to_ticket_context(self, issue: dict[str, Any]) -> TicketContext:
        fields = issue.get("fields", {})
        comments = fields.get("comment", {}).get("comments", [])
        ticket_key = str(issue.get("key") or "")

        description = self._plain_text(fields.get("description"))
        attachments_raw = fields.get("attachment", []) or []
        attachment_names = [
            att.get("filename", "") for att in attachments_raw if att.get("filename")
        ]

        # Download, cache and extract text for all attachments via the central service.
        # Errors per attachment are recorded inside the service; they do not raise here.
        att_texts: list[str] = []
        try:
            results = process_ticket_attachments(
                ticket_key=ticket_key,
                attachments_meta=attachments_raw,
                auth=self.auth,
            )
            for r in results:
                if r.skipped or r.error:
                    if r.error:
                        att_texts.append(
                            f"--- [Erro de Extração: {r.filename}] ---\n{r.error}"
                        )
                    continue
                if r.text:
                    header = f"--- [Texto extraído do anexo: {r.filename}]"
                    if r.text_truncated:
                        header += " [TRUNCADO]"
                    header += " ---"
                    att_texts.append(f"{header}\n{r.text}")
        except Exception as exc:
            print(f"[JIRA CLIENT] Aviso ao processar anexos de {ticket_key}: {exc}")

        # Augment description so that SapDiagnosisEngine / extract_signal can
        # see the attachment content without requiring interface changes.
        if att_texts:
            description += "\n\n" + "\n\n".join(att_texts)

        return TicketContext(
            key=ticket_key,
            summary=str(fields.get("summary") or ""),
            description=description,
            comments=[self._plain_text(c.get("body")) for c in comments],
            labels=list(fields.get("labels") or []),
            components=[c.get("name", "") for c in fields.get("components", [])],
            attachments=attachment_names,
            attachment_texts=att_texts,
            raw=issue,
        )

    def _plain_text(self, value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, str):
            return value
        if isinstance(value, dict):
            if "text" in value:
                return str(value["text"])
            return "\n".join(self._plain_text(item) for item in value.get("content", []))
        if isinstance(value, list):
            return "\n".join(self._plain_text(item) for item in value)
        return str(value)

    def _to_adf(self, text: str) -> dict[str, Any]:
        paragraphs = []
        for line in text.splitlines() or [text]:
            paragraphs.append({
                "type": "paragraph",
                "content": [{"type": "text", "text": line or " "}],
            })
        return {"type": "doc", "version": 1, "content": paragraphs}

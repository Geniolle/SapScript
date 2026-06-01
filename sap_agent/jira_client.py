from __future__ import annotations

from dataclasses import dataclass
from typing import Any

import requests

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
            "fields": ["summary", "description", "comment", "labels", "components"],
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
        return TicketContext(
            key=issue.get("key", ""),
            summary=str(fields.get("summary") or ""),
            description=self._plain_text(fields.get("description")),
            comments=[self._plain_text(comment.get("body")) for comment in comments],
            labels=list(fields.get("labels") or []),
            components=[component.get("name", "") for component in fields.get("components", [])],
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

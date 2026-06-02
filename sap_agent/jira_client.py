from __future__ import annotations

from dataclasses import dataclass
import io
from typing import Any

from PIL import Image
import requests
import winocr

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
        
        description = self._plain_text(fields.get("description"))
        attachments = fields.get("attachment", []) or []
        
        # Extrair texto das capturas de ecrã (prints) anexadas
        image_text = self._extract_text_from_images(attachments)
        if image_text:
            description += f"\n\n{image_text}"

        return TicketContext(
            key=issue.get("key", ""),
            summary=str(fields.get("summary") or ""),
            description=description,
            comments=[self._plain_text(comment.get("body")) for comment in comments],
            labels=list(fields.get("labels") or []),
            components=[component.get("name", "") for component in fields.get("components", [])],
            raw=issue,
        )

    def _extract_text_from_images(self, attachments: list[dict[str, Any]]) -> str:
        extracted_texts = []
        for att in attachments:
            mime_type = str(att.get("mimeType") or "").lower()
            filename = str(att.get("filename") or "").lower()
            
            # Verificar se é uma imagem (png, jpg, jpeg, gif)
            if not (mime_type.startswith("image/") or filename.endswith((".png", ".jpg", ".jpeg", ".gif"))):
                continue
                
            content_url = att.get("content")
            if not content_url:
                continue
                
            try:
                # Descarregar o anexo de imagem em memória
                res = requests.get(content_url, auth=self.auth, timeout=15)
                res.raise_for_status()
                
                # Ler com Pillow
                image = Image.open(io.BytesIO(res.content))
                
                # Executar o Windows OCR
                ocr_result = winocr.recognize_pil_sync(image)
                text = str(ocr_result.get("text") or "").strip()
                if text:
                    extracted_texts.append(f"--- [Texto extraído do print: {att.get('filename')}] ---\n{text}")
            except Exception as e:
                print(f"[OCR WARNING] Erro ao processar anexo {att.get('filename')}: {e}")
                
        return "\n\n".join(extracted_texts)

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

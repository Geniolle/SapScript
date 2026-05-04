from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from sap_session import ensure_sap_access_from_env


BOOL_TRUE = {"1", "true", "yes", "on", "sim", "s"}


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in BOOL_TRUE


def _safe_name(value: str) -> str:
    text = re.sub(r"[^a-zA-Z0-9._-]+", "_", str(value or "").strip())
    text = re.sub(r"_+", "_", text).strip("._")
    return text or "step"


def _resolve_doc_output_dir(base_dir: Path, row_context: Dict[str, str]) -> Path:
    xlsx_path = Path(str(row_context.get("xlsx_path", "")).strip())
    if xlsx_path.exists() and xlsx_path.parent.exists():
        return xlsx_path.parent.resolve()

    template = str(os.getenv("WORKFLOW_DOC_OUTPUT_DIR", "") or "").strip()
    if template:
        rendered = template.format_map({k: str(v or "") for k, v in row_context.items()})
        folder = Path(rendered)
        if not folder.is_absolute():
            folder = (base_dir / folder).resolve()
        return folder

    return (base_dir / "cache" / "documentacao").resolve()


@dataclass
class WorkflowDocumentation:
    enabled: bool
    workflow_name: str
    ticket_key: str
    category: str
    output_dir: Path
    doc_path: Path
    image_dir: Path
    started_at: str
    entries: List[Dict[str, str]] = field(default_factory=list)

    @classmethod
    def from_env(
        cls,
        *,
        base_dir: Path,
        row_context: Dict[str, str],
        workflow_name: str,
    ) -> "WorkflowDocumentation":
        enabled = _to_bool(os.getenv("WORKFLOW_DOC_ENABLED", "false"))
        ticket_key = str(row_context.get("ticket_key", "") or "").strip().upper()
        category = str(row_context.get("categoria_sap", "") or "").strip()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        output_dir = _resolve_doc_output_dir(base_dir, row_context)
        safe_ticket = _safe_name(ticket_key or "ticket")
        safe_category = _safe_name(category or "workflow")
        doc_name = f"Documentacao_{safe_ticket}_{safe_category}_{stamp}.docx"
        image_dir = output_dir / f"_evidencias_{safe_ticket}_{stamp}"

        return cls(
            enabled=enabled,
            workflow_name=workflow_name,
            ticket_key=ticket_key,
            category=category,
            output_dir=output_dir,
            doc_path=(output_dir / doc_name),
            image_dir=image_dir,
            started_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )

    def capture_step(
        self,
        *,
        step_name: str,
        row_context: Dict[str, str],
        note: str = "",
        snapshot_override: Dict[str, str] | None = None,
        allow_live_capture: bool = True,
    ) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        request_number = str(row_context.get("request_number", "") or "").strip().upper()
        status_type = ""
        status_text = ""
        image_path = ""

        if snapshot_override:
            status_type = str(snapshot_override.get("status_type", "") or "").strip().upper()
            status_text = str(snapshot_override.get("status_text", "") or "").strip()
            image_path = str(snapshot_override.get("image_path", "") or "").strip()
            snap_ts = str(snapshot_override.get("timestamp", "") or "").strip()
            if snap_ts:
                timestamp = snap_ts
        elif self.enabled and allow_live_capture:
            session = self._try_get_session()
            if session is not None:
                status_type, status_text = self._read_status_bar(session)
                image_path = self._capture_sap_screen(session, step_name=step_name)

        self.entries.append(
            {
                "timestamp": timestamp,
                "step": step_name,
                "request_number": request_number,
                "status_type": status_type,
                "status_text": status_text,
                "image_path": image_path,
                "note": note,
            }
        )

    def capture_runtime_snapshot(self, *, step_name: str, row_context: Dict[str, str]) -> Dict[str, str]:
        if not self.enabled:
            return {}

        snapshot: Dict[str, str] = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "step": step_name,
            "request_number": str(row_context.get("request_number", "") or "").strip().upper(),
            "status_type": "",
            "status_text": "",
            "image_path": "",
        }

        session = self._try_get_session()
        if session is None:
            return snapshot

        status_type, status_text = self._read_status_bar(session)
        image_path = self._capture_sap_screen(session, step_name=f"{step_name}_config_validacao")
        snapshot["status_type"] = status_type
        snapshot["status_text"] = status_text
        snapshot["image_path"] = image_path
        return snapshot

    def capture_runtime_snapshot_with_retry(
        self,
        *,
        step_name: str,
        row_context: Dict[str, str],
        attempts: int = 4,
        wait_s: float = 0.4,
    ) -> Dict[str, str]:
        snapshot: Dict[str, str] = {}
        total = max(1, int(attempts))
        for idx in range(total):
            snapshot = self.capture_runtime_snapshot(step_name=step_name, row_context=row_context)
            image_path = str(snapshot.get("image_path", "") or "").strip()
            if image_path and Path(image_path).exists():
                return snapshot
            if idx < total - 1:
                import time
                time.sleep(max(0.1, float(wait_s)))
        return snapshot

    def finalize(self, *, row_context: Dict[str, str], success: bool, error: str = "") -> str:
        if not self.enabled:
            return ""

        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self.image_dir.mkdir(parents=True, exist_ok=True)
            self._build_word_document(row_context=row_context, success=success, error=error)
            return str(self.doc_path)
        except Exception as exc:
            logging.warning("Falha ao gerar documento Word de evidencias: %s", exc)
            return ""

    def _try_get_session(self):
        try:
            return ensure_sap_access_from_env(
                key=os.getenv("WORKFLOW_SAP_KEY", "S4DCLNT100"),
                timeout_s=8,
                load_env=False,
            )
        except Exception as exc:
            logging.warning("Nao foi possivel obter sessao SAP para screenshot: %s", exc)
            return None

    def _read_status_bar(self, session) -> tuple[str, str]:
        try:
            sbar = session.findById("wnd[0]/sbar")
        except Exception:
            return "", ""
        msg_type = str(getattr(sbar, "MessageType", "") or "").strip().upper()
        msg_text = str(getattr(sbar, "Text", "") or "").strip()
        return msg_type, msg_text

    def _capture_sap_screen(self, session, *, step_name: str) -> str:
        try:
            self.image_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.now().strftime("%H%M%S")
            file_name = f"{len(self.entries) + 1:02d}_{_safe_name(step_name)}_{stamp}.bmp"
            image_path = self.image_dir / file_name
            wnd = session.findById("wnd[0]")

            try:
                wnd.hardCopy(str(image_path), 2)
            except Exception:
                try:
                    wnd.HardCopy(str(image_path), 2)
                except Exception:
                    try:
                        wnd.hardCopy(str(image_path))
                    except Exception:
                        wnd.HardCopy(str(image_path))

            return str(image_path) if image_path.exists() else ""
        except Exception as exc:
            logging.warning("Falha ao tirar screenshot SAP no step '%s': %s", step_name, exc)
            return ""

    def _build_word_document(self, *, row_context: Dict[str, str], success: bool, error: str) -> None:
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError("pywin32 nao disponivel para gerar Word automaticamente.") from exc

        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            doc = app.Documents.Add()
            sel = app.Selection

            def _line(text: str = ""):
                sel.TypeText(str(text))
                sel.TypeParagraph()

            sel.Font.Bold = True
            _line("Documentacao de Configuracao SAP")
            sel.Font.Bold = False
            _line(f"Data/Hora de inicio: {self.started_at}")
            _line(f"Workflow: {self.workflow_name}")
            _line(f"Ticket: {self.ticket_key or '-'}")
            _line(f"Categoria SAP: {self.category or '-'}")
            _line(f"Ordem de Transporte: {row_context.get('request_number', '') or '-'}")
            _line(f"Resultado final: {'CONCLUIDO' if success else 'FALHOU'}")
            if error:
                _line(f"Erro: {error}")
            _line("")

            for index, entry in enumerate(self.entries, start=1):
                sel.Font.Bold = True
                _line(f"Passo {index}: {entry.get('step', '-')}")
                sel.Font.Bold = False
                _line(f"Momento: {entry.get('timestamp', '-')}")
                _line(f"Ordem de Transporte: {entry.get('request_number', '-') or '-'}")
                msg_type = entry.get("status_type", "") or "-"
                msg_text = entry.get("status_text", "") or "-"
                _line(f"Status SAP: [{msg_type}] {msg_text}")
                note = entry.get("note", "") or ""
                if note:
                    _line(f"Observacao: {note}")

                image_path = str(entry.get("image_path", "") or "").strip()
                if image_path and Path(image_path).exists():
                    sel.InlineShapes.AddPicture(
                        FileName=str(Path(image_path).resolve()),
                        LinkToFile=False,
                        SaveWithDocument=True,
                    )
                    sel.TypeParagraph()
                else:
                    _line("Captura de ecra: nao disponivel")
                _line("")

            doc.SaveAs(str(self.doc_path), FileFormat=12)
        finally:
            if doc is not None:
                doc.Close(False)
            if app is not None:
                app.Quit()
            pythoncom.CoUninitialize()

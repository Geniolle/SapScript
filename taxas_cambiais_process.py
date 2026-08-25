# -*- coding: utf-8 -*-
"""
Taxas Cambiais

Processo diario para:
    1. Ligar ao Microsoft Graph com credenciais do .env
    2. Localizar a pasta de origem no Outlook/Office 365
    3. Descarregar os anexos dos emails encontrados
    4. Guardar os anexos na pasta local configurada
    5. Mover os emails para a pasta de backup

O processo foi desenhado para correr de forma repetivel em execucao diaria.
Nao guarda segredos no codigo e nao depende de browser manual.
"""

from __future__ import annotations

import argparse
import base64
import json
import logging
import os
import re
import sys
from dataclasses import dataclass, field, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
from dotenv import load_dotenv


# =============================================================================
# (1) CONFIGURACAO BASE
# =============================================================================

ROOT_DIR = Path(__file__).resolve().parent
GRAPH_BASE_URL_DEFAULT = "https://graph.microsoft.com/v1.0"
AUTHORITY_URL_DEFAULT = "https://login.microsoftonline.com"
GRAPH_SCOPE_DEFAULT = "https://graph.microsoft.com/.default"
SOURCE_FOLDER_DEFAULT = "Diarios"
BACKUP_FOLDER_DEFAULT = "Backup Taxas"
OUTPUT_DIR_DEFAULT = Path.home() / "Downloads" / "Taxas Cambiais"
REQUEST_TIMEOUT_S = 60
MAX_MESSAGES_PER_RUN_DEFAULT = 200
MAX_ATTACHMENTS_PER_MESSAGE_DEFAULT = 50


if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _print(*args: Any, **kwargs: Any) -> None:
    kwargs.setdefault("flush", True)
    print(*args, **kwargs)


@dataclass(frozen=True)
class TaxasCambiaisConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    mailbox_upn: str
    source_folder: str
    backup_folder: str
    output_dir: Path
    graph_base_url: str = GRAPH_BASE_URL_DEFAULT
    authority_url: str = AUTHORITY_URL_DEFAULT
    graph_scope: str = GRAPH_SCOPE_DEFAULT
    skip_inline_attachments: bool = False
    only_unread: bool = False
    max_messages: int = MAX_MESSAGES_PER_RUN_DEFAULT
    max_attachments_per_message: int = MAX_ATTACHMENTS_PER_MESSAGE_DEFAULT
    page_size: int = 50


@dataclass
class AttachmentDownload:
    name: str
    local_path: str
    size: int
    content_type: str
    is_inline: bool


@dataclass
class MessageRunResult:
    message_id: str
    subject: str
    received_date_time: str
    source_folder: str
    backup_folder: str
    attachments: list[AttachmentDownload] = field(default_factory=list)
    moved: bool = False
    skipped: bool = False
    error: str = ""


@dataclass
class RunSummary:
    mailbox_upn: str
    source_folder: str
    backup_folder: str
    output_dir: str
    processed_messages: int = 0
    moved_messages: int = 0
    downloaded_attachments: int = 0
    skipped_messages: int = 0
    failed_messages: int = 0
    messages: list[MessageRunResult] = field(default_factory=list)


# =============================================================================
# (2) CONFIGURACAO E VALIDACOES
# =============================================================================

def load_project_dotenv() -> None:
    """Carrega o .env do projeto sem imprimir segredos."""

    candidates = [
        ROOT_DIR / ".env",
        Path.cwd() / ".env",
    ]

    for env_path in candidates:
        if env_path.exists():
            load_dotenv(env_path, override=False)


def _first_env(*names: str, default: str = "", required: bool = True) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value

    if required:
        joined = ", ".join(names)
        raise RuntimeError(f"Falta definir uma das variaveis de ambiente: {joined}")

    return default


def _to_bool(value: str | bool | None) -> bool:
    if isinstance(value, bool):
        return value
    return str(value or "").strip().lower() in {"1", "true", "yes", "on", "sim", "s"}


def _parse_int(value: str | int | None, default: int) -> int:
    try:
        if value is None:
            return default
        return int(str(value).strip())
    except Exception:
        return default


def _resolve_output_dir(raw_value: str | None) -> Path:
    value = str(raw_value or "").strip()
    if not value:
        return OUTPUT_DIR_DEFAULT
    return Path(value).expanduser().resolve()


def load_config() -> TaxasCambiaisConfig:
    load_project_dotenv()

    tenant_id = _first_env("TAXAS_CAMBIAIS_TENANT_ID")
    client_id = _first_env("TAXAS_CAMBIAIS_CLIENT_ID")
    client_secret = _first_env("TAXAS_CAMBIAIS_CLIENT_SECRET")
    mailbox_upn = _first_env("TAXAS_CAMBIAIS_MAILBOX_UPN")
    source_folder = os.getenv("TAXAS_CAMBIAIS_SOURCE_FOLDER", SOURCE_FOLDER_DEFAULT).strip()
    backup_folder = os.getenv("TAXAS_CAMBIAIS_BACKUP_FOLDER", BACKUP_FOLDER_DEFAULT).strip()

    return TaxasCambiaisConfig(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        mailbox_upn=mailbox_upn,
        source_folder=source_folder or SOURCE_FOLDER_DEFAULT,
        backup_folder=backup_folder or BACKUP_FOLDER_DEFAULT,
        output_dir=_resolve_output_dir(os.getenv("TAXAS_CAMBIAIS_OUTPUT_DIR")),
        graph_base_url=os.getenv("TAXAS_CAMBIAIS_GRAPH_BASE_URL", GRAPH_BASE_URL_DEFAULT).strip() or GRAPH_BASE_URL_DEFAULT,
        authority_url=os.getenv("TAXAS_CAMBIAIS_AUTHORITY_URL", AUTHORITY_URL_DEFAULT).strip() or AUTHORITY_URL_DEFAULT,
        graph_scope=os.getenv("TAXAS_CAMBIAIS_GRAPH_SCOPE", GRAPH_SCOPE_DEFAULT).strip() or GRAPH_SCOPE_DEFAULT,
        skip_inline_attachments=_to_bool(os.getenv("TAXAS_CAMBIAIS_SKIP_INLINE_ATTACHMENTS", "false")),
        only_unread=_to_bool(os.getenv("TAXAS_CAMBIAIS_ONLY_UNREAD", "false")),
        max_messages=_parse_int(os.getenv("TAXAS_CAMBIAIS_MAX_MESSAGES"), MAX_MESSAGES_PER_RUN_DEFAULT),
        max_attachments_per_message=_parse_int(
            os.getenv("TAXAS_CAMBIAIS_MAX_ATTACHMENTS_PER_MESSAGE"),
            MAX_ATTACHMENTS_PER_MESSAGE_DEFAULT,
        ),
        page_size=max(1, _parse_int(os.getenv("TAXAS_CAMBIAIS_PAGE_SIZE"), 50)),
    )


def sanitize_filename(value: str, default: str = "ficheiro") -> str:
    text = str(value or "").strip()
    if not text:
        return default

    text = re.sub(r"[<>:\"/\\|?*\x00-\x1f]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip(" .")
    text = re.sub(r"\.{2,}", ".", text)
    return text or default


def _slugify(value: str, *, default: str = "item") -> str:
    text = sanitize_filename(value, default=default)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("._-")
    return text or default


def _safe_iso_datetime(value: str) -> datetime:
    text = str(value or "").strip()
    if not text:
        return datetime.now(timezone.utc)
    normalized = text.replace("Z", "+00:00")
    try:
        return datetime.fromisoformat(normalized)
    except ValueError:
        return datetime.now(timezone.utc)


def _dedupe_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    for index in range(2, 1000):
        candidate = parent / f"{stem} ({index}){suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Nao foi possivel gerar um nome unico para: {path.name}")


def _extract_folder_segments(folder_path: str) -> list[str]:
    raw_segments = re.split(r"[\\/]+", str(folder_path or "").strip())
    return [segment.strip() for segment in raw_segments if segment and segment.strip()]


# =============================================================================
# (3) GRAPH AUTH E CLIENT
# =============================================================================

def _graph_token_url(cfg: TaxasCambiaisConfig) -> str:
    return f"{cfg.authority_url.rstrip('/')}/{cfg.tenant_id}/oauth2/v2.0/token"


def get_graph_access_token(cfg: TaxasCambiaisConfig) -> str:
    missing = []
    if not cfg.tenant_id:
        missing.append("TAXAS_CAMBIAIS_TENANT_ID")
    if not cfg.client_id:
        missing.append("TAXAS_CAMBIAIS_CLIENT_ID")
    if not cfg.client_secret:
        missing.append("TAXAS_CAMBIAIS_CLIENT_SECRET")
    if missing:
        raise RuntimeError(f"Faltam variaveis obrigatorias: {', '.join(missing)}")

    response = requests.post(
        _graph_token_url(cfg),
        data={
            "client_id": cfg.client_id,
            "client_secret": cfg.client_secret,
            "scope": cfg.graph_scope,
            "grant_type": "client_credentials",
        },
        timeout=REQUEST_TIMEOUT_S,
    )
    if response.status_code >= 400:
        raise RuntimeError(
            "Falha ao obter token do Microsoft Graph: "
            f"{response.status_code} {response.text.strip()}"
        )

    payload = response.json()
    token = str(payload.get("access_token") or "").strip()
    if not token:
        raise RuntimeError("Microsoft Graph devolveu um token vazio.")
    return token


def _graph_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }


def graph_request(
    method: str,
    url: str,
    *,
    token: str,
    json_body: dict[str, Any] | None = None,
    params: dict[str, Any] | None = None,
) -> dict[str, Any]:
    response = requests.request(
        method=method.upper(),
        url=url,
        headers=_graph_headers(token),
        json=json_body,
        params=params,
        timeout=REQUEST_TIMEOUT_S,
    )

    if response.status_code >= 400:
        raise RuntimeError(
            f"Microsoft Graph respondeu com erro {response.status_code} em {url}: "
            f"{response.text.strip()}"
        )

    if response.status_code == 204:
        return {}

    if not response.text.strip():
        return {}

    return response.json()


def graph_list_all(url: str, *, token: str) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    next_url: str | None = url

    while next_url:
        payload = graph_request("GET", next_url, token=token)
        raw_items = payload.get("value", [])
        if isinstance(raw_items, list):
            items.extend([item for item in raw_items if isinstance(item, dict)])
        next_url = payload.get("@odata.nextLink")
        if next_url is not None and not isinstance(next_url, str):
            next_url = str(next_url)

    return items


# =============================================================================
# (4) FOLDERS, MENSAGENS E ANEXOS
# =============================================================================

def _mail_folder_collection_url(cfg: TaxasCambiaisConfig, parent_id: str | None = None) -> str:
    if parent_id:
        return f"{cfg.graph_base_url.rstrip('/')}/users/{cfg.mailbox_upn}/mailFolders/{parent_id}/childFolders?$top=200"
    return f"{cfg.graph_base_url.rstrip('/')}/users/{cfg.mailbox_upn}/mailFolders?$top=200"


def _find_folder_in_collection(
    folders: list[dict[str, Any]],
    folder_name: str,
) -> dict[str, Any] | None:
    target = str(folder_name or "").strip().casefold()
    for folder in folders:
        if str(folder.get("displayName") or "").strip().casefold() == target:
            return folder
    return None


def create_mail_folder(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    folder_name: str,
    parent_id: str | None = None,
) -> dict[str, Any]:
    url = _mail_folder_collection_url(cfg, parent_id=parent_id)
    payload = {"displayName": folder_name}
    return graph_request("POST", url, token=token, json_body=payload)


def resolve_mail_folder_id(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    folder_path: str,
    create_missing_leaf: bool = False,
) -> tuple[str, str]:
    segments = _extract_folder_segments(folder_path)
    if not segments:
        raise RuntimeError("A pasta do Outlook nao foi informada.")

    parent_id: str | None = None
    resolved_path: list[str] = []

    for index, segment in enumerate(segments, start=1):
        folder_url = _mail_folder_collection_url(cfg, parent_id=parent_id)
        folders = graph_list_all(folder_url, token=token)
        match = _find_folder_in_collection(folders, segment)

        if match is None and create_missing_leaf and index == len(segments):
            created = create_mail_folder(cfg, token=token, folder_name=segment, parent_id=parent_id)
            match = created if isinstance(created, dict) else None

        if match is None:
            scope = "raiz" if parent_id is None else parent_id
            raise RuntimeError(
                f"Folder '{segment}' nao encontrada em '{folder_path}' (scope: {scope})."
            )

        parent_id = str(match.get("id") or "").strip()
        if not parent_id:
            raise RuntimeError(f"Folder '{segment}' encontrada sem id valido.")
        resolved_path.append(str(match.get("displayName") or segment).strip())

    return parent_id, " / ".join(resolved_path)


def list_messages_in_folder(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    folder_id: str,
) -> list[dict[str, Any]]:
    select = "id,subject,receivedDateTime,hasAttachments,isRead,from,parentFolderId"
    filter_parts = []
    if cfg.only_unread:
        filter_parts.append("isRead eq false")

    params: dict[str, Any] = {
        "$top": cfg.page_size,
        "$select": select,
        "$orderby": "receivedDateTime asc",
    }
    if filter_parts:
        params["$filter"] = " and ".join(filter_parts)

    url = f"{cfg.graph_base_url.rstrip('/')}/users/{cfg.mailbox_upn}/mailFolders/{folder_id}/messages"
    items: list[dict[str, Any]] = []
    next_url: str | None = url

    while next_url and len(items) < cfg.max_messages:
        payload = graph_request("GET", next_url, token=token, params=params if next_url == url else None)
        raw_items = payload.get("value", [])
        if isinstance(raw_items, list):
            for item in raw_items:
                if isinstance(item, dict):
                    items.append(item)
                    if len(items) >= cfg.max_messages:
                        break
        next_url = payload.get("@odata.nextLink")
        if next_url is not None and not isinstance(next_url, str):
            next_url = str(next_url)

    return items[: cfg.max_messages]


def list_message_attachments(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    message_id: str,
) -> list[dict[str, Any]]:
    url = (
        f"{cfg.graph_base_url.rstrip('/')}/users/{cfg.mailbox_upn}"
        f"/messages/{message_id}/attachments"
    )
    attachments = graph_list_all(f"{url}?$top={cfg.max_attachments_per_message}", token=token)
    if not attachments:
        return []

    filtered: list[dict[str, Any]] = []
    for item in attachments:
        if not isinstance(item, dict):
            continue
        if cfg.skip_inline_attachments and _to_bool(item.get("isInline")):
            continue
        filtered.append(item)
    return filtered


def get_attachment_bytes(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    message_id: str,
    attachment_id: str,
) -> tuple[bytes, dict[str, Any]]:
    url = (
        f"{cfg.graph_base_url.rstrip('/')}/users/{cfg.mailbox_upn}"
        f"/messages/{message_id}/attachments/{attachment_id}"
    )
    payload = graph_request("GET", url, token=token)
    content_b64 = str(payload.get("contentBytes") or "").strip()
    if not content_b64:
        raise RuntimeError("Anexo sem contentBytes no Microsoft Graph.")

    try:
        attachment_bytes = base64.b64decode(content_b64, validate=False)
    except Exception as exc:
        raise RuntimeError(f"Falha ao decodificar o anexo: {exc}") from exc

    return attachment_bytes, payload


def move_message_to_folder(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    message_id: str,
    destination_folder_id: str,
) -> dict[str, Any]:
    url = (
        f"{cfg.graph_base_url.rstrip('/')}/users/{cfg.mailbox_upn}"
        f"/messages/{message_id}/move"
    )
    return graph_request("POST", url, token=token, json_body={"destinationId": destination_folder_id})


def _message_folder_name(message: dict[str, Any]) -> str:
    received = _safe_iso_datetime(str(message.get("receivedDateTime") or ""))
    stamp = received.astimezone(timezone.utc).strftime("%Y%m%d_%H%M%S")
    subject = _slugify(str(message.get("subject") or "sem_assunto"), default="sem_assunto")
    message_id = _slugify(str(message.get("id") or "mensagem"), default="mensagem")
    return f"{stamp}_{subject}_{message_id}"


def _ensure_unique_attachment_path(directory: Path, attachment_name: str) -> Path:
    safe_name = _slugify(attachment_name, default="anexo")
    candidate = directory / safe_name
    return _dedupe_path(candidate)


def download_message_attachments(
    cfg: TaxasCambiaisConfig,
    *,
    token: str,
    message: dict[str, Any],
    output_root: Path,
) -> MessageRunResult:
    message_id = str(message.get("id") or "").strip()
    subject = str(message.get("subject") or "").strip()
    received_date_time = str(message.get("receivedDateTime") or "").strip()
    run_result = MessageRunResult(
        message_id=message_id,
        subject=subject,
        received_date_time=received_date_time,
        source_folder=cfg.source_folder,
        backup_folder=cfg.backup_folder,
    )

    message_dir = output_root / _message_folder_name(message)
    attachments_dir = message_dir / "anexos"
    attachments_dir.mkdir(parents=True, exist_ok=True)

    attachments = list_message_attachments(cfg, token=token, message_id=message_id)
    if not attachments:
        run_result.skipped = True
        try:
            move_message_to_folder(
                cfg,
                token=token,
                message_id=message_id,
                destination_folder_id=_resolve_backup_folder_id(cfg, token=token),
            )
            run_result.moved = True
        except Exception as exc:
            run_result.error = f"Email sem anexos, mas falha ao mover: {exc}"
        return run_result

    for attachment in attachments:
        attachment_id = str(attachment.get("id") or "").strip()
        attachment_name = str(attachment.get("name") or "").strip()
        content_type = str(attachment.get("contentType") or "").strip()
        is_inline = _to_bool(attachment.get("isInline"))

        if not attachment_id:
            run_result.error = "Anexo sem id valido."
            return run_result

        try:
            attachment_bytes, attachment_payload = get_attachment_bytes(
                cfg,
                token=token,
                message_id=message_id,
                attachment_id=attachment_id,
            )
            file_name = attachment_name or str(attachment_payload.get("name") or "anexo")
            target_path = _ensure_unique_attachment_path(attachments_dir, file_name)
            target_path.write_bytes(attachment_bytes)
            run_result.attachments.append(
                AttachmentDownload(
                    name=file_name,
                    local_path=str(target_path),
                    size=len(attachment_bytes),
                    content_type=content_type,
                    is_inline=is_inline,
                )
            )
        except Exception as exc:
            run_result.error = f"Falha no anexo '{attachment_name or attachment_id}': {exc}"
            return run_result

    try:
        move_message_to_folder(
            cfg,
            token=token,
            message_id=message_id,
            destination_folder_id=_resolve_backup_folder_id(cfg, token=token),
        )
        run_result.moved = True
    except Exception as exc:
        run_result.error = f"Anexos descarregados, mas falha ao mover email: {exc}"

    return run_result


def _resolve_backup_folder_id(cfg: TaxasCambiaisConfig, *, token: str) -> str:
    folder_id, _ = resolve_mail_folder_id(
        cfg,
        token=token,
        folder_path=cfg.backup_folder,
        create_missing_leaf=True,
    )
    return folder_id


# =============================================================================
# (5) EXECUCAO PRINCIPAL
# =============================================================================

def run_taxas_cambiais(cfg: TaxasCambiaisConfig) -> RunSummary:
    token = get_graph_access_token(cfg)
    cfg.output_dir.mkdir(parents=True, exist_ok=True)

    source_folder_id, resolved_source_path = resolve_mail_folder_id(
        cfg,
        token=token,
        folder_path=cfg.source_folder,
        create_missing_leaf=False,
    )
    backup_folder_id = _resolve_backup_folder_id(cfg, token=token)
    messages = list_messages_in_folder(cfg, token=token, folder_id=source_folder_id)

    summary = RunSummary(
        mailbox_upn=cfg.mailbox_upn,
        source_folder=resolved_source_path,
        backup_folder=cfg.backup_folder,
        output_dir=str(cfg.output_dir),
    )

    if not messages:
        return summary

    for message in messages:
        summary.processed_messages += 1
        try:
            result = download_message_attachments(
                cfg,
                token=token,
                message=message,
                output_root=cfg.output_dir,
            )
            summary.downloaded_attachments += len(result.attachments)
            if result.moved:
                summary.moved_messages += 1
            if result.skipped and not result.attachments:
                summary.skipped_messages += 1
            if result.error:
                summary.failed_messages += 1
            elif not result.moved and result.attachments and backup_folder_id:
                # Se houve anexos e o email nao foi movido, tratamos como falha.
                summary.failed_messages += 1
            summary.messages.append(result)
        except Exception as exc:
            summary.failed_messages += 1
            summary.messages.append(
                MessageRunResult(
                    message_id=str(message.get("id") or "").strip(),
                    subject=str(message.get("subject") or "").strip(),
                    received_date_time=str(message.get("receivedDateTime") or "").strip(),
                    source_folder=cfg.source_folder,
                    backup_folder=cfg.backup_folder,
                    error=str(exc),
                )
            )

    return summary


def _write_run_manifest(cfg: TaxasCambiaisConfig, summary: RunSummary) -> Path:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    manifest_path = cfg.output_dir / f"taxas_cambiais_{timestamp}.json"
    payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "summary": {
            "mailbox_upn": summary.mailbox_upn,
            "source_folder": summary.source_folder,
            "backup_folder": summary.backup_folder,
            "output_dir": summary.output_dir,
            "processed_messages": summary.processed_messages,
            "moved_messages": summary.moved_messages,
            "downloaded_attachments": summary.downloaded_attachments,
            "skipped_messages": summary.skipped_messages,
            "failed_messages": summary.failed_messages,
        },
        "messages": [
            {
                "message_id": item.message_id,
                "subject": item.subject,
                "received_date_time": item.received_date_time,
                "source_folder": item.source_folder,
                "backup_folder": item.backup_folder,
                "attachments": [
                    {
                        "name": att.name,
                        "local_path": att.local_path,
                        "size": att.size,
                        "content_type": att.content_type,
                        "is_inline": att.is_inline,
                    }
                    for att in item.attachments
                ],
                "moved": item.moved,
                "skipped": item.skipped,
                "error": item.error,
            }
            for item in summary.messages
        ],
    }
    manifest_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path


def executar(ambiente_cockpit: str | None = None, **kwargs: Any) -> dict[str, Any]:
    del ambiente_cockpit  # O processo nao depende do cockpit SAP.
    cfg = load_config()

    cfg = replace(
        cfg,
        source_folder=str(kwargs.get("source_folder") or cfg.source_folder).strip(),
        backup_folder=str(kwargs.get("backup_folder") or cfg.backup_folder).strip(),
        output_dir=_resolve_output_dir(str(kwargs.get("output_dir") or cfg.output_dir)),
        only_unread=_to_bool(kwargs.get("only_unread")) if "only_unread" in kwargs else cfg.only_unread,
        skip_inline_attachments=_to_bool(kwargs.get("skip_inline_attachments")) if "skip_inline_attachments" in kwargs else cfg.skip_inline_attachments,
        max_messages=_parse_int(kwargs.get("max_messages"), cfg.max_messages) if kwargs.get("max_messages") is not None else cfg.max_messages,
        page_size=max(1, _parse_int(kwargs.get("page_size"), cfg.page_size)) if kwargs.get("page_size") is not None else cfg.page_size,
    )

    summary = run_taxas_cambiais(cfg)
    manifest_path = _write_run_manifest(cfg, summary)
    return {
        "mailbox_upn": summary.mailbox_upn,
        "source_folder": summary.source_folder,
        "backup_folder": summary.backup_folder,
        "output_dir": summary.output_dir,
        "processed_messages": summary.processed_messages,
        "moved_messages": summary.moved_messages,
        "downloaded_attachments": summary.downloaded_attachments,
        "skipped_messages": summary.skipped_messages,
        "failed_messages": summary.failed_messages,
        "manifest_path": str(manifest_path),
    }


# =============================================================================
# (6) CLI
# =============================================================================

def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Descarrega anexos da pasta Diarios do Office 365 e move os emails para Backup Taxas."
    )
    parser.add_argument("--source-folder", default=os.getenv("TAXAS_CAMBIAIS_SOURCE_FOLDER", SOURCE_FOLDER_DEFAULT))
    parser.add_argument("--backup-folder", default=os.getenv("TAXAS_CAMBIAIS_BACKUP_FOLDER", BACKUP_FOLDER_DEFAULT))
    parser.add_argument("--output-dir", default=os.getenv("TAXAS_CAMBIAIS_OUTPUT_DIR", str(OUTPUT_DIR_DEFAULT)))
    parser.add_argument("--only-unread", action="store_true", help="Processa apenas emails nao lidos.")
    parser.add_argument("--skip-inline-attachments", action="store_true", help="Ignora anexos inline.")
    parser.add_argument("--max-messages", type=int, default=_parse_int(os.getenv("TAXAS_CAMBIAIS_MAX_MESSAGES"), MAX_MESSAGES_PER_RUN_DEFAULT))
    parser.add_argument("--page-size", type=int, default=_parse_int(os.getenv("TAXAS_CAMBIAIS_PAGE_SIZE"), 50))
    return parser


def main(argv: list[str] | None = None) -> int:
    load_project_dotenv()

    parser = build_arg_parser()
    args = parser.parse_args(argv)

    try:
        _print("=" * 84)
        _print("TAXAS CAMBIAIS - DOWNLOAD DE ANEXOS DO OFFICE 365")
        _print("=" * 84)

        cfg = load_config()
        cfg = TaxasCambiaisConfig(
            **{
                **cfg.__dict__,
                "source_folder": str(args.source_folder).strip() or cfg.source_folder,
                "backup_folder": str(args.backup_folder).strip() or cfg.backup_folder,
                "output_dir": _resolve_output_dir(str(args.output_dir)),
                "only_unread": bool(args.only_unread),
                "skip_inline_attachments": bool(args.skip_inline_attachments),
                "max_messages": int(args.max_messages),
                "page_size": max(1, int(args.page_size)),
            }
        )

        summary = run_taxas_cambiais(cfg)
        manifest_path = _write_run_manifest(cfg, summary)

        _print("")
        _print("Execucao concluida.")
        _print(f"Mailbox             : {summary.mailbox_upn}")
        _print(f"Pasta origem        : {summary.source_folder}")
        _print(f"Pasta backup        : {summary.backup_folder}")
        _print(f"Emails processados  : {summary.processed_messages}")
        _print(f"Emails movidos      : {summary.moved_messages}")
        _print(f"Anexos descarregados : {summary.downloaded_attachments}")
        _print(f"Saltados            : {summary.skipped_messages}")
        _print(f"Falhas              : {summary.failed_messages}")
        _print(f"Manifest            : {manifest_path}")
        return 0 if summary.failed_messages == 0 else 1

    except KeyboardInterrupt:
        _print("")
        _print("Execucao interrompida pelo utilizador.")
        return 130
    except Exception as exc:
        _print(f"Erro: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

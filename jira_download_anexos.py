import argparse
import os
import re
from pathlib import Path
from typing import Dict, List
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv


def require_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Variavel de ambiente obrigatoria ausente: {name}")
    return value


def normalize_base_url(value: str) -> str:
    cleaned = value.strip().rstrip("/")
    if not cleaned:
        raise RuntimeError("JIRA_DADOS_COMP_HASH esta vazio.")
    if not cleaned.startswith(("http://", "https://")):
        cleaned = f"https://{cleaned}"

    parsed = urlparse(cleaned)
    host = parsed.hostname or ""
    if not host or "." not in host:
        raise RuntimeError(
            "JIRA_DADOS_COMP_HASH invalido. Use a URL base do Jira, ex: https://empresa.atlassian.net"
        )

    return cleaned


def normalize_api_path(value: str) -> str:
    cleaned = value.strip().strip("/")
    if not cleaned:
        return "rest/api/3"
    if "/" not in cleaned:
        return "rest/api/3"
    return cleaned


def safe_filename(filename: str) -> str:
    sanitized = re.sub(r'[<>:"/\\|?*]+', "_", filename).strip()
    return sanitized or "anexo_sem_nome"


def get_issue_attachments(
    base_url: str,
    api_path: str,
    issue_key: str,
    auth: tuple[str, str],
) -> List[Dict]:
    issue_url = f"{base_url}/{api_path}/issue/{issue_key}"
    response = requests.get(
        issue_url,
        params={"fields": "attachment"},
        auth=auth,
        headers={"Accept": "application/json"},
        timeout=30,
    )

    if response.status_code == 404:
        raise RuntimeError(f"Ticket nao encontrado: {issue_key}")
    if response.status_code == 401:
        raise RuntimeError("Falha de autenticacao no Jira (401).")
    if response.status_code == 403:
        raise RuntimeError("Sem permissao para ler este ticket no Jira (403).")

    response.raise_for_status()
    payload = response.json()
    fields = payload.get("fields", {})
    attachments = fields.get("attachment", [])
    return attachments if isinstance(attachments, list) else []


def download_attachment(
    attachment: Dict,
    issue_folder: Path,
    auth: tuple[str, str],
    overwrite: bool,
) -> str:
    filename = safe_filename(str(attachment.get("filename", "anexo")))
    content_url = attachment.get("content")

    if not content_url:
        return f"[SKIP] URL de conteudo ausente para: {filename}"

    target = issue_folder / filename
    if target.exists() and not overwrite:
        return f"[SKIP] Ja existe: {target}"

    with requests.get(content_url, auth=auth, stream=True, timeout=60) as response:
        if response.status_code == 401:
            return f"[ERRO] Sem autenticacao para baixar: {filename}"
        if response.status_code == 403:
            return f"[ERRO] Sem permissao para baixar: {filename}"
        response.raise_for_status()

        with open(target, "wb") as file_obj:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file_obj.write(chunk)

    return f"[OK] Baixado: {target}"


def download_issue_attachments(
    *,
    base_url: str,
    api_path: str,
    issue_key: str,
    auth: tuple[str, str],
    output_base: Path,
    overwrite: bool,
    verbose: bool = True,
) -> Dict[str, int | str]:
    normalized_issue = issue_key.strip().upper()
    issue_folder = output_base / normalized_issue
    issue_folder.mkdir(parents=True, exist_ok=True)

    if verbose:
        print(f"Ticket: {normalized_issue}")

    attachments = get_issue_attachments(
        base_url=base_url,
        api_path=api_path,
        issue_key=normalized_issue,
        auth=auth,
    )

    downloaded = 0
    skipped = 0
    errors = 0

    if not attachments:
        if verbose:
            print("  [INFO] Sem anexos.")
        return {
            "issue_key": normalized_issue,
            "downloaded": downloaded,
            "skipped": skipped,
            "errors": errors,
        }

    for attachment in attachments:
        message = download_attachment(
            attachment=attachment,
            issue_folder=issue_folder,
            auth=auth,
            overwrite=overwrite,
        )

        if verbose:
            print(f"  {message}")

        if message.startswith("[OK]"):
            downloaded += 1
        elif message.startswith("[SKIP]"):
            skipped += 1
        else:
            errors += 1

    return {
        "issue_key": normalized_issue,
        "downloaded": downloaded,
        "skipped": skipped,
        "errors": errors,
    }


def main() -> None:
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Baixa anexos de tickets Jira para uma pasta local."
    )
    parser.add_argument(
        "issue_keys",
        nargs="+",
        help="Chave(s) do ticket Jira, ex: IZ-56680",
    )
    parser.add_argument(
        "--output",
        default=r"C:\Jira",
        help=r"Pasta base de destino (default: C:\Jira).",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Sobrescreve ficheiros ja existentes.",
    )
    args = parser.parse_args()

    jira_email = require_env("JIRA_EMAIL")
    jira_token = require_env("JIRA_TOKEN")
    jira_base = normalize_base_url(require_env("JIRA_DADOS_COMP_HASH"))
    jira_api_path = normalize_api_path(os.getenv("JIRA_DADOS_HASH", "rest/api/3"))

    output_base = Path(args.output).resolve()
    output_base.mkdir(parents=True, exist_ok=True)
    auth = (jira_email, jira_token)

    total_downloaded = 0
    total_skipped = 0
    total_errors = 0

    for issue_key in args.issue_keys:
        stats = download_issue_attachments(
            base_url=jira_base,
            api_path=jira_api_path,
            issue_key=issue_key,
            auth=auth,
            output_base=output_base,
            overwrite=args.overwrite,
            verbose=True,
        )
        total_downloaded += int(stats["downloaded"])
        total_skipped += int(stats["skipped"])
        total_errors += int(stats["errors"])

    print(
        "Concluido. "
        f"Baixados: {total_downloaded} | Ignorados: {total_skipped} | Erros: {total_errors}"
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Erro: {exc}")
        raise SystemExit(1)

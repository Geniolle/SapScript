import logging
import os
import time
from pathlib import Path

from dotenv import load_dotenv

from jira_download_anexos import (
    download_issue_attachments,
    normalize_api_path,
    normalize_base_url,
    require_env,
)

load_dotenv()

import main as sheet_main


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


def get_int_env(name: str, default: int) -> int:
    raw = os.getenv(name, str(default)).strip()
    try:
        value = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"Variavel {name} invalida: '{raw}'") from exc
    if value <= 0:
        raise RuntimeError(f"Variavel {name} deve ser maior que zero.")
    return value


def carregar_config() -> dict:
    poll_seconds = get_int_env("POLL_SECONDS", 300)
    jira_base = normalize_base_url(require_env("JIRA_DADOS_COMP_HASH"))
    jira_api_path = normalize_api_path(os.getenv("JIRA_DADOS_HASH", "rest/api/3"))
    jira_email = require_env("JIRA_EMAIL")
    jira_token = require_env("JIRA_TOKEN")
    output_dir = Path(os.getenv("JIRA_DOWNLOAD_DIR", r"C:\Jira")).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    return {
        "poll_seconds": poll_seconds,
        "jira_base": jira_base,
        "jira_api_path": jira_api_path,
        "jira_email": jira_email,
        "jira_token": jira_token,
        "output_dir": output_dir,
    }


def obter_tickets_da_sheet() -> dict[str, set[str]]:
    client = sheet_main.criar_cliente_gspread()
    dados = sheet_main.obter_dados_sheet(
        client,
        sheet_main.SPREADSHEET_ID,
        sheet_main.SHEET_NAME,
    )
    linhas = sheet_main.filtrar_linhas(
        dados,
        sheet_main.RESPONSAVEL_ALVO,
        sheet_main.SUPPLIER_ALVO,
        sheet_main.ESTADO_ALVO,
    )
    pares = sheet_main.extrair_chave_categoria(linhas)

    tickets: dict[str, set[str]] = {}
    for item in pares:
        chave = str(item.get("Chave", "")).strip().upper()
        categoria = str(item.get("IT SALSA - Categoria SAP", "")).strip()
        if not chave:
            continue
        if chave not in tickets:
            tickets[chave] = set()
        if categoria:
            tickets[chave].add(categoria)

    return tickets


def executar_ciclo(config: dict) -> None:
    tickets = obter_tickets_da_sheet()
    logging.info(
        "Sheet '%s': %s ticket(s) para validar.",
        sheet_main.SHEET_NAME,
        len(tickets),
    )

    total_downloaded = 0
    total_skipped = 0
    total_errors = 0

    for ticket in sorted(tickets):
        categorias = ", ".join(sorted(tickets[ticket])) if tickets[ticket] else "-"
        logging.info(
            "Ticket %s | Categoria(s): %s",
            ticket,
            categorias,
        )

        try:
            stats = download_issue_attachments(
                base_url=config["jira_base"],
                api_path=config["jira_api_path"],
                issue_key=ticket,
                auth=(config["jira_email"], config["jira_token"]),
                output_base=config["output_dir"],
                overwrite=False,
                verbose=False,
            )
        except Exception as exc:
            total_errors += 1
            logging.exception("Erro a processar ticket %s: %s", ticket, exc)
            continue

        downloaded = int(stats["downloaded"])
        skipped = int(stats["skipped"])
        errors = int(stats["errors"])

        total_downloaded += downloaded
        total_skipped += skipped
        total_errors += errors

        logging.info(
            "Ticket %s | Baixados=%s | Ignorados=%s | Erros=%s",
            ticket,
            downloaded,
            skipped,
            errors,
        )

    logging.info(
        "Ciclo concluido | Baixados=%s | Ignorados=%s | Erros=%s",
        total_downloaded,
        total_skipped,
        total_errors,
    )


def main() -> None:
    config = carregar_config()
    logging.info(
        "Daemon iniciado | Poll=%ss | Sheet=%s | Destino=%s",
        config["poll_seconds"],
        sheet_main.SHEET_NAME,
        config["output_dir"],
    )

    while True:
        try:
            executar_ciclo(config)
        except Exception as exc:
            logging.exception("Erro no ciclo principal: %s", exc)

        logging.info(
            "A aguardar %s segundos para o proximo ciclo.",
            config["poll_seconds"],
        )
        time.sleep(config["poll_seconds"])


if __name__ == "__main__":
    main()

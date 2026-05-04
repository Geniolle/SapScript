import os
import time
import json
import re
import logging
import unicodedata
from pathlib import Path
from typing import List, Dict, Any

import gspread
from google.oauth2.service_account import Credentials

from sap_session import ensure_sap_access_from_env, load_dotenv_manual, session_info
from workflow_engine import execute_workflows

load_dotenv_manual()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

BASE_DIR = Path(__file__).resolve().parent

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "1oYAw6cUP2-tN7Q6fBCJi2JQaahEz4kIZ4rjxWBO_cY8")
SHEET_NAME = os.getenv("SHEET_NAME", "DADOS")
RESPONSAVEL_ALVO = os.getenv("RESPONSAVEL_ALVO", "Clayton Lopes")
SUPPLIER_ALVO = os.getenv("SUPPLIER_ALVO", "Evolutive")
ESTADO_ALVO = os.getenv("ESTADO_ALVO", "In Review")
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "300"))
SAP_LOGIN_ON_MAIN = os.getenv("SAP_LOGIN_ON_MAIN", "true")
RUN_ONCE = os.getenv("RUN_ONCE", "true")

# Se existir variavel de ambiente, usa ela.
# Senao, tenta o credentials.json ao lado do main.py.
GOOGLE_CREDENTIALS_JSON = os.getenv(
    "GOOGLE_CREDENTIALS_JSON",
    str(BASE_DIR / "credentials.json")
)


def _normalize_header(text: str) -> str:
    value = str(text or "").strip()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"\s+", " ", value).strip().upper()
    return value


def criar_cliente_gspread() -> gspread.Client:
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]

    credentials_path = Path(GOOGLE_CREDENTIALS_JSON).resolve()

    if not credentials_path.exists():
        raise FileNotFoundError(
            f"Ficheiro de credenciais nao encontrado em: {credentials_path}"
        )

    credentials = Credentials.from_service_account_file(
        str(credentials_path),
        scopes=scopes
    )

    return gspread.authorize(credentials)


def obter_dados_sheet(
    client: gspread.Client,
    spreadsheet_id: str,
    sheet_name: str
) -> List[List[Any]]:
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)
    return worksheet.get_all_values()


def encontrar_indices_cabecalho(cabecalho: List[str]) -> Dict[str, int]:
    idx_por_coluna = {
        _normalize_header(nome): i
        for i, nome in enumerate(cabecalho)
    }

    idx_responsavel = idx_por_coluna.get("RESPONSAVEL", -1)
    idx_supplier = idx_por_coluna.get("SUPPLIER", -1)
    idx_estado = idx_por_coluna.get("ESTADO", -1)

    if idx_responsavel == -1:
        raise ValueError('Coluna "Responsavel" nao encontrada no cabecalho.')

    if idx_supplier == -1:
        raise ValueError('Coluna "Supplier" nao encontrada no cabecalho.')

    if idx_estado == -1:
        raise ValueError('Coluna "Estado" nao encontrada no cabecalho.')

    return {
        "RESPONSAVEL": idx_responsavel,
        "SUPPLIER": idx_supplier,
        "ESTADO": idx_estado,
    }


def filtrar_linhas(
    dados: List[List[Any]],
    responsavel_alvo: str,
    supplier_alvo: str,
    estado_alvo: str
) -> List[Dict[str, Any]]:
    if not dados:
        logging.warning("A sheet esta vazia.")
        return []

    if len(dados) < 2:
        logging.warning("A sheet possui apenas cabecalho ou nao possui linhas de dados.")
        return []

    cabecalho = dados[0]
    indices = encontrar_indices_cabecalho(cabecalho)

    idx_responsavel = indices["RESPONSAVEL"]
    idx_supplier = indices["SUPPLIER"]
    idx_estado = indices["ESTADO"]

    linhas_encontradas = []

    for i, linha in enumerate(dados[1:], start=2):
        valor_responsavel = str(linha[idx_responsavel]).strip() if idx_responsavel < len(linha) else ""
        valor_supplier = str(linha[idx_supplier]).strip() if idx_supplier < len(linha) else ""
        valor_estado = str(linha[idx_estado]).strip() if idx_estado < len(linha) else ""

        if (
            valor_responsavel == responsavel_alvo
            and valor_supplier == supplier_alvo
            and valor_estado == estado_alvo
        ):
            linha_dict = {}

            for j, nome_coluna in enumerate(cabecalho):
                nome_coluna_limpo = str(nome_coluna).strip()
                linha_dict[nome_coluna_limpo] = linha[j] if j < len(linha) else ""

            linhas_encontradas.append({
                "numero_linha": i,
                "dados": linha_dict
            })

    return linhas_encontradas


def extrair_chave_categoria(
    linhas: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    resultado = []

    for item in linhas:
        dados_linha = item.get("dados", {})
        resultado.append({
            "numero_linha": item.get("numero_linha"),
            "Chave": str(dados_linha.get("Chave", "")).strip(),
            "IT SALSA - Categoria SAP": str(
                dados_linha.get("IT SALSA - Categoria SAP", "")
            ).strip(),
        })

    return resultado


def _to_bool(value: str) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on", "sim", "s"}


def garantir_sessao_sap() -> None:
    if not _to_bool(SAP_LOGIN_ON_MAIN):
        logging.info("SAP_LOGIN_ON_MAIN desativado.")
        return

    session = ensure_sap_access_from_env(
        key=os.getenv("WORKFLOW_SAP_KEY", "S4DCLNT100"),
        timeout_s=40,
        load_env=True,
    )
    info = session_info(session)
    logging.info(
        "Sessao SAP pronta | Sistema=%s | Cliente=%s | User=%s",
        info["system_name"],
        info["client"],
        info["user"],
    )


def executar_rotina() -> None:
    logging.info("A iniciar rotina.")
    logging.info(
        "Filtros aplicados | Responsavel='%s' | Supplier='%s' | Estado='%s' | Sheet='%s'",
        RESPONSAVEL_ALVO,
        SUPPLIER_ALVO,
        ESTADO_ALVO,
        SHEET_NAME
    )
    logging.info("Credenciais em uso: %s", Path(GOOGLE_CREDENTIALS_JSON).resolve())

    client = criar_cliente_gspread()
    dados = obter_dados_sheet(client, SPREADSHEET_ID, SHEET_NAME)
    linhas = filtrar_linhas(dados, RESPONSAVEL_ALVO, SUPPLIER_ALVO, ESTADO_ALVO)
    chaves_categorias = extrair_chave_categoria(linhas)

    logging.info("Total de linhas encontradas: %s", len(linhas))
    logging.info(
        "Total de pares Chave + IT SALSA - Categoria SAP: %s",
        len(chaves_categorias)
    )

    for item in chaves_categorias:
        logging.info(
            "Linha %s | Chave='%s' | IT SALSA - Categoria SAP='%s'",
            item["numero_linha"],
            item["Chave"],
            item["IT SALSA - Categoria SAP"]
        )

    if linhas:
        garantir_sessao_sap()
    else:
        logging.info("Sem tickets elegiveis nesta execucao; login SAP nao necessario.")

    try:
        execute_workflows(linhas, base_dir=BASE_DIR)
    except Exception as workflow_error:
        logging.exception("Erro ao executar workflows: %s", workflow_error)

    for item in linhas:
        logging.info(
            "Linha %s | %s",
            item["numero_linha"],
            json.dumps(item["dados"], ensure_ascii=False)
        )


def main() -> None:
    if _to_bool(RUN_ONCE):
        try:
            executar_rotina()
        except Exception as e:
            logging.exception("Erro durante a execucao da rotina: %s", e)
        return

    while True:
        try:
            executar_rotina()
        except Exception as e:
            logging.exception("Erro durante a execucao da rotina: %s", e)

        logging.info("A aguardar %s segundos para a proxima execucao.", POLL_SECONDS)
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()

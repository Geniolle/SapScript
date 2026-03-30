import os
import re

# Caminhos
BASE_DIR = r"C:\SAP Script"
PROCESSOS_DIR = os.path.join(BASE_DIR, "Processos")
SAPLOGON_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"

# Texto exigido (sem ícone). O ícone é aplicado no logger.
MSG_RZ11_SCRIPTING = 'Ativar na transação RZ11 o nome do parametro "sapgui/user_scripting" alterar para "TRUE"'

# Módulo de lista/pesquisa de requests
PESQUISAR_REQUEST_PATH = os.path.join(PROCESSOS_DIR, "pesquisar_request.py")

# Ambientes SAP
AMBIENTES = {
    "1": ("DEV", "DESENVOLVIMENTO (S4H)"),
    "2": ("QAD", "QUALIDADE (S4H)"),
    "3": ("PRD", "PRODUÇÃO (S4H)"),
    "4": ("CUA", "CUA (PRD)")
}

MAPA_SISTEMA = {
    "DEV": "S4D",
    "QAD": "S4Q",
    "PRD": "S4P",
    "CUA": "SPA"
}

CLIENTES_POR_AMBIENTE = {
    "DEV": "100",
    "QAD": "100",
    "PRD": "100",
    "CUA": "001"
}

# Sleeps (ajusta se o SAP “se perder” nos inputs)
SLEEP_UI = 0.25
SLEEP_ACTION = 0.40

# Scan para extrair TRKORR do ecrã (sem status bar)
SCAN_MAX_DEPTH = 6
SCAN_MAX_NODES = 2500

REQ_PATTERNS = [
    re.compile(r"\b[A-Z0-9]{3,4}K\d{6,}\b"),      # ex.: S4QK900416 / S4DK951499
    re.compile(r"\b[A-Z0-9]{3,4}[A-Z]\d{6,}\b"),  # fallback
]
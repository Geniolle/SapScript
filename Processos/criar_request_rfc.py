# -*- coding: utf-8 -*-
"""
criar_request_rfc.py

Objetivo:
- Criar Transport Requests (Customizing ou Workbench) no SAP 100% via RFC
- Utiliza a combinação ABAP standard TR_EXT_CREATE_REQUEST (Cabeçalho) + TR40_TASK_ADD (Tarefa do Utilizador)
- Oferece alto desempenho (criação instantânea em < 1s) sem abrir janelas do SAP GUI
- Mantém interface compatível com o script legado criar_request.py (GUI SE10)
"""

import sys
import time
import os
import functools
from typing import Optional, Any, Tuple

if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

print = functools.partial(print, flush=True)

try:
    from pyrfc import Connection, ABAPApplicationError
    HAS_PYRFC = True
except Exception as exc_import:
    HAS_PYRFC = False
    _PYRFC_IMPORT_ERROR = exc_import

try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.getcwd(), ".env"))
except Exception:
    pass


# ─────────────────────────────────────────────────────────────────────────────
# 1. HELPER DE CONEXÃO RFC COM SUPORTE A ALIASES
# ─────────────────────────────────────────────────────────────────────────────

MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P", "CUA": "SPA"}
TARGET_SISTEMA = {"DEV": "S4Q", "S4D": "S4Q", "QAD": "S4P", "S4Q": "S4P"}


def obter_conexao_rfc(ambiente: Optional[str] = None) -> tuple[Any, str, str]:
    """
    Obtém uma conexão pyrfc.Connection com base no ficheiro .env.
    Retorna: (conn, system_code, user)
    """
    if not HAS_PYRFC:
        raise RuntimeError(f"A biblioteca 'pyrfc' não está disponível: {_PYRFC_IMPORT_ERROR}")

    system_up = (ambiente or os.getenv("SAP_SYSTEM") or "DEV").upper().strip()

    ALIAS_MAP = {
        "DEV": ["DEV", "S4D", "S4DCLNT100"],
        "S4D": ["DEV", "S4D", "S4DCLNT100"],
        "S4DCLNT100": ["DEV", "S4D", "S4DCLNT100"],
        "QAD": ["QAD", "S4Q", "S4QCLNT100"],
        "S4Q": ["QAD", "S4Q", "S4QCLNT100"],
        "S4QCLNT100": ["QAD", "S4Q", "S4QCLNT100"],
        "PRD": ["PRD", "S4P", "S4PCLNT100"],
        "S4P": ["PRD", "S4P", "S4PCLNT100"],
        "S4PCLNT100": ["PRD", "S4P", "S4PCLNT100"],
    }
    keys_to_try = ALIAS_MAP.get(system_up, [system_up])
    system_code = MAPA_SISTEMA.get(system_up, system_up)

    ashost = ""
    for k in keys_to_try:
        ashost = os.getenv(f"SAP_ASHOST_{k}", "").strip()
        if ashost:
            break
    if not ashost:
        ashost = os.getenv("SAP_ASHOST", "").strip()

    sysnr = ""
    for k in keys_to_try:
        sysnr = os.getenv(f"SAP_SYSNR_{k}", "").strip()
        if sysnr:
            break
    if not sysnr:
        sysnr = os.getenv("SAP_SYSNR", "00").strip() or "00"

    client = ""
    for k in keys_to_try:
        client = os.getenv(f"SAP_CLIENT_{k}", "").strip()
        if client:
            break
    if not client:
        client = os.getenv("SAP_CLIENT", "100").strip() or "100"

    user = ""
    for k in keys_to_try:
        user = os.getenv(f"SAP_USER_{k}", "").strip()
        if user:
            break
    if not user:
        user = os.getenv("SAP_USER", "").strip()

    lang = ""
    for k in keys_to_try:
        lang = os.getenv(f"SAP_LANGUAGE_{k}", "").strip()
        if lang:
            break
    if not lang:
        lang = os.getenv("SAP_LANGUAGE", "PT").strip() or "PT"

    passwd = ""
    for k in keys_to_try:
        passwd = (
            os.getenv(f"SAP_PASSWORD_{k}")
            or os.getenv(f"SAP_PASSWORD_{k}CLNT{client}")
            or ""
        ).strip()
        if passwd:
            break
    if not passwd:
        passwd = (
            os.getenv("SAP_PASSWD")
            or os.getenv("SAP_PASSWORD")
            or ""
        ).strip()

    if not ashost or not user or not passwd:
        raise ValueError(
            f"Faltam credenciais SAP no ficheiro .env para o sistema '{system_up}'. "
            f"Verifique SAP_ASHOST_{system_up}, SAP_USER_{system_up} e SAP_PASSWORD_{system_up}."
        )

    conn = Connection(ashost=ashost, sysnr=sysnr, client=client, user=user, passwd=passwd, lang=lang)
    return conn, system_code, user


# ─────────────────────────────────────────────────────────────────────────────
# 2. FUNÇÃO PRINCIPAL DE CRIAÇÃO VIA RFC (TR_EXT_CREATE_REQUEST + TR40_TASK_ADD)
# ─────────────────────────────────────────────────────────────────────────────

def criar_nova_request_rfc(
    ambiente: str = "DEV",
    tipo: str = "customizing",
    descricao: str = "",
    target_system: Optional[str] = None,
) -> Tuple[str, str]:
    """
    Cria uma nova Transport Request completa (Cabeçalho + Tarefa de Utilizador) no SAP via RFC.
    
    Parâmetros:
    - ambiente: 'DEV', 'QAD', 'S4D', etc.
    - tipo: 'customizing' (ou '1') para Customizing; 'workbench' (ou '2') para Workbench.
    - descricao: Texto curto da request (máx. 60 caracteres).
    - target_system: Sistema de destino (ex: 'S4Q'). Se None, deriva automaticamente do ambiente.
    
    Retorna:
    - (request_number, task_number) ex: ("S4DK953543", "S4DK953544")
    """
    conn, system_code, user = obter_conexao_rfc(ambiente)

    # Determinar tipo ABAP: 'W' = Customizing, 'K' = Workbench
    tipo_str = str(tipo or "").strip().lower()
    if tipo_str in ("2", "k", "workbench"):
        abap_req_type = "K"
        tipo_lbl = "Workbench"
        category_wbo = "K"
    else:
        abap_req_type = "W"
        tipo_lbl = "Customizing"
        category_wbo = "C"

    # Tratar descrição
    desc_limpa = (descricao or "").strip()[:60]
    if not desc_limpa:
        desc_limpa = f"REQUEST {tipo_lbl.upper()} VIA RFC ({user})"

    # Target system
    if not target_system:
        system_up = (ambiente or "DEV").upper().strip()
        target_system = TARGET_SISTEMA.get(system_up, "S4Q")

    print(f"🚀 A criar Request {tipo_lbl} no SAP via RFC ({system_code})...")
    print(f"   - Descrição: '{desc_limpa}'")
    print(f"   - Alvo Transporte: '{target_system}'")
    print(f"   - Autor: '{user}'")

    req_number = ""
    task_number = ""

    # Passo 1: Criar Cabeçalho da Request via TR_EXT_CREATE_REQUEST
    try:
        res = conn.call(
            "TR_EXT_CREATE_REQUEST",
            IV_AUTHOR=user,
            IV_REQUEST_TYPE=abap_req_type,
            IV_TEXT=desc_limpa,
            IV_TARGET=target_system
        )
        req_number = str(res.get("ES_REQ_ID") or "").strip().upper()
        if not req_number and isinstance(res.get("ES_REQ_HEADER"), dict):
            req_number = str(res["ES_REQ_HEADER"].get("REQ_ID") or "").strip().upper()
    except Exception as exc1:
        print(f"⚠️ Método TR_EXT_CREATE_REQUEST falhou: {exc1}. A tentar CTS_WBO_CREATE_REQUEST...")

    # Fallback Método 1b: CTS_WBO_CREATE_REQUEST
    if not req_number:
        try:
            res_wbo = conn.call(
                "CTS_WBO_CREATE_REQUEST",
                CATEGORY=category_wbo,
                DESCRIPTION=desc_limpa,
                OWNER=user,
                SID=system_code
            )
            req_number = str(res_wbo.get("REQUEST") or "").strip().upper()
        except Exception as exc2:
            print(f"⚠️ Método CTS_WBO_CREATE_REQUEST falhou: {exc2}...")

    if not req_number:
        raise RuntimeError(f"Não foi possível obter o número do cabeçalho da request no sistema {system_code}.")

    # Passo 2: Criar Tarefa do Utilizador sob a Request via TR40_TASK_ADD
    try:
        res_task = conn.call(
            "TR40_TASK_ADD",
            IV_TRKORR=req_number,
            TT_USERLIST=[{"AS4USER": user}]
        )
        user_list = res_task.get("TT_USERLIST") or []
        if user_list and isinstance(user_list, list):
            row = user_list[0]
            task_number = str(row.get("CORRECTION") or row.get("REPAIR") or "").strip().upper()
    except Exception as exc_t:
        print(f"⚠️ Aviso: Não foi possível gerar tarefa automática sob a request {req_number}: {exc_t}")

    print(f"✅ Request e Tarefa criadas com sucesso via RFC: Request={req_number} | Tarefa={task_number or '(sem tarefa)'}")
    print(f"REQUEST_NUMBER={req_number}")

    # Atualizar variáveis de ambiente globais
    os.environ["SAP_ULTIMA_REQUEST"] = req_number
    if task_number:
        os.environ["SAP_ULTIMA_TAREFA"] = task_number

    return req_number, task_number


# ─────────────────────────────────────────────────────────────────────────────
# 3. INTERFACE DE COMPATIBILIDADE COM O EXECUTAR DO COCKPIT
# ─────────────────────────────────────────────────────────────────────────────

def executar(
    ambiente_cockpit=None,
    request_ctx=None,
    request_transporte=None,
    modo_nao_interativo=False,
    tipo_ordem="customizing",
    descricao_request="",
    system_name="",
    client="",
    chamado_pelo_main=False,
    **kwargs
) -> str:
    """
    Função wrapper compatível com o contrato de execução do SAP Cockpit.
    """
    req_recebida = ""
    if isinstance(request_ctx, dict):
        req_recebida = str(request_ctx.get("request_number", "")).strip().upper()
    if not req_recebida:
        req_recebida = str(request_transporte or "").strip().upper()

    if req_recebida:
        print(f"REQUEST_NUMBER={req_recebida}")
        return req_recebida

    ambiente = ambiente_cockpit or system_name or "DEV"
    req_number, _ = criar_nova_request_rfc(
        ambiente=ambiente,
        tipo=tipo_ordem,
        descricao=descricao_request
    )
    return req_number


# ─────────────────────────────────────────────────────────────────────────────
# 4. EXECUÇÃO VIA TERMINAL CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Criar Transport Request no SAP via RFC.")
    parser.add_argument("--ambiente", type=str, default="DEV", help="Ambiente SAP (DEV, QAD, S4D)")
    parser.add_argument("--tipo", type=str, default="customizing", help="Tipo de Ordem (customizing ou workbench)")
    parser.add_argument("--desc", type=str, default="", help="Descrição / Texto curto da Request")
    parser.add_argument("--target", type=str, default=None, help="Sistema alvo de transporte (ex: S4Q)")

    args = parser.parse_args()

    try:
        req, task = criar_nova_request_rfc(
            ambiente=args.ambiente,
            tipo=args.tipo,
            descricao=args.desc,
            target_system=args.target
        )
        print(f"\n🎉 Sucesso! Request: {req} | Tarefa: {task}")
    except Exception as exc:
        print(f"\n❌ Erro na criação da Request via RFC: {exc}")
        sys.exit(1)

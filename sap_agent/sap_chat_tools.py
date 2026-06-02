"""sap_chat_tools.py – Ferramentas de consulta SAP para o chat interativo.

Este módulo detecta a intenção de consulta SAP numa mensagem de utilizador,
conecta ao SAP via RFC (usando as credenciais do .env) e devolve os dados
reais como texto estruturado para ser injetado no prompt do Gemini.

Se pyrfc não estiver disponível (ex: no container Docker sem SAP RFC SDK),
devolve um bloco descritivo de "o que consultar" para que o Gemini possa
orientar o utilizador manualmente.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Any

# ──────────────────────────────────────────────
# Padrões de deteção de intenção
# ──────────────────────────────────────────────

# Números que parecem ordens internas CO (8 dígitos, ex: 6000066481 → 10 dígitos também)
_RE_ORDER_NUMBER = re.compile(
    r"\b(\d{7,12})\b"
)

# Números que parecem Purchase Orders (4500000000 - 10 dígitos começando com 45)
_RE_PO_NUMBER = re.compile(r"\b(45\d{8})\b")

# Números que parecem WBS (ex: E-2024-0001 ou PROJECT.1.1)
_RE_WBS = re.compile(r"\b([A-Z]{1,6}[-./]\d{1,6}[-./]\d{1,6}[-./]?\d{0,4})\b", re.IGNORECASE)

# Documentos FI (ex: 1000000001 ou 2000000001 com 10 dígitos)
_RE_FI_DOC = re.compile(r"\b([12]\d{9})\b")

# Palavras de intenção SAP
_RE_INTENT = re.compile(
    r"\b(entr[ae]|acede|aceda|abr[ae]|analisa|analise|verifica|verifique|consulta|consulte|"
    r"mostra|mostre|vai|vá|abre|veja|ver|look|check|open|show)\b.{0,60}"
    r"\b(pedido|ordem|order|po|documento|doc|asset|imobilizado|wbs|projeto|project)\b",
    re.IGNORECASE,
)

_RE_DIRECT_INTENT = re.compile(
    r"\b(pedido|ordem|order)\b.{0,20}\b(\d{7,12})\b",
    re.IGNORECASE,
)


@dataclass
class SapQueryResult:
    """Resultado de uma consulta ao SAP para injeção no contexto do Gemini."""

    object_type: str  # 'internal_order', 'po', 'fi_doc', 'wbs', 'asset', 'unknown'
    object_number: str
    data_blocks: list[str]  # Blocos de texto formatado para incluir no prompt
    error: str | None = None
    is_real_data: bool = False  # True se os dados vieram mesmo do SAP


def detect_sap_intent(message: str) -> tuple[str | None, str | None]:
    """Tenta detetar a intenção de consulta SAP na mensagem.

    Returns:
        (object_type, object_number) ou (None, None) se não detetado.
    """
    # Deteção direta: "pedido 6000066481" ou "ordem 6000066481"
    m = _RE_DIRECT_INTENT.search(message)
    if m:
        keyword = m.group(1).lower()
        number = m.group(2)
        if "po" in keyword:
            return ("po", number)
        # PO começa com 45
        if number.startswith("45") and len(number) == 10:
            return ("po", number)
        return ("internal_order", number)

    # Deteção por intenção + número
    has_intent = bool(_RE_INTENT.search(message))
    if not has_intent:
        return (None, None)

    # PO
    m = _RE_PO_NUMBER.search(message)
    if m:
        return ("po", m.group(1))

    # Ordem interna (8-12 dígitos, não começa com 45)
    m = _RE_ORDER_NUMBER.search(message)
    if m:
        number = m.group(1)
        if not number.startswith("45"):
            return ("internal_order", number)
        return ("po", number)

    return (None, None)


def query_sap_object(
    object_type: str,
    object_number: str,
    company_code: str | None = None,
) -> SapQueryResult:
    """Conecta ao SAP e consulta os dados do objeto indicado.

    Usa pyrfc via SapRfcClient. Se pyrfc não estiver disponível,
    retorna um bloco de orientação manual.
    """
    try:
        from sap_agent.config import SapConnectionConfig
        from sap_agent.safety import SafetyGuard
        from sap_agent.sap_rfc_client import SapRfcClient, SapRfcUnavailable
    except ImportError:
        from .config import SapConnectionConfig
        from .safety import SafetyGuard
        from .sap_rfc_client import SapRfcClient, SapRfcUnavailable

    # Carregar configuração RFC do .env
    try:
        cfg = SapConnectionConfig.from_env()
    except RuntimeError as exc:
        return SapQueryResult(
            object_type=object_type,
            object_number=object_number,
            data_blocks=[_build_manual_guidance(object_type, object_number, str(exc))],
            error=str(exc),
            is_real_data=False,
        )

    # SafetyGuard permissivo para leitura de tabelas do chat (sem whitelist fixa)
    guard = SafetyGuard.build(
        allow_write_operations=False,
        allowed_functions=[],  # sem whitelist → qualquer RFC read é permitida
        allowed_tables=[],     # sem whitelist → qualquer tabela é permitida
    )
    client = SapRfcClient(config=cfg, safety_guard=guard)

    # Testar conexão antes de prosseguir — deteta pyrfc indisponível rapidamente
    try:
        client.ping()
    except SapRfcUnavailable as exc:
        # pyrfc não está instalado no container — devolver orientação manual
        return SapQueryResult(
            object_type=object_type,
            object_number=object_number,
            data_blocks=[_build_manual_guidance(object_type, object_number, str(exc))],
            error=str(exc),
            is_real_data=False,
        )
    except Exception as exc:
        # Erro de ligação (host inacessível, credenciais, etc.)
        error_msg = str(exc)
        return SapQueryResult(
            object_type=object_type,
            object_number=object_number,
            data_blocks=[
                f"⚠️ Não foi possível conectar ao SAP RFC: {error_msg}\n\n"
                + _build_manual_guidance(object_type, object_number, "")
            ],
            error=error_msg,
            is_real_data=False,
        )

    try:
        if object_type == "internal_order":
            return _query_internal_order(client, object_number)
        elif object_type == "po":
            return _query_purchase_order(client, object_number)
        elif object_type == "fi_doc":
            return _query_fi_document(client, object_number, company_code or "")
        elif object_type == "asset":
            return _query_asset(client, object_number, company_code or "")
        else:
            return _query_generic(client, object_number)
    except Exception as exc:
        # Fallback: orientação manual
        return SapQueryResult(
            object_type=object_type,
            object_number=object_number,
            data_blocks=[_build_manual_guidance(object_type, object_number, str(exc))],
            error=str(exc),
            is_real_data=False,
        )


# ──────────────────────────────────────────────
# Consultas específicas por tipo de objeto
# ──────────────────────────────────────────────

def _query_internal_order(client: Any, order_number: str) -> SapQueryResult:
    """Consulta AUFK (mestre de ordens) e AUAK (regras de liquidação)."""
    blocks: list[str] = []

    # AUFK – Cabeçalho da ordem interna
    try:
        rows = client.read_table(
            "AUFK",
            fields=["AUFNR", "AUART", "BUKRS", "KOSTL", "KTEXT", "ERNAM", "ERDAT", "STAT", "OBJNR"],
            options=[f"AUFNR = '{order_number.zfill(12)}'"],
            rowcount=5,
        )
        if rows:
            r = rows[0]
            blocks.append(
                f"**AUFK – Mestre da Ordem Interna {order_number}:**\n"
                f"- Tipo de Ordem (AUART): {r.get('AUART', '-')}\n"
                f"- Empresa (BUKRS): {r.get('BUKRS', '-')}\n"
                f"- Centro de Custo (KOSTL): {r.get('KOSTL', '-')}\n"
                f"- Descrição (KTEXT): {r.get('KTEXT', '-')}\n"
                f"- Criado por: {r.get('ERNAM', '-')} em {r.get('ERDAT', '-')}\n"
                f"- Estado (STAT): {r.get('STAT', '-')}\n"
                f"- Nr. Objeto (OBJNR): {r.get('OBJNR', '-')}"
            )
        else:
            blocks.append(f"**AUFK:** Ordem {order_number} não encontrada na AUFK.")
    except Exception as exc:
        blocks.append(f"**AUFK:** Erro ao ler AUFK: {exc}")

    # AUAK – Regras de Liquidação (Settlements)
    try:
        rows = client.read_table(
            "AUAK",
            fields=["AUFNR", "LFDNR", "KSCHL", "PROZS", "BEWAR", "KOSTL", "ANLNR", "ANLN2", "PSPNR"],
            options=[f"AUFNR = '{order_number.zfill(12)}'"],
            rowcount=20,
        )
        if rows:
            lines = [f"**AUAK – Regras de Liquidação (Ordem {order_number}):**"]
            for r in rows:
                lines.append(
                    f"  - Linha {r.get('LFDNR', '-')}: Tipo {r.get('KSCHL', '-')}, "
                    f"% {r.get('PROZS', '-')}, Tipo Custo {r.get('BEWAR', '-')}, "
                    f"C.Custo {r.get('KOSTL', '-')}, Imob {r.get('ANLNR', '-')}/{r.get('ANLN2', '-')}, "
                    f"WBS {r.get('PSPNR', '-')}"
                )
            blocks.append("\n".join(lines))
        else:
            blocks.append(f"**AUAK:** Nenhuma regra de liquidação encontrada para a ordem {order_number}.")
    except Exception as exc:
        blocks.append(f"**AUAK:** Erro ao ler AUAK: {exc}")

    # AUFP – Posições (Orçamentos/Custos planeados)
    try:
        rows = client.read_table(
            "COSP",
            fields=["OBJNR", "GJAHR", "VERSN", "WRTTP", "KSTAR", "BEKNZ", "WTG001"],
            options=[f"OBJNR = 'OR{order_number.zfill(12)}'", "AND GJAHR = '2024'"],
            rowcount=10,
        )
        if rows:
            lines = [f"**COSP – Custos CO (Ordem {order_number}, 2024):**"]
            for r in rows:
                lines.append(
                    f"  - Classe custo {r.get('KSTAR', '-')}: {r.get('WTG001', '-')} "
                    f"[versão {r.get('VERSN', '-')}, tipo {r.get('WRTTP', '-')}]"
                )
            blocks.append("\n".join(lines))
    except Exception as exc:
        blocks.append(f"**COSP:** Aviso ao ler custos: {exc}")

    return SapQueryResult(
        object_type="internal_order",
        object_number=order_number,
        data_blocks=blocks,
        is_real_data=True,
    )


def _query_purchase_order(client: Any, po_number: str) -> SapQueryResult:
    """Consulta EKKO (cabeçalho PO) e EKPO (posições PO)."""
    blocks: list[str] = []

    try:
        rows = client.read_table(
            "EKKO",
            fields=["EBELN", "BUKRS", "BSART", "LIFNR", "BEDAT", "EKORG", "EKGRP", "WAERS", "PROCSTAT"],
            options=[f"EBELN = '{po_number.zfill(10)}'"],
            rowcount=1,
        )
        if rows:
            r = rows[0]
            blocks.append(
                f"**EKKO – Cabeçalho PO {po_number}:**\n"
                f"- Empresa: {r.get('BUKRS', '-')}\n"
                f"- Tipo doc: {r.get('BSART', '-')}\n"
                f"- Fornecedor: {r.get('LIFNR', '-')}\n"
                f"- Data: {r.get('BEDAT', '-')}\n"
                f"- Org. Compras: {r.get('EKORG', '-')}\n"
                f"- Moeda: {r.get('WAERS', '-')}\n"
                f"- Estado: {r.get('PROCSTAT', '-')}"
            )
        else:
            blocks.append(f"**EKKO:** PO {po_number} não encontrada.")
    except Exception as exc:
        blocks.append(f"**EKKO:** Erro: {exc}")

    try:
        rows = client.read_table(
            "EKPO",
            fields=["EBELN", "EBELP", "MATNR", "TXZ01", "MENGE", "MEINS", "NETPR", "WAERS", "ELIKZ"],
            options=[f"EBELN = '{po_number.zfill(10)}'"],
            rowcount=20,
        )
        if rows:
            lines = [f"**EKPO – Posições PO {po_number}:**"]
            for r in rows:
                lines.append(
                    f"  - Pos {r.get('EBELP', '-')}: Mat {r.get('MATNR', '-')} "
                    f"'{r.get('TXZ01', '-')}' Qtd {r.get('MENGE', '-')} {r.get('MEINS', '-')} "
                    f"@ {r.get('NETPR', '-')} {r.get('WAERS', '-')} [Ent:{r.get('ELIKZ', '-')}]"
                )
            blocks.append("\n".join(lines))
    except Exception as exc:
        blocks.append(f"**EKPO:** Erro: {exc}")

    return SapQueryResult(
        object_type="po",
        object_number=po_number,
        data_blocks=blocks,
        is_real_data=True,
    )


def _query_fi_document(client: Any, doc_number: str, company_code: str) -> SapQueryResult:
    """Consulta BKPF (cabeçalho doc FI)."""
    blocks: list[str] = []
    bukrs = company_code or os.getenv("SAP_BUKRS", "1000")
    try:
        rows = client.read_table(
            "BKPF",
            fields=["BUKRS", "BELNR", "GJAHR", "BLART", "BUDAT", "CPUDT", "USNAM", "TCODE", "XBLNR", "BKTXT"],
            options=[
                f"BUKRS = '{bukrs}'",
                f"AND BELNR = '{doc_number.zfill(10)}'",
            ],
            rowcount=5,
        )
        if rows:
            r = rows[0]
            blocks.append(
                f"**BKPF – Documento FI {doc_number}:**\n"
                f"- Empresa: {r.get('BUKRS', '-')}, Exercício: {r.get('GJAHR', '-')}\n"
                f"- Tipo: {r.get('BLART', '-')}, Data contab: {r.get('BUDAT', '-')}\n"
                f"- Transação: {r.get('TCODE', '-')}, Utilizador: {r.get('USNAM', '-')}\n"
                f"- Ref externa: {r.get('XBLNR', '-')}, Texto: {r.get('BKTXT', '-')}"
            )
        else:
            blocks.append(f"**BKPF:** Documento {doc_number} não encontrado para empresa {bukrs}.")
    except Exception as exc:
        blocks.append(f"**BKPF:** Erro: {exc}")
    return SapQueryResult(
        object_type="fi_doc",
        object_number=doc_number,
        data_blocks=blocks,
        is_real_data=True,
    )


def _query_asset(client: Any, asset_number: str, company_code: str) -> SapQueryResult:
    """Consulta ANLA (mestre de imobilizados)."""
    blocks: list[str] = []
    bukrs = company_code or os.getenv("SAP_BUKRS", "1000")
    try:
        rows = client.read_table(
            "ANLA",
            fields=["BUKRS", "ANLN1", "ANLN2", "AKTIV", "DEAKT", "TXT50", "ANLKL", "KOSTL", "AUFNR"],
            options=[
                f"BUKRS = '{bukrs}'",
                f"AND ANLN1 = '{asset_number.zfill(12)}'",
            ],
            rowcount=5,
        )
        if rows:
            r = rows[0]
            blocks.append(
                f"**ANLA – Imobilizado {asset_number}:**\n"
                f"- Empresa: {r.get('BUKRS', '-')}, Sub-nr: {r.get('ANLN2', '-')}\n"
                f"- Descrição: {r.get('TXT50', '-')}\n"
                f"- Classe: {r.get('ANLKL', '-')}, Centro custo: {r.get('KOSTL', '-')}\n"
                f"- Ordem orig: {r.get('AUFNR', '-')}\n"
                f"- Data ativ: {r.get('AKTIV', '-')}, Data desativ: {r.get('DEAKT', '-')}"
            )
        else:
            blocks.append(f"**ANLA:** Imobilizado {asset_number} não encontrado para empresa {bukrs}.")
    except Exception as exc:
        blocks.append(f"**ANLA:** Erro: {exc}")
    return SapQueryResult(
        object_type="asset",
        object_number=asset_number,
        data_blocks=blocks,
        is_real_data=True,
    )


def _query_generic(client: Any, number: str) -> SapQueryResult:
    """Tenta detetar o tipo de objeto e consultar."""
    blocks = [
        f"Número {number} detetado. A tentar identificar tipo de objeto...",
        _build_manual_guidance("unknown", number, "Tipo de objeto não reconhecido automaticamente."),
    ]
    return SapQueryResult(
        object_type="unknown",
        object_number=number,
        data_blocks=blocks,
        is_real_data=False,
    )


def _build_manual_guidance(object_type: str, number: str, error: str) -> str:
    """Constrói um bloco de orientação manual quando o SAP não está acessível."""
    guidance_map = {
        "internal_order": (
            f"Para analisar a Ordem Interna {number} manualmente:\n"
            f"- SE16N → AUFK com AUFNR = {number.zfill(12)} (mestre da ordem)\n"
            f"- SE16N → AUAK com AUFNR = {number.zfill(12)} (regras de liquidação)\n"
            f"- KO03 → Visualizar Ordem Interna {number}\n"
            f"- KO8G → Liquidação coletiva de ordens\n"
            f"- CJ8G → Apropriação real (se for ordem de projeto)"
        ),
        "po": (
            f"Para analisar a Purchase Order {number} manualmente:\n"
            f"- ME23N → Visualizar PO {number}\n"
            f"- SE16N → EKKO com EBELN = {number.zfill(10)}\n"
            f"- SE16N → EKPO com EBELN = {number.zfill(10)}"
        ),
        "fi_doc": (
            f"Para analisar o Documento FI {number} manualmente:\n"
            f"- FB03 → Visualizar documento {number}\n"
            f"- SE16N → BKPF com BELNR = {number.zfill(10)}"
        ),
        "asset": (
            f"Para analisar o Imobilizado {number} manualmente:\n"
            f"- AS03 → Visualizar imobilizado {number}\n"
            f"- SE16N → ANLA com ANLN1 = {number.zfill(12)}"
        ),
    }
    base = guidance_map.get(object_type, f"Objeto {number}: verificar tipo na SE16N / SE11.")
    note = f"\n_(Nota: acesso direto via RFC não disponível: {error})_" if error else ""
    return base + note

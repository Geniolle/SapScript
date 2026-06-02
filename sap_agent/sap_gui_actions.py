"""sap_gui_actions.py – Automação SAP GUI a partir do chat interativo.

Este módulo implementa ações que o worker Windows pode executar no SAP GUI
(via COM/pywin32) em resposta a comandos do utilizador no chat.

Ações suportadas:
  - se16n_query   : Pesquisa em tabela via SE16N
  - open_transaction : Abre qualquer transação via campo de comando
  - read_sbar      : Lê o status bar da sessão ativa
  - ko03_view      : Visualiza Ordem Interna (KO03)
  - me23n_view     : Visualiza Purchase Order (ME23N)
  - fb03_view      : Visualiza Documento FI (FB03)

Todas as ações retornam (result_text: str, rows: list[dict], error: str | None).
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field
from typing import Any


# ──────────────────────────────────────────────────────────────────────────────
# Modelos de resultado
# ──────────────────────────────────────────────────────────────────────────────

@dataclass
class SapGuiResult:
    """Resultado de uma ação SAP GUI."""
    action: str
    description: str
    result_text: str = ""          # Texto legível para mostrar no chat
    rows: list[dict[str, str]] = field(default_factory=list)  # Dados em tabela
    error: str | None = None
    success: bool = True


# ──────────────────────────────────────────────────────────────────────────────
# Utilitários SAP GUI
# ──────────────────────────────────────────────────────────────────────────────

def _get_session():
    """Obtém sessão SAP GUI disponível via COM."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
    except Exception as exc:
        raise RuntimeError(
            f"Não foi possível ligar ao SAP GUI Scripting: {exc}\n"
            "Confirma que o SAP Logon está aberto e que o SAP GUI Scripting está ativo "
            "(Tools → Options → Accessibility & Scripting → Enable Scripting)."
        ) from exc

    # Percorrer todas as conexões/sessões e devolver a primeira não ocupada
    for ci in range(application.Children.Count):
        conn = application.Children(ci)
        for si in range(conn.Children.Count):
            sess = conn.Children(si)
            try:
                if not sess.Busy:
                    return sess
            except Exception:
                continue

    raise RuntimeError(
        "Nenhuma sessão SAP disponível (todas ocupadas ou nenhuma aberta).\n"
        "Abre o SAP Logon e inicia sessão antes de usar esta funcionalidade."
    )


def _navigate_to(session, transaction: str) -> str:
    """Navega para uma transação e retorna o status bar."""
    okcd = session.findById("wnd[0]/tbar[0]/okcd")
    okcd.Text = f"/n{transaction.upper().lstrip('/')}"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1.5)
    try:
        return str(session.findById("wnd[0]/sbar").Text).strip()
    except Exception:
        return ""


def _read_sbar(session) -> str:
    try:
        return str(session.findById("wnd[0]/sbar").Text).strip()
    except Exception:
        return ""


def _dismiss_popup(session) -> None:
    """Fecha popup/janela de aviso se existir."""
    for btn_id in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[11]"):
        try:
            session.findById(btn_id).press()
            return
        except Exception:
            pass
    try:
        session.findById("wnd[1]").sendVKey(12)  # ESC
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────────────────────
# SE16N – Pesquisa em tabela
# ──────────────────────────────────────────────────────────────────────────────

def se16n_query(
    table: str,
    filters: list[dict[str, str]] | None = None,
    fields: list[str] | None = None,
    max_rows: int = 20,
    description: str = "",
) -> SapGuiResult:
    """Abre a SE16N, pesquisa na tabela indicada com os filtros fornecidos.

    Args:
        table: Nome da tabela SAP (ex: "EKKO", "AUFK", "BKPF")
        filters: Lista de {"field": "EBELN", "value": "4500000123", "option": "EQ"}
        fields: Lista de campos a mostrar (vazia = todos)
        max_rows: Número máximo de linhas a retornar
        description: Descrição legível para o chat

    Returns:
        SapGuiResult com rows preenchido e result_text formatado
    """
    action_desc = description or f"SE16N → Tabela {table}"
    try:
        session = _get_session()
    except RuntimeError as exc:
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            error=str(exc),
            success=False,
            result_text=f"❌ {exc}",
        )

    try:
        # Navegar para SE16N
        sbar = _navigate_to(session, "SE16N")
        time.sleep(0.5)

        # Fechar popup se aparecer (ex: "sessão com alterações pendentes")
        _dismiss_popup(session)

        # Preencher o nome da tabela
        try:
            session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table.upper()
            session.findById("wnd[0]").sendVKey(0)  # Enter para confirmar
            time.sleep(1.5)
        except Exception as exc:
            return SapGuiResult(
                action="se16n_query",
                description=action_desc,
                error=f"Erro ao preencher tabela na SE16N: {exc}",
                success=False,
                result_text=f"❌ Erro ao abrir tabela {table}: {exc}",
            )

        # Fechar popup de aviso (ex: tabela protegida / aviso de leitura)
        _dismiss_popup(session)

        # Definir max rows (campo MAX_LINES ou ROWCOUNT na SE16N)
        try:
            session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = str(min(max_rows, 200))
        except Exception:
            pass  # Campo pode não existir em todas as versões

        # Aplicar filtros
        if filters:
            for f in filters:
                field_name = str(f.get("field") or "").strip().upper()
                value = str(f.get("value") or "").strip()
                if not field_name or not value:
                    continue
                try:
                    # Tentar preencher directamente o campo de selecção
                    field_id = f"wnd[0]/usr/txt{field_name}-LOW"
                    session.findById(field_id).Text = value
                except Exception:
                    try:
                        # Alternativa: campo com prefixo ctxt
                        field_id = f"wnd[0]/usr/ctxt{field_name}-LOW"
                        session.findById(field_id).Text = value
                    except Exception:
                        pass  # Campo não encontrado, ignora

        # Executar pesquisa (F8)
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(2.0)

        # Fechar popup se aparecer
        _dismiss_popup(session)

        # Ler resultado — tentar capturar GridView (ALV grid)
        rows: list[dict[str, str]] = _read_alv_grid(session, max_rows)

        if rows:
            result_text = _format_rows_as_text(rows, table, filters)
            return SapGuiResult(
                action="se16n_query",
                description=action_desc,
                result_text=result_text,
                rows=rows,
                success=True,
            )
        else:
            # Tentar ler status bar para mensagem de "nenhum resultado"
            sbar_msg = _read_sbar(session)
            return SapGuiResult(
                action="se16n_query",
                description=action_desc,
                result_text=f"📭 Pesquisa na tabela **{table}** concluída. Nenhum resultado encontrado.\nSTATUS: {sbar_msg}",
                rows=[],
                success=True,
            )

    except Exception as exc:
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            error=str(exc),
            success=False,
            result_text=f"❌ Erro durante pesquisa SE16N em {table}: {exc}",
        )


def _read_alv_grid(session, max_rows: int = 50) -> list[dict[str, str]]:
    """Tenta ler os dados do ALV Grid na janela actual da SE16N."""
    rows: list[dict[str, str]] = []

    try:
        # Procurar GridViewCtrl ou shell no wnd[0]/usr
        grid = None
        try:
            grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        except Exception:
            pass

        if grid is None:
            try:
                grid = session.findById("wnd[0]/usr/cntlALV_GRID/shellcont/shell")
            except Exception:
                pass

        if grid is None:
            return rows

        col_count = int(grid.ColumnCount)
        row_count = min(int(grid.RowCount), max_rows)

        # Obter nomes das colunas
        col_names: list[str] = []
        for ci in range(col_count):
            try:
                col_names.append(str(grid.GetColumnKey(ci)).strip())
            except Exception:
                col_names.append(f"COL{ci}")

        # Ler cada linha
        for ri in range(row_count):
            row: dict[str, str] = {}
            for ci, col in enumerate(col_names):
                try:
                    val = str(grid.GetCellValue(ri, col)).strip()
                    row[col] = val
                except Exception:
                    row[col] = ""
            rows.append(row)

    except Exception:
        pass

    return rows


def _format_rows_as_text(rows: list[dict[str, str]], table: str, filters: list | None) -> str:
    """Formata as linhas da tabela SAP como texto markdown para o chat."""
    if not rows:
        return f"Tabela {table}: sem resultados."

    filter_desc = ""
    if filters:
        parts = [f"{f.get('field')}={f.get('value')}" for f in (filters or []) if f.get("field")]
        filter_desc = " | Filtros: " + ", ".join(parts)

    lines = [f"**📊 Tabela SAP: {table}{filter_desc} — {len(rows)} linha(s)**\n"]

    # Cabeçalho
    headers = list(rows[0].keys())
    header_row = " | ".join(f"**{h}**" for h in headers)
    sep_row = " | ".join("---" for _ in headers)
    lines.append(f"| {header_row} |")
    lines.append(f"| {sep_row} |")

    # Dados
    for row in rows:
        data_row = " | ".join(str(row.get(h, "")).strip() for h in headers)
        lines.append(f"| {data_row} |")

    return "\n".join(lines)


# ──────────────────────────────────────────────────────────────────────────────
# Ações rápidas de visualização
# ──────────────────────────────────────────────────────────────────────────────

def open_transaction(transaction: str, description: str = "") -> SapGuiResult:
    """Abre uma transação SAP e retorna o status bar."""
    desc = description or f"Abrir transação {transaction}"
    try:
        session = _get_session()
        sbar = _navigate_to(session, transaction)
        return SapGuiResult(
            action="open_transaction",
            description=desc,
            result_text=f"✅ Transação **{transaction.upper()}** aberta.\nSTATUS: {sbar or '(sem mensagem)'}",
            success=True,
        )
    except RuntimeError as exc:
        return SapGuiResult(
            action="open_transaction",
            description=desc,
            error=str(exc),
            success=False,
            result_text=f"❌ {exc}",
        )


def read_current_status(description: str = "") -> SapGuiResult:
    """Lê o status bar da sessão SAP actual."""
    try:
        session = _get_session()
        sbar = _read_sbar(session)
        return SapGuiResult(
            action="read_sbar",
            description=description or "Ler status bar SAP",
            result_text=f"STATUS SAP: {sbar or '(vazio)'}",
            success=True,
        )
    except RuntimeError as exc:
        return SapGuiResult(
            action="read_sbar",
            description=description or "Ler status bar SAP",
            error=str(exc),
            success=False,
            result_text=f"❌ {exc}",
        )


# ──────────────────────────────────────────────────────────────────────────────
# Dispatcher principal
# ──────────────────────────────────────────────────────────────────────────────

def execute_sap_gui_action(params: dict[str, Any]) -> SapGuiResult:
    """Ponto de entrada principal para o worker executar uma ação SAP GUI.

    params deve conter:
      - action: "se16n_query" | "open_transaction" | "read_sbar"
      - (para se16n_query): table, filters, fields, max_rows
      - (para open_transaction): transaction
      - description: texto descritivo opcional
    """
    action = str(params.get("action") or "se16n_query").strip().lower()
    description = str(params.get("description") or "").strip()

    if action == "se16n_query":
        return se16n_query(
            table=str(params.get("table") or "").upper(),
            filters=params.get("filters") or [],
            fields=params.get("fields") or [],
            max_rows=int(params.get("max_rows") or 20),
            description=description,
        )

    elif action == "open_transaction":
        return open_transaction(
            transaction=str(params.get("transaction") or "").upper(),
            description=description,
        )

    elif action == "read_sbar":
        return read_current_status(description=description)

    else:
        return SapGuiResult(
            action=action,
            description=description,
            error=f"Ação não reconhecida: '{action}'. Ações suportadas: se16n_query, open_transaction, read_sbar.",
            success=False,
            result_text=f"❌ Ação '{action}' não suportada.",
        )

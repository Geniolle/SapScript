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

AUTHORIZATION_ALLOWED_TABLES = {
    "USZBVSYS",
    "USLA04",
    "USL04",
    "USRSYSACTT",
}

AUTHORIZATION_ALLOWED_FIELDS = {
    "USZBVSYS": {
        "BNAME",
        "SUBSYSTEM",
    },
    "USLA04": {
        "BNAME",
        "SUBSYSTEM",
        "AGR_NAME",
        "FROM_DAT",
        "TO_DAT",
        "ORG_FLAG",
    },
    "USL04": {
        "BNAME",
        "SUBSYSTEM",
        "PROFILE",
    },
    "USRSYSACTT": {
        "SUBSYSTEM",
        "LANGU",
        "AGR_NAME",
        "TEXT",
    },
}

def _find_selection_field(container, suffix: str) -> Any:
    """Procura recursivamente um elemento Gui(C)TextField cujo ID ou Name termina com o sufixo."""
    try:
        name = str(getattr(container, "Name", "")).strip().upper()
        elem_id = str(getattr(container, "Id", "")).strip().upper()
        suffix_upper = suffix.strip().upper()
        
        if name.endswith(suffix_upper) or elem_id.endswith(suffix_upper):
            type_name = str(getattr(container, "Type", ""))
            if "TextField" in type_name or "CTextField" in type_name:
                if getattr(container, "Changeable", True):
                    return container
    except Exception:
        pass

    try:
        children = getattr(container, "Children", None)
        if children:
            for i in range(children.Count):
                found = _find_selection_field(children.Element(i), suffix)
                if found:
                    return found
    except Exception:
        pass

    return None

def se16n_query_with_session(
    session: Any,
    *,
    table: str,
    filters: list[dict[str, str]],
    fields: list[str] | None = None,
    max_rows: int = 5000,
    strict_filters: bool = True,
) -> SapGuiResult:
    table_upper = table.strip().upper()
    action_desc = f"SE16N → Tabela {table_upper}"

    # Se for execução estrita, barrar tabelas fora da allowlist
    if strict_filters and table_upper not in AUTHORIZATION_ALLOWED_TABLES:
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            error=f"Tabela '{table_upper}' não está na allowlist de autorizações.",
            success=False,
            result_text=f"❌ Tabela não autorizada: {table_upper}",
        )

    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16N"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        _dismiss_popup(session)
    except Exception as exc:
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            error=f"Erro ao navegar para SE16N: {exc}",
            success=False,
            result_text=f"❌ Erro ao navegar para SE16N: {exc}",
        )

    try:
        session.findById("wnd[0]/usr/ctxtGD-TAB").Text = table_upper
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        _dismiss_popup(session)
    except Exception as exc:
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            error=f"Erro ao preencher tabela GD-TAB na SE16N: {exc}",
            success=False,
            result_text=f"❌ Erro ao abrir tabela {table_upper} na SE16N: {exc}",
        )

    try:
        session.findById("wnd[0]/usr/txtGD-MAX_LINES").Text = str(max_rows)
    except Exception:
        pass

    applied_filters = []
    usr_area = session.findById("wnd[0]/usr")

    if filters:
        for f in filters:
            field_name = str(f.get("field") or "").strip().upper()
            value = str(f.get("value") or "").strip()
            if not field_name or not value:
                continue

            if strict_filters:
                allowed_fields = AUTHORIZATION_ALLOWED_FIELDS.get(table_upper, set())
                if field_name not in allowed_fields:
                    return SapGuiResult(
                        action="se16n_query",
                        description=action_desc,
                        error=f"Campo '{field_name}' na tabela '{table_upper}' não está na allowlist.",
                        success=False,
                        result_text=f"❌ Campo não autorizado na tabela {table_upper}.",
                    )

            suffix = f"{field_name}-LOW"
            field_element = _find_selection_field(usr_area, suffix)

            if field_element:
                try:
                    field_element.Text = value
                    time.sleep(0.1)
                    if str(field_element.Text).strip() == value:
                        applied_filters.append(field_name)
                    else:
                        field_element.Text = value
                        time.sleep(0.2)
                        if str(field_element.Text).strip() == value:
                            applied_filters.append(field_name)
                except Exception:
                    pass

            if field_name not in applied_filters:
                if strict_filters:
                    return SapGuiResult(
                        action="se16n_query",
                        description=action_desc,
                        error=f"Não foi possível aplicar de forma segura o filtro {field_name} na tabela {table_upper}.",
                        success=False,
                        result_text=f"❌ Filtro obrigatório não aplicado na tabela {table_upper}.",
                    )

    try:
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(2.0)
        _dismiss_popup(session)
    except Exception as exc:
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            error=f"Erro ao executar pesquisa na SE16N: {exc}",
            success=False,
            result_text=f"❌ Erro ao executar pesquisa: {exc}",
        )

    rows = _read_alv_grid(session, max_rows)

    # Voltar para a tela inicial
    try:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(0.5)
        _dismiss_popup(session)
    except Exception:
        pass

    if rows:
        result_text = _format_rows_as_text(rows, table_upper, filters)
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            result_text=result_text,
            rows=rows,
            success=True,
        )
    else:
        sbar_msg = _read_sbar(session)
        return SapGuiResult(
            action="se16n_query",
            description=action_desc,
            result_text=f"📭 Pesquisa na tabela **{table_upper}** concluída. Nenhum resultado encontrado.\nSTATUS: {sbar_msg}",
            rows=[],
            success=True,
        )


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

    # Note: Para compatibilidade legada, executamos com strict_filters=False
    return se16n_query_with_session(
        session,
        table=table,
        filters=filters or [],
        fields=fields,
        max_rows=max_rows,
        strict_filters=False,
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


def collect_sap_components(container) -> list[dict[str, Any]]:
    """Percorre recursivamente o container e recolhe propriedades dos componentes relevantes."""
    components = []
    
    def walk(node):
        if node is None:
            return
        
        comp_info = {}
        for prop in ["Id", "Name", "Type", "Text", "Changeable", "Left", "Top", "Width", "Height"]:
            try:
                val = getattr(node, prop, None)
                if prop in ["Left", "Top", "Width", "Height"]:
                    comp_info[prop.lower()] = int(val) if val is not None else 0
                elif prop == "Changeable":
                    comp_info[prop.lower()] = bool(val) if val is not None else False
                else:
                    comp_info[prop.lower()] = str(val) if val is not None else ""
            except Exception:
                comp_info[prop.lower()] = 0 if prop in ["Left", "Top", "Width", "Height"] else (False if prop == "Changeable" else "")
        
        comp_type = comp_info.get("type", "")
        if comp_type in ("GuiLabel", "GuiTextField", "GuiCTextField"):
            is_visible = True
            try:
                if getattr(node, "Visible", True) is False:
                    is_visible = False
            except Exception:
                pass
            if is_visible and comp_info.get("width", 0) > 0 and comp_info.get("height", 0) > 0:
                comp_info["element"] = node
                components.append(comp_info)
        
        try:
            children = getattr(node, "Children", None)
            if children:
                for i in range(children.Count):
                    walk(children.Element(i))
        except Exception:
            pass

    walk(container)
    return components


def find_se16_low_field_by_label(session, field_name: str) -> Any:
    """Localiza o campo LOW para o field_name na tela de seleção SE16 clássica usando posicionamento por label."""
    try:
        usr = session.findById("wnd[0]/usr")
    except Exception:
        return None
        
    components = collect_sap_components(usr)
    
    field_upper = field_name.upper()
    translated_terms = {
        "BNAME": ["BNAME", "UTILIZADOR", "USER", "NOME"],
        "SUBSYSTEM": ["SUBSYSTEM", "SISTEMA", "SYSTEM", "LOGICAL SYSTEM"]
    }.get(field_upper, [field_upper])
    
    target_label = None
    for c in components:
        if c.get("type") == "GuiLabel":
            text = c.get("text", "").strip().upper()
            if text in translated_terms or text.replace(":", "") in translated_terms or any(term == text for term in translated_terms):
                target_label = c
                break
            if any(term in text for term in translated_terms):
                target_label = c
                break

    if not target_label:
        return None
    
    label_top = target_label.get("top", 0)
    label_left = target_label.get("left", 0)
    
    candidates = []
    for c in components:
        if c.get("type") not in ("GuiTextField", "GuiCTextField"):
            continue
        if not c.get("changeable"):
            continue
            
        tf_top = c.get("top", 0)
        tf_left = c.get("left", 0)
        
        # Mesma linha vertical com tolerância de 8 pixels
        if abs(tf_top - label_top) <= 8:
            if tf_left > label_left:
                candidates.append(c)
                
    if not candidates:
        return None
        
    candidates.sort(key=lambda x: x.get("left", 0))
    best_candidate = candidates[0]
    return best_candidate.get("element")


def _find_se16_field(session, field_name: str) -> Any:
    """Procura um elemento de seleção na SE16 para o campo dado."""
    try:
        usr = session.findById("wnd[0]/usr")
    except Exception:
        return None
        
    field_upper = field_name.upper()

    # 1. Tentar Name/Id exato ou contendo o nome técnico (fallback semântico por Name e ID)
    # Procurar primeiro na árvore por elementos editáveis contendo o nome técnico no ID ou Name (evitando -HIGH)
    components = collect_sap_components(usr)
    editables = [c for c in components if c.get("changeable") and c.get("type") in ("GuiTextField", "GuiCTextField")]
    
    for c in editables:
        name = c.get("name", "").upper()
        elem_id = c.get("id", "").upper()
        if field_upper in name or field_upper in elem_id:
            if "-HIGH" not in elem_id and "-HIGH" not in name:
                return c.get("element")

    # 2. Se não encontrar por ID/Name semântico, usar associação label -> campo pela posição
    elem = find_se16_low_field_by_label(session, field_name)
    if elem:
        return elem
        
    # 3. Outros fallbacks baseados em sufixos clássicos conhecidos
    for suffix in [f"{field_upper}-LOW", f"SO_{field_upper[:4]}-LOW", f"GD-{field_upper}-LOW", field_upper]:
        elem = _find_selection_field(usr, suffix)
        if elem:
            return elem

    # 4. Fallback posicional para tabelas conhecidas
    if field_upper == "BNAME":
        for idx in ("I1-LOW", "I1"):
            elem = _find_selection_field(usr, idx)
            if elem:
                return elem
    elif field_upper == "SUBSYSTEM":
        for idx in ("I2-LOW", "I2"):
            elem = _find_selection_field(usr, idx)
            if elem:
                return elem

    return None


def wait_for_se16_selection_screen(session, table: str, timeout_s: float = 15.0) -> bool:
    """Aguarda até que o ecrã de seleção da SE16 para a tabela dada esteja carregado e editável."""
    start_time = time.time()
    table_upper = table.upper()
    while time.time() - start_time < timeout_s:
        try:
            active_window = session.ActiveWindow
            title = str(getattr(active_window, "Text", "")).upper()
            if table_upper not in title:
                time.sleep(0.5)
                continue
            
            usr = session.findById("wnd[0]/usr")
            components = collect_sap_components(usr)
            editables = [c for c in components if c.get("changeable") and c.get("type") in ("GuiTextField", "GuiCTextField")]
            
            if not editables:
                time.sleep(0.5)
                continue
                
            if table_upper == "USZBVSYS":
                has_bname = any("BNAME" in str(c.get("text", "")).upper() for c in components)
                if not has_bname:
                    time.sleep(0.5)
                    continue
            
            return True
        except Exception:
            time.sleep(0.5)
    return False


def _find_grid_view(container) -> Any:
    for path in ("wnd[0]/usr/cntlGRID1/shellcont/shell", "wnd[0]/usr/cntlALV_GRID/shellcont/shell"):
        try:
            grid = container.findById(path)
            if grid:
                return grid
        except Exception:
            pass
    
    # Busca recursiva no container
    try:
        t = str(getattr(container, "Type", ""))
        if "GridView" in t or ("Shell" in t and "GridView" in str(getattr(container, "SubType", ""))):
            return container
    except Exception:
        pass
        
    try:
        children = getattr(container, "Children", None)
        if children:
            for i in range(children.Count):
                found = _find_grid_view(children.Element(i))
                if found:
                    return found
    except Exception:
        pass
        
    return None


import re

def _read_se16_output(session, max_rows: int = 5000) -> list[dict[str, str]]:
    usr = session.findById("wnd[0]/usr")
    
    # 1. Fallback: Se for uma GridView (ALV), usar a rotina existente
    grid_view = _find_grid_view(session)
    if grid_view:
        return _read_alv_grid(session, max_rows)
        
    # 2. Caso contrário, tratar como lista clássica ABAP
    cells = []
    try:
        for i in range(usr.Children.Count):
            child = usr.Children.Element(i)
            child_id = str(getattr(child, "Id", ""))
            child_type = str(getattr(child, "Type", ""))
            
            if "lbl[" in child_id or "txt[" in child_id or child_type in ("GuiLabel", "GuiTextField"):
                text = str(getattr(child, "Text", "")).strip()
                match = re.search(r"\[(\d+),(\d+)\]", child_id)
                if match:
                    row_idx = int(match.group(1))
                    col_idx = int(match.group(2))
                    cells.append({
                        "row": row_idx,
                        "col": col_idx,
                        "text": text
                    })
    except Exception:
        pass

    if not cells:
        return []

    rows_map = {}
    for cell in cells:
        row_idx = cell["row"]
        if row_idx not in rows_map:
            rows_map[row_idx] = []
        rows_map[row_idx].append(cell)

    sorted_rows = []
    for row_idx in sorted(rows_map.keys()):
        row_cells = sorted(rows_map[row_idx], key=lambda x: x["col"])
        full_text = "".join(c["text"] for c in row_cells)
        if not full_text or re.match(r"^[-|_+ *]+$", full_text):
            continue
        sorted_rows.append((row_idx, row_cells))

    cleaned_rows = []
    for row_idx, row_cells in sorted_rows:
        filtered_cells = [c for c in row_cells if c["text"] != "|"]
        if filtered_cells:
            cleaned_rows.append(filtered_cells)

    if not cleaned_rows:
        return []

    header_row_idx = -1
    known_headers = {"MANDT", "BNAME", "SUBSYSTEM", "AGR_NAME", "FROM_DAT", "TO_DAT", "ORG_FLAG", "PROFILE"}
    
    for idx, row in enumerate(cleaned_rows):
        row_texts = [c["text"].upper() for c in row]
        if any(h in row_texts for h in known_headers):
            header_row_idx = idx
            break

    if header_row_idx == -1:
        header_row_idx = 0

    parsed_rows = []
    for row in cleaned_rows[header_row_idx + 1:]:
        row_data = {}
        for c in row:
            col_pos = c["col"]
            best_header = None
            min_dist = 999999
            for h_cell in cleaned_rows[header_row_idx]:
                dist = abs(h_cell["col"] - col_pos)
                if dist < min_dist:
                    min_dist = dist
                    best_header = h_cell["text"].upper()
            if best_header:
                row_data[best_header] = c["text"]
        
        if row_data:
            parsed_rows.append(row_data)

    return parsed_rows[:max_rows]


def se16_query_with_session(
    session: Any,
    *,
    table: str,
    filters: list[dict[str, str]],
    fields: list[str] | None = None,
    max_rows: int = 5000,
    strict_filters: bool = True,
) -> SapGuiResult:
    table_upper = table.strip().upper()
    action_desc = f"SE16 → Tabela {table_upper}"

    if strict_filters and table_upper not in AUTHORIZATION_ALLOWED_TABLES:
        return SapGuiResult(
            action="se16_query",
            description=action_desc,
            error=f"Tabela '{table_upper}' não está na allowlist de autorizações.",
            success=False,
            result_text=f"❌ Tabela não autorizada: {table_upper}",
        )

    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSE16"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        _dismiss_popup(session)
    except Exception as exc:
        return SapGuiResult(
            action="se16_query",
            description=action_desc,
            error=f"Erro ao navegar para SE16: {exc}",
            success=False,
            result_text=f"❌ Erro ao navegar para SE16: {exc}",
        )

    try:
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = table_upper
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1.5)
        _dismiss_popup(session)
    except Exception as exc:
        return SapGuiResult(
            action="se16_query",
            description=action_desc,
            error=f"Erro ao preencher tabela DATABROWSE-TABLENAME na SE16: {exc}",
            success=False,
            result_text=f"❌ Erro ao abrir tabela {table_upper} na SE16: {exc}",
        )

    # ESPERAR O ECRÃ DE SELEÇÃO
    screen_ready = wait_for_se16_selection_screen(session, table_upper, timeout_s=15)
    if not screen_ready:
        print(f"[AUTH][SE16][WARN] Ecrã de seleção para {table_upper} não carregou completamente no tempo limite.")
    else:
        print(f"[AUTH][SE16] Ecrã de seleção {table_upper} carregado.")

    try:
        max_lines_elem = _find_selection_field(session.findById("wnd[0]/usr"), "MAX_LINES")
        if max_lines_elem:
            max_lines_elem.Text = str(max_rows)
    except Exception:
        pass

    applied_filters = []
    if filters:
        for f in filters:
            field_name = str(f.get("field") or "").strip().upper()
            value = str(f.get("value") or "").strip()
            if not field_name or not value:
                continue

            if strict_filters:
                allowed_fields = AUTHORIZATION_ALLOWED_FIELDS.get(table_upper, set())
                if field_name not in allowed_fields:
                    return SapGuiResult(
                        action="se16_query",
                        description=action_desc,
                        error=f"Campo '{field_name}' na tabela '{table_upper}' não está na allowlist.",
                        success=False,
                        result_text=f"❌ Campo não autorizado na tabela {table_upper}.",
                    )

            print(f"[AUTH][SE16] A localizar filtro {field_name}.")
            
            usr_area = session.findById("wnd[0]/usr")
            components = collect_sap_components(usr_area)
            translated_terms = {
                "BNAME": ["BNAME", "UTILIZADOR", "USER", "NOME"],
                "SUBSYSTEM": ["SUBSYSTEM", "SISTEMA", "SYSTEM", "LOGICAL SYSTEM"]
            }.get(field_name, [field_name])
            
            label_found = False
            for comp in components:
                if comp.get("type") == "GuiLabel":
                    text = comp.get("text", "").strip().upper()
                    if text in translated_terms or text.replace(":", "") in translated_terms or any(term in text for term in translated_terms):
                        label_found = True
                        break
            if label_found:
                print(f"[AUTH][SE16] Label {field_name} encontrada.")
            else:
                print(f"[AUTH][SE16] Label {field_name} não encontrada.")

            field_element = _find_se16_field(session, field_name)

            if not field_element:
                labels_found = [comp.get("text", "").strip() for comp in components if comp.get("type") == "GuiLabel"]
                editables_found = [comp.get("id", "").strip() for comp in components if comp.get("changeable") and comp.get("type") in ("GuiTextField", "GuiCTextField")]
                error_msg = f"Não foi possível localizar o campo {field_name} no ecrã de seleção da tabela {table_upper} na SE16."
                tech_log = (
                    f"Tabela: {table_upper}\n"
                    f"Filtro ausente: {field_name}\n"
                    f"Labels encontradas: {labels_found}\n"
                    f"Quantidade de campos editáveis: {len(editables_found)}\n"
                    f"IDs candidatos: {editables_found}"
                )
                print(f"[AUTH][SE16][ERROR] {error_msg}\nDetalhes técnicos:\n{tech_log}")
                if strict_filters:
                    return SapGuiResult(
                        action="se16_query",
                        description=action_desc,
                        error=f"{error_msg}\n{tech_log}",
                        success=False,
                        result_text=f"❌ {error_msg}",
                    )
                continue

            try:
                field_id = getattr(field_element, "Id", "N/D")
                print(f"[AUTH][SE16][DEBUG] {field_name} field id: {field_id}")
                print(f"[AUTH][SE16] Campo LOW de {field_name} encontrado.")
            except Exception:
                pass

            try:
                field_element.setFocus()
                field_element.Text = value
                time.sleep(0.1)
                written_value = str(field_element.Text).strip()
                if written_value.upper() == value.upper():
                    applied_filters.append(field_name)
                else:
                    field_element.Text = value
                    time.sleep(0.2)
                    written_value = str(field_element.Text).strip()
                    if written_value.upper() == value.upper():
                        applied_filters.append(field_name)
            except Exception as exc:
                print(f"[AUTH][SE16][ERROR] Erro ao preencher campo {field_name}: {exc}")
                pass

            if field_name in applied_filters:
                print(f"[AUTH][SE16] {field_name} preenchido e validado.")
            else:
                error_msg = f"Não foi possível localizar o campo {field_name} no ecrã de seleção da tabela {table_upper} na SE16."
                print(f"[AUTH][SE16][ERROR] {error_msg}")
                if strict_filters:
                    return SapGuiResult(
                        action="se16_query",
                        description=action_desc,
                        error=error_msg,
                        success=False,
                        result_text=f"❌ {error_msg}",
                    )

        print("[AUTH][SE16] Todos os filtros obrigatórios foram aplicados. A executar F8.")

    try:
        session.findById("wnd[0]").sendVKey(8)
        time.sleep(2.0)
        _dismiss_popup(session)
    except Exception as exc:
        return SapGuiResult(
            action="se16_query",
            description=action_desc,
            error=f"Erro ao executar pesquisa na SE16: {exc}",
            success=False,
            result_text=f"❌ Erro ao executar pesquisa: {exc}",
        )

    rows = _read_se16_output(session, max_rows)

    try:
        session.findById("wnd[0]").sendVKey(3)
        time.sleep(0.5)
        _dismiss_popup(session)
    except Exception:
        pass

    if rows:
        result_text = _format_rows_as_text(rows, table_upper, filters)
        return SapGuiResult(
            action="se16_query",
            description=action_desc,
            result_text=result_text,
            rows=rows,
            success=True,
        )
    else:
        sbar_msg = _read_sbar(session)
        return SapGuiResult(
            action="se16_query",
            description=action_desc,
            result_text=f"📭 Pesquisa na tabela **{table_upper}** concluída. Nenhum resultado encontrado.\nSTATUS: {sbar_msg}",
            rows=[],
            success=True,
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

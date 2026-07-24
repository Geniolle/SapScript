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


def get_sap_children(obj):
    children = []
    try:
        collection = obj.Children
        count = int(collection.Count)
    except Exception:
        return children

    for index in range(count):
        child = None
        try:
            child = collection.Item(index)
        except Exception:
            try:
                child = collection(index)
            except Exception:
                try:
                    child = collection.Element(index)
                except Exception:
                    child = None
        if child is not None:
            children.append(child)
    return children


def collect_sap_components(root, max_depth=12, max_nodes=3000):
    result = []
    visited_ids = set()

    def visit(component, depth):
        if component is None:
            return
        if depth > max_depth:
            return
        if len(result) >= max_nodes:
            return

        try:
            component_id = str(getattr(component, "Id", ""))
        except Exception:
            component_id = ""

        visit_key = component_id or id(component)
        if visit_key in visited_ids:
            return

        visited_ids.add(visit_key)
        result.append(component)

        for child in get_sap_children(component):
            visit(child, depth + 1)

    visit(root, 0)
    return result


def describe_sap_component(component):
    def safe_attr(name, default=None):
        try:
            return getattr(component, name)
        except Exception:
            return default

    def safe_int(val):
        try:
            return int(val) if val is not None else None
        except (ValueError, TypeError):
            return None

    return {
        "id": str(safe_attr("Id", "") or ""),
        "name": str(safe_attr("Name", "") or ""),
        "type": str(safe_attr("Type", "") or ""),
        "text": str(safe_attr("Text", "") or ""),
        "changeable": safe_attr("Changeable", None),
        "enabled": safe_attr("Enabled", None),
        "visible": safe_attr("Visible", None),
        "left": safe_int(safe_attr("Left", None)),
        "top": safe_int(safe_attr("Top", None)),
        "width": safe_int(safe_attr("Width", None)),
        "height": safe_int(safe_attr("Height", None)),
    }


def is_sap_true(value):
    if value is True:
        return True
    if isinstance(value, int):
        return value != 0
    return str(value or "").strip().lower() in {"true", "1", "-1", "yes"}


def is_editable_sap_field(component) -> bool:
    desc = describe_sap_component(component)
    comp_type = desc["type"]
    if comp_type not in {"GuiTextField", "GuiCTextField", "GuiPasswordField"}:
        return False

    enabled = desc["enabled"]
    if enabled is False or str(enabled).strip().lower() in {"false", "0", "no"}:
        return False

    changeable = desc["changeable"]
    if changeable is None:
        return True
    return is_sap_true(changeable)


def normalize_label(value):
    return str(value or "").strip().upper().rstrip(":")


def is_sap_label(component, label_name: str) -> bool:
    desc = describe_sap_component(component)
    if desc["type"] in {"GuiTextField", "GuiCTextField", "GuiPasswordField"}:
        return False

    norm_text = normalize_label(desc["text"])
    norm_target = normalize_label(label_name)

    if norm_text == norm_target:
        return True

    if desc["type"] == "GuiLabel":
        translated_terms = {
            "BNAME": ["BNAME", "UTILIZADOR", "USER", "NOME"],
            "SUBSYSTEM": ["SUBSYSTEM", "SISTEMA", "SYSTEM", "LOGICAL SYSTEM"]
        }.get(norm_target, [norm_target])
        if norm_text in translated_terms or any(term in norm_text for term in translated_terms):
            return True

    return False


def get_se16_usr_components(session) -> list:
    raw_components = []
    try:
        wnd0 = session.findById("wnd[0]")
        if wnd0:
            raw_components = collect_sap_components(wnd0)
    except Exception:
        pass

    try:
        usr = session.findById("wnd[0]/usr")
        if usr:
            usr_comps = collect_sap_components(usr)
            seen = set()
            combined = []
            for c in raw_components + usr_comps:
                try:
                    c_id = str(getattr(c, "Id", ""))
                except Exception:
                    c_id = ""
                key = c_id or id(c)
                if key not in seen:
                    seen.add(key)
                    combined.append(c)
            raw_components = combined
    except Exception:
        pass

    filtered = []
    for c in raw_components:
        try:
            c_id = str(getattr(c, "Id", ""))
        except Exception:
            c_id = ""
        if "/usr/" in c_id or "wnd[0]" not in c_id:
            filtered.append(c)
    return filtered


def find_se16_low_field_by_label_with_components(components: list, field_name: str) -> Any:
    field_upper = field_name.upper()
    target_label = None
    for c in components:
        if is_sap_label(c, field_upper):
            target_label = describe_sap_component(c)
            break

    if not target_label:
        return None

    label_top = target_label.get("top")
    label_left = target_label.get("left")
    if label_top is None or label_left is None:
        return None

    candidates = []
    for c in components:
        if not is_editable_sap_field(c):
            continue

        desc = describe_sap_component(c)
        tf_top = desc.get("top")
        tf_left = desc.get("left")
        if tf_top is None or tf_left is None:
            continue

        v_diff = abs(tf_top - label_top)
        if v_diff <= 8:
            if tf_left > label_left:
                candidates.append((v_diff, tf_left - label_left, c))

    if not candidates:
        return None

    candidates.sort(key=lambda x: (x[0], x[1]))
    best_candidate = candidates[0][2]
    return best_candidate


def find_se16_low_field_by_label(session, field_name: str) -> Any:
    """Localiza o campo LOW para o field_name na tela de seleção SE16 clássica usando posicionamento por label."""
    components = get_se16_usr_components(session)
    return find_se16_low_field_by_label_with_components(components, field_name)


def get_direct_user_area_children(session):
    user_area = session.findById("wnd[0]/usr")
    collection = user_area.Children
    count = int(collection.Count)
    result = []

    for index in range(count):
        component = None
        try:
            component = collection.Item(index)
        except Exception:
            try:
                component = collection(index)
            except Exception:
                try:
                    component = collection.Element(index)
                except Exception:
                    component = None

        if component is not None:
            result.append(
                {
                    "index": index,
                    "component": component,
                }
            )
    return result


def normalize_se16_caption(value):
    return str(value or "").strip().upper().rstrip(":")


def try_write_se16_field(component, value):
    try:
        try:
            component.setFocus()
        except Exception:
            pass

        component.Text = str(value)
        time.sleep(0.1)
        actual = str(getattr(component, "Text", "") or "").strip()
        return actual.upper() == str(value).strip().upper()
    except Exception:
        return False


class Se16FilterNotFoundError(RuntimeError):
    pass


def safe_component_id(component):
    try:
        return str(getattr(component, "Id", ""))
    except Exception:
        return ""


def find_and_fill_se16_low_field(session, *, field_name, value):
    children = get_direct_user_area_children(session)
    expected_caption = normalize_se16_caption(field_name)

    translated_terms = {
        "BNAME": ["BNAME", "UTILIZADOR", "USER", "NOME"],
        "SUBSYSTEM": ["SUBSYSTEM", "SISTEMA", "SYSTEM", "LOGICAL SYSTEM"]
    }.get(expected_caption, [expected_caption])

    for position, entry in enumerate(children):
        component = entry["component"]
        try:
            text = normalize_se16_caption(component.Text)
        except Exception:
            text = ""

        if text != expected_caption and text not in translated_terms and not any(term in text for term in translated_terms):
            continue

        candidate_positions = [
            position + 1,
            position + 2,
        ]

        for candidate_position in candidate_positions:
            if candidate_position >= len(children):
                continue

            candidate = children[candidate_position]["component"]
            try:
                cand_text = str(getattr(candidate, "Text", "")).strip().lower()
                cand_type = str(getattr(candidate, "Type", ""))
                if "Button" in cand_type:
                    continue
                if cand_text in ("to", "até", "a"):
                    continue
                if normalize_se16_caption(cand_text) in {"BNAME", "SUBSYSTEM"}:
                    continue
            except Exception:
                pass

            if try_write_se16_field(candidate, value):
                return {
                    "field_name": field_name,
                    "caption_index": entry["index"],
                    "input_index": children[candidate_position]["index"],
                    "input_id": safe_component_id(candidate),
                    "component": candidate,
                }

    raise Se16FilterNotFoundError(
        f"Não foi possível localizar o campo LOW de {field_name} na SE16."
    )


def fill_se16_field_with_fallbacks(session, field_name: str, value: str, table: str) -> Any:
    field_upper = field_name.upper()
    table_upper = table.upper()

    # 1. iteração sequencial de children + próximo elemento
    try:
        res = find_and_fill_se16_low_field(session, field_name=field_upper, value=value)
        if res:
            return res["component"]
    except Exception:
        pass

    # 2. Fallbacks de IDs específicos para USZBVSYS
    if table_upper == "USZBVSYS":
        specific_ids = {
            "BNAME": [
                "wnd[0]/usr/ctxtI1-LOW",
                "wnd[0]/usr/txtI1-LOW",
            ],
            "SUBSYSTEM": [
                "wnd[0]/usr/ctxtI3-LOW",
                "wnd[0]/usr/txtI3-LOW",
            ]
        }.get(field_upper, [])

        for spec_id in specific_ids:
            try:
                comp = session.findById(spec_id)
                if comp and try_write_se16_field(comp, value):
                    return comp
            except Exception:
                pass

    # 3. Fallback de IDs semânticos/conhecidos gerais
    general_ids = {
        "BNAME": [
            "wnd[0]/usr/txtBNAME-LOW",
            "wnd[0]/usr/txtI1-LOW",
            "wnd[0]/usr/txtI1",
            "wnd[0]/usr/ctxtBNAME-LOW",
            "wnd[0]/usr/ctxtI1-LOW",
            "wnd[0]/usr/ctxtI1",
        ],
        "SUBSYSTEM": [
            "wnd[0]/usr/ctxtSUBSYSTEM-LOW",
            "wnd[0]/usr/ctxtI2-LOW",
            "wnd[0]/usr/ctxtI2",
            "wnd[0]/usr/txtSUBSYSTEM-LOW",
            "wnd[0]/usr/txtI2-LOW",
            "wnd[0]/usr/txtI2",
        ]
    }.get(field_upper, [])

    for gen_id in general_ids:
        try:
            comp = session.findById(gen_id)
            if comp and try_write_se16_field(comp, value):
                return comp
        except Exception:
            pass

    # 4. Fallback geométrico (Left/Top)
    try:
        comp = find_se16_low_field_by_label(session, field_upper)
        if comp and try_write_se16_field(comp, value):
            return comp
    except Exception:
        pass

    # 5. Fallback por sufixos clássicos do _find_selection_field
    usr_area = None
    try:
        usr_area = session.findById("wnd[0]/usr")
    except Exception:
        pass

    if usr_area:
        for suffix in [f"{field_upper}-LOW", f"SO_{field_upper[:4]}-LOW", f"GD-{field_upper}-LOW", field_upper]:
            try:
                comp = _find_selection_field(usr_area, suffix)
                if comp and try_write_se16_field(comp, value):
                    return comp
            except Exception:
                pass

        if field_upper == "BNAME":
            for idx in ("I1-LOW", "I1"):
                try:
                    comp = _find_selection_field(usr_area, idx)
                    if comp and try_write_se16_field(comp, value):
                        return comp
                except Exception:
                    pass
        elif field_upper == "SUBSYSTEM":
            for idx in ("I2-LOW", "I2"):
                try:
                    comp = _find_selection_field(usr_area, idx)
                    if comp and try_write_se16_field(comp, value):
                        return comp
                except Exception:
                    pass

    return None


def _find_se16_field(session, field_name: str) -> Any:
    """Procura um elemento de seleção na SE16 para o campo dado (compatibilidade com testes)."""
    field_upper = field_name.upper()

    try:
        children = get_direct_user_area_children(session)
        expected_caption = normalize_se16_caption(field_upper)
        translated_terms = {
            "BNAME": ["BNAME", "UTILIZADOR", "USER", "NOME"],
            "SUBSYSTEM": ["SUBSYSTEM", "SISTEMA", "SYSTEM", "LOGICAL SYSTEM"]
        }.get(expected_caption, [expected_caption])

        for position, entry in enumerate(children):
            component = entry["component"]
            try:
                text = normalize_se16_caption(component.Text)
            except Exception:
                text = ""

            if text == expected_caption or text in translated_terms or any(term in text for term in translated_terms):
                candidate_positions = [position + 1, position + 2]
                for pos in candidate_positions:
                    if pos < len(children):
                        cand = children[pos]["component"]
                        try:
                            cand_type = str(getattr(cand, "Type", ""))
                            cand_text = str(getattr(cand, "Text", "")).strip().lower()
                            if "Button" not in cand_type and cand_text not in ("to", "até", "a"):
                                return cand
                        except Exception:
                            pass
    except Exception:
        pass

    components = get_se16_usr_components(session)
    editables = [c for c in components if is_editable_sap_field(c)]

    general_ids = {
        "BNAME": [
            "wnd[0]/usr/txtBNAME-LOW",
            "wnd[0]/usr/txtI1-LOW",
            "wnd[0]/usr/txtI1",
            "wnd[0]/usr/ctxtBNAME-LOW",
            "wnd[0]/usr/ctxtI1-LOW",
            "wnd[0]/usr/ctxtI1",
        ],
        "SUBSYSTEM": [
            "wnd[0]/usr/ctxtSUBSYSTEM-LOW",
            "wnd[0]/usr/ctxtI2-LOW",
            "wnd[0]/usr/ctxtI2",
            "wnd[0]/usr/txtSUBSYSTEM-LOW",
            "wnd[0]/usr/txtI2-LOW",
            "wnd[0]/usr/txtI2",
        ]
    }.get(field_upper, [])

    for c in editables:
        desc = describe_sap_component(c)
        if desc["id"] in general_ids:
            return c

    for c in editables:
        desc = describe_sap_component(c)
        name = desc["name"].upper()
        elem_id = desc["id"].upper()
        if field_upper in name or field_upper in elem_id:
            if "-HIGH" not in elem_id and "-HIGH" not in name:
                return c

    elem = find_se16_low_field_by_label_with_components(components, field_name)
    if elem:
        return elem

    usr_area = None
    try:
        usr_area = session.findById("wnd[0]/usr")
    except Exception:
        pass

    if usr_area:
        for suffix in [f"{field_upper}-LOW", f"SO_{field_upper[:4]}-LOW", f"GD-{field_upper}-LOW", field_upper]:
            elem = _find_selection_field(usr_area, suffix)
            if elem:
                return elem

        if field_upper == "BNAME":
            for idx in ("I1-LOW", "I1"):
                elem = _find_selection_field(usr_area, idx)
                if elem:
                    return elem
        elif field_upper == "SUBSYSTEM":
            for idx in ("I2-LOW", "I2"):
                elem = _find_selection_field(usr_area, idx)
                if elem:
                    return elem

    return None


class Se16FieldDiscoveryError(RuntimeError):
    pass


def build_safe_discovery_error(snapshot):
    labels = snapshot.get("labels", [])
    editable_fields = snapshot.get("editable_fields", [])
    labels_text = [l.get("text", "") for l in labels]
    editables_ids = [e.get("id", "") for e in editable_fields]
    return (
        f"Não foi possível localizar os campos obrigatórios da SE16.\n"
        f"Labels encontradas: {labels_text}\n"
        f"Campos editáveis encontrados: {editables_ids}"
    )


def wait_for_se16_fields(session, timeout_s=15):
    deadline = time.monotonic() + timeout_s
    last_snapshot = {"labels": [], "editable_fields": []}

    while time.monotonic() < deadline:
        try:
            children = get_direct_user_area_children(session)
            labels = []
            editable_fields = []

            for entry in children:
                comp = entry["component"]
                try:
                    text = normalize_se16_caption(comp.Text)
                    comp_type = str(getattr(comp, "Type", ""))
                except Exception:
                    text = ""
                    comp_type = ""

                translated_bname = ["BNAME", "UTILIZADOR", "USER", "NOME"]
                translated_sub = ["SUBSYSTEM", "SISTEMA", "SYSTEM", "LOGICAL SYSTEM"]

                if text in translated_bname or text in translated_sub:
                    labels.append({
                        "id": safe_component_id(comp),
                        "text": text,
                        "type": comp_type
                    })

                if "TextField" in comp_type or "CTextField" in comp_type:
                    editable_fields.append({
                        "id": safe_component_id(comp),
                        "type": comp_type
                    })

            last_snapshot = {
                "labels": labels,
                "editable_fields": editable_fields,
            }
            if labels and editable_fields:
                return [entry["component"] for entry in children]
        except Exception:
            pass
        time.sleep(0.4)

    raise Se16FieldDiscoveryError(
        build_safe_discovery_error(last_snapshot)
    )


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
            
            components = wait_for_se16_fields(session, timeout_s=2.0)
            if components:
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
    filters_applied = {}
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
            
            components = get_se16_usr_components(session)
            total_n = len(components)
            labels_n = sum(1 for c in components if describe_sap_component(c)["type"] == "GuiLabel")
            text_fields_n = sum(1 for c in components if describe_sap_component(c)["type"] in ("GuiTextField", "GuiCTextField", "GuiPasswordField"))
            containers_n = total_n - labels_n - text_fields_n
            print(f"[AUTH][SE16][DEBUG] Total de componentes: {total_n}")
            print(f"[AUTH][SE16][DEBUG] Labels encontrados: {labels_n}")
            print(f"[AUTH][SE16][DEBUG] Campos de texto encontrados: {text_fields_n}")
            print(f"[AUTH][SE16][DEBUG] Contentores encontrados: {containers_n}")
            
            editables = [c for c in components if is_editable_sap_field(c)]
            for c in editables:
                desc = describe_sap_component(c)
                print(f"[AUTH][SE16][DEBUG] Candidato - Id: {desc['id']}, Name: {desc['name']}, Type: {desc['type']}, Text: {desc['text']}, Changeable: {desc['changeable']}, Left: {desc['left']}, Top: {desc['top']}")

            field_element = _find_se16_field(session, field_name)

            if not field_element:
                labels_found = [describe_sap_component(comp).get("text", "").strip() for comp in components if describe_sap_component(comp).get("type") == "GuiLabel"]
                editables_found = [describe_sap_component(comp).get("id", "").strip() for comp in components if is_editable_sap_field(comp)]
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
                    filters_applied[field_name] = True
                    applied_filters.append(field_name)
                else:
                    field_element.Text = value
                    time.sleep(0.2)
                    written_value = str(field_element.Text).strip()
                    if written_value.upper() == value.upper():
                        filters_applied[field_name] = True
                        applied_filters.append(field_name)
            except Exception as exc:
                print(f"[AUTH][SE16][ERROR] Erro ao preencher campo {field_name}: {exc}")
                pass

            if filters_applied.get(field_name):
                print(f"[AUTH][SE16] {field_name} preenchido e validado.")
            else:
                error_msg = f"Não foi possível preencher/validar o campo {field_name} no ecrã de seleção da tabela {table_upper} na SE16."
                print(f"[AUTH][SE16][ERROR] {error_msg}")
                if strict_filters:
                    return SapGuiResult(
                        action="se16_query",
                        description=action_desc,
                        error=error_msg,
                        success=False,
                        result_text=f"❌ {error_msg}",
                    )

        # Validar se todos os filtros necessários foram aplicados antes de prosseguir
        expected_filters = {str(f.get("field")).strip().upper() for f in filters if f.get("field")}
        applied_keys = {k for k, v in filters_applied.items() if v}
        
        if strict_filters and expected_filters != applied_keys:
            error_msg = f"Validação de filtros falhou. Filtros esperados: {expected_filters}, aplicados: {applied_keys}."
            print(f"[AUTH][SE16][ERROR] {error_msg}")
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

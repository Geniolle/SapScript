# -*- coding: utf-8 -*-
"""
Analise_Configuracao_SAP.py

Leitura genérica e SOMENTE DE CONSULTA de tabelas SAP através do SAP GUI
Scripting. O objetivo é permitir alterar, diretamente no VS Code, as tabelas,
filtros e campos que devem ser analisados sem mexer na lógica do programa.

O login/reutilização de sessão segue o padrão do projeto através de
`sap_session.py` e das variáveis já existentes no `.env`.

Execução típica no terminal do VS Code:
    .venv\Scripts\python.exe "Relatórios\Analise_Configuracao_SAP.py"

IMPORTANTE:
- Não grava nem altera dados no SAP.
- Por defeito usa SE16H apenas para leitura.
- Deixa os parâmetros variáveis concentrados na secção CONFIGURAÇÃO abaixo.
"""
from __future__ import annotations

import csv
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any


# =============================================================================
# CONFIGURAÇÃO — ALTERAR AQUI NO VS CODE
# =============================================================================

# Sistema/mandante já configurado no .env do projeto.
SAP_KEY = "S4DCLNT100"

# SE16H é a opção preferida para consultas. Se necessário pode mudar para SE16N.
TRANSACTION = "SE16H"

# Abre a consulta num novo modo SAP para não retirar o utilizador da transação
# em que está a trabalhar. Se o limite de modos estiver atingido, usa a sessão
# atual como fallback.
ABRIR_NOVO_MODO = True

# Fechar automaticamente o modo criado por este script no final.
FECHAR_MODO_NO_FIM = False

# Máximo de linhas a pedir e a devolver por tabela.
MAX_ROWS = 200

# Geração de ficheiros com o resultado em /cache.
GERAR_JSON = True
GERAR_CSV = False

# -----------------------------------------------------------------------------
# CONSULTAS
# -----------------------------------------------------------------------------
# Cada bloco aceita:
#   nome           : título que aparece no relatório
#   tabela         : tabela/view SAP
#   filtros        : lista de filtros. EQ é o mais seguro no SAP GUI.
#                    Ex.: {"campo": "LAND1", "valor": "PT", "opcao": "EQ"}
#   campos_saida   : [] = devolver todas as colunas encontradas no ALV.
#                    Ou informe apenas os campos desejados.
#
# Para analisar outra configuração, normalmente só é necessário alterar esta
# lista. A lógica abaixo não precisa ser modificada.
# -----------------------------------------------------------------------------
CONSULTAS = [
    {
        "nome": "Métodos de pagamento por país - Portugal",
        "tabela": "T042Z",
        "filtros": [
            {"campo": "LAND1", "valor": "PT", "opcao": "EQ"},
        ],
        "campos_saida": [],
    },

    # EXEMPLO — copie/descomente e altere quando quiser analisar outra tabela:
    # {
    #     "nome": "Exemplo de consulta",
    #     "tabela": "T001",
    #     "filtros": [
    #         {"campo": "BUKRS", "valor": "2100", "opcao": "EQ"},
    #     ],
    #     "campos_saida": ["BUKRS", "BUTXT", "LAND1", "WAERS"],
    # },
]


# =============================================================================
# FIM DA CONFIGURAÇÃO
# =============================================================================

ROOT_DIR = Path(__file__).resolve().parents[1]
CACHE_DIR = ROOT_DIR / "cache"

if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from sap_session import (  # noqa: E402
    ensure_sap_access_from_env,
    load_dotenv_manual,
    session_info,
)


if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# =============================================================================
# Utilitários SAP GUI
# =============================================================================

def _wait_not_busy(session, timeout_s: float = 20.0) -> None:
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            if not session.Busy:
                return
        except Exception:
            return
        time.sleep(0.1)
    raise TimeoutError("Timeout a aguardar o SAP terminar o processamento.")


def _status_bar(session) -> tuple[str, str]:
    try:
        sbar = session.findById("wnd[0]/sbar")
        return (
            str(getattr(sbar, "MessageType", "") or "").strip().upper(),
            str(getattr(sbar, "Text", "") or "").strip(),
        )
    except Exception:
        return "", ""


def _raise_if_sap_error(session, context: str) -> None:
    msg_type, msg = _status_bar(session)
    if msg_type in {"E", "A"} and msg:
        raise RuntimeError(f"{context}: {msg}")


def _set_text(session, candidates: list[str], value: str) -> str | None:
    for element_id in candidates:
        try:
            element = session.findById(element_id)
            element.Text = str(value)
            return element_id
        except Exception:
            continue
    return None


def _walk_children(root, max_nodes: int = 12000):
    stack = [root]
    seen = 0
    while stack and seen < max_nodes:
        obj = stack.pop()
        seen += 1
        yield obj
        try:
            count = obj.Children.Count
        except Exception:
            continue
        for idx in range(count - 1, -1, -1):
            try:
                stack.append(obj.Children(idx))
            except Exception:
                continue


def _open_transaction(session, transaction: str) -> None:
    okcd = session.findById("wnd[0]/tbar[0]/okcd")
    okcd.Text = f"/n{transaction.upper()}"
    session.findById("wnd[0]").sendVKey(0)
    _wait_not_busy(session)
    time.sleep(0.4)
    _raise_if_sap_error(session, f"Não foi possível abrir {transaction}")


def _open_transaction_new_mode(session, transaction: str):
    """Abre /oTRANSACTION e devolve a nova sessão. Fallback: sessão atual."""
    try:
        connection = session.Parent
        before_ids = set()
        for idx in range(connection.Children.Count):
            try:
                before_ids.add(str(connection.Children(idx).Id))
            except Exception:
                pass

        okcd = session.findById("wnd[0]/tbar[0]/okcd")
        okcd.Text = f"/o{transaction.upper()}"
        session.findById("wnd[0]").sendVKey(0)

        deadline = time.time() + 12
        while time.time() < deadline:
            for idx in range(connection.Children.Count):
                candidate = connection.Children(idx)
                try:
                    candidate_id = str(candidate.Id)
                except Exception:
                    candidate_id = ""
                if candidate_id and candidate_id not in before_ids:
                    _wait_not_busy(candidate)
                    return candidate
            time.sleep(0.2)
    except Exception:
        pass

    print("⚠️  Não foi possível abrir novo modo; a sessão atual será utilizada.")
    _open_transaction(session, transaction)
    return session


def _open_analysis_session(base_session):
    if ABRIR_NOVO_MODO:
        return _open_transaction_new_mode(base_session, TRANSACTION)
    _open_transaction(base_session, TRANSACTION)
    return base_session


def _close_session_window(session) -> None:
    try:
        session.findById("wnd[0]").close()
        time.sleep(0.3)
    except Exception:
        return
    for button in (
        "wnd[1]/usr/btnSPOP-OPTION1",
        "wnd[1]/tbar[0]/btn[0]",
    ):
        try:
            session.findById(button).press()
            return
        except Exception:
            continue


def _set_table_name(session, table: str) -> None:
    field_id = _set_text(
        session,
        [
            "wnd[0]/usr/ctxtGD-TAB",
            "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
            "wnd[0]/usr/ctxtTABNAME",
        ],
        table.upper(),
    )
    if not field_id:
        raise RuntimeError(
            f"Não encontrei o campo de tabela na {TRANSACTION}. "
            "Confirme que a transação abriu corretamente."
        )

    session.findById("wnd[0]").sendVKey(0)
    _wait_not_busy(session)
    time.sleep(0.5)
    _raise_if_sap_error(session, f"Tabela {table} não pôde ser carregada")


def _set_max_rows(session, max_rows: int) -> None:
    candidates = [
        "wnd[0]/usr/txtMAX_SEL",
        "wnd[0]/usr/txtGD-MAXROWS",
        "wnd[0]/usr/txtGD-MAX_LINES",
        "wnd[0]/usr/txtMAX_HITS",
    ]
    _set_text(session, candidates, str(max_rows))


def _find_selection_table(session):
    roots = []
    for root_id in ("wnd[0]/usr", "wnd[0]"):
        try:
            roots.append(session.findById(root_id))
        except Exception:
            pass

    for root in roots:
        for obj in _walk_children(root):
            try:
                obj_id = str(obj.Id or "")
                name = str(obj.Name or "")
                type_name = str(obj.Type or "")
            except Exception:
                continue

            marker = f"{obj_id} {name}".upper()
            if "SAPLSE16NSELFIELDS_TC" in marker:
                return obj
            if "GUITABLECONTROL" in type_name.upper() and "SELFIELDS" in marker:
                return obj

    return None


def _parse_cell_coordinates(element_id: str) -> tuple[int, int] | None:
    if "[" not in element_id or "]" not in element_id:
        return None
    try:
        payload = element_id.rsplit("[", 1)[1].split("]", 1)[0]
        col_raw, row_raw = payload.split(",", 1)
        return int(col_raw), int(row_raw)
    except Exception:
        return None


def _selection_columns(table_control) -> dict[str, tuple[int, str]]:
    """Descobre as colunas técnicas FIELDNAME / LOW / OPTION do table control."""
    found: dict[str, tuple[int, str]] = {}

    try:
        children_count = table_control.Children.Count
    except Exception:
        return found

    for idx in range(children_count):
        try:
            child = table_control.Children(idx)
            child_id = str(child.Id)
            coords = _parse_cell_coordinates(child_id)
            if not coords:
                continue
            col, row = coords
            if row != 0:
                continue
            name = str(getattr(child, "Name", "") or "").upper()
            prefix = child_id.rsplit("[", 1)[0].split("/")[-1]
        except Exception:
            continue

        if "FIELDNAME" in name:
            found["FIELDNAME"] = (col, prefix)
        elif name.endswith("-LOW") or name == "GS_SELFIELDS-LOW":
            found["LOW"] = (col, prefix)
        elif name.endswith("-OPTION") or name in {"GS_SELFIELDS-OPTION", "OPTION"}:
            found["OPTION"] = (col, prefix)
        elif name.endswith("-HIGH") or name == "GS_SELFIELDS-HIGH":
            found["HIGH"] = (col, prefix)

    # Fallbacks conhecidos da família SE16N/SE16H.
    found.setdefault("FIELDNAME", (13, "txtGS_SELFIELDS-FIELDNAME"))
    found.setdefault("LOW", (2, "ctxtGS_SELFIELDS-LOW"))
    return found


def _find_field_visible_row(session, table_control, field_name: str) -> tuple[int, dict] | None:
    """Percorre o table control e devolve a linha VISÍVEL do campo técnico."""
    field_name = field_name.strip().upper()
    columns = _selection_columns(table_control)
    field_col, field_prefix = columns["FIELDNAME"]

    try:
        row_count = int(table_control.RowCount)
        visible_count = max(1, int(table_control.VisibleRowCount))
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler os campos de seleção: {exc}") from exc

    positions = list(range(0, max(row_count, 1), visible_count))
    if positions[-1] != max(0, row_count - visible_count):
        positions.append(max(0, row_count - visible_count))

    for position in sorted(set(positions)):
        try:
            table_control.VerticalScrollbar.Position = position
            _wait_not_busy(session, timeout_s=5)
            time.sleep(0.1)
        except Exception:
            pass

        for visible_row in range(min(visible_count, row_count)):
            field_id = (
                f"{table_control.Id}/{field_prefix}"
                f"[{field_col},{visible_row}]"
            )
            try:
                current = str(session.findById(field_id).Text or "").strip().upper()
            except Exception:
                continue
            if current == field_name:
                return visible_row, columns

    return None


def _set_selection_cell(session, table_control, prefix: str, col: int, row: int, value: str) -> bool:
    element_id = f"{table_control.Id}/{prefix}[{col},{row}]"
    try:
        session.findById(element_id).Text = str(value)
        return True
    except Exception:
        return False


def _apply_filter(session, table_control, filter_cfg: dict[str, Any]) -> None:
    field = str(filter_cfg.get("campo") or "").strip().upper()
    value = str(filter_cfg.get("valor") or "").strip()
    option = str(filter_cfg.get("opcao") or "EQ").strip().upper()
    high = str(filter_cfg.get("high") or "").strip()

    if not field:
        return

    located = _find_field_visible_row(session, table_control, field)
    if not located:
        raise RuntimeError(f"Campo de filtro '{field}' não encontrado na tabela.")

    visible_row, columns = located
    low_col, low_prefix = columns["LOW"]
    if not _set_selection_cell(session, table_control, low_prefix, low_col, visible_row, value):
        raise RuntimeError(f"Não foi possível preencher {field}={value}.")

    # EQ é o default SAP e não precisa ser alterado. Para outras opções tentamos
    # preencher a coluna OPTION se ela for editável nesta versão da SE16H/SE16N.
    if option and option != "EQ" and "OPTION" in columns:
        option_col, option_prefix = columns["OPTION"]
        if not option_prefix.lower().startswith("btn"):
            if not _set_selection_cell(
                session, table_control, option_prefix, option_col, visible_row, option
            ):
                raise RuntimeError(
                    f"Não foi possível definir a opção {option} para o campo {field}."
                )

    if high and "HIGH" in columns:
        high_col, high_prefix = columns["HIGH"]
        _set_selection_cell(session, table_control, high_prefix, high_col, visible_row, high)


def _apply_filters(session, filters: list[dict[str, Any]]) -> None:
    if not filters:
        return

    table_control = _find_selection_table(session)
    if table_control is None:
        raise RuntimeError(
            "Não foi encontrado o controlo de campos de seleção da SE16H/SE16N."
        )

    for filter_cfg in filters:
        _apply_filter(session, table_control, filter_cfg)


def _execute_query(session) -> None:
    # F8 — funciona em SE16H/SE16N.
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    except Exception:
        session.findById("wnd[0]").sendVKey(8)
    _wait_not_busy(session, timeout_s=30)
    time.sleep(0.7)
    _raise_if_sap_error(session, "Erro ao executar a consulta")


def _score_grid_candidate(obj) -> int:
    try:
        rows = int(obj.RowCount)
        cols = int(obj.ColumnCount)
    except Exception:
        return -1
    if rows < 0 or cols <= 0:
        return -1

    score = 10
    if rows > 0:
        score += 10
    try:
        _ = obj.GetColumnKey(0)
        score += 20
    except Exception:
        pass
    try:
        if rows:
            first_key = str(obj.GetColumnKey(0))
            _ = obj.GetCellValue(0, first_key)
            score += 20
    except Exception:
        pass
    return score


def _find_best_result_grid(session):
    direct_candidates = [
        "wnd[0]/usr/cntlRESULT/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlALV_GRID/shellcont/shell",
        "wnd[0]/usr/shellcont/shell",
    ]
    for element_id in direct_candidates:
        try:
            obj = session.findById(element_id)
            if _score_grid_candidate(obj) >= 0:
                return obj
        except Exception:
            continue

    roots = []
    for root_id in ("wnd[0]/usr", "wnd[0]"):
        try:
            roots.append(session.findById(root_id))
        except Exception:
            pass

    best = None
    best_score = -1
    for root in roots:
        for obj in _walk_children(root):
            score = _score_grid_candidate(obj)
            if score > best_score:
                best = obj
                best_score = score
    return best


def _column_keys(grid) -> list[str]:
    keys: list[str] = []
    try:
        col_count = int(grid.ColumnCount)
    except Exception:
        return keys

    for idx in range(col_count):
        try:
            key = str(grid.GetColumnKey(idx) or "").strip()
        except Exception:
            key = ""
        if key and key not in keys:
            keys.append(key)
    return keys


def _read_result_grid(
    session,
    max_rows: int,
    output_fields: list[str] | None = None,
) -> list[dict[str, str]]:
    grid = _find_best_result_grid(session)
    if grid is None:
        msg_type, msg = _status_bar(session)
        if msg:
            print(f"ℹ️  STATUS SAP: {msg_type} {msg}")
        return []

    keys = _column_keys(grid)
    requested = [str(x).strip().upper() for x in (output_fields or []) if str(x).strip()]
    if requested:
        selected_keys = [key for key in keys if key.upper() in requested]
        missing = [field for field in requested if field not in {k.upper() for k in keys}]
        if missing:
            print(f"⚠️  Campos de saída não encontrados no ALV: {', '.join(missing)}")
    else:
        selected_keys = keys

    try:
        row_count = min(int(grid.RowCount), int(max_rows))
    except Exception:
        row_count = 0

    rows: list[dict[str, str]] = []
    for row_idx in range(row_count):
        row: dict[str, str] = {}
        for key in selected_keys:
            try:
                row[key] = str(grid.GetCellValue(row_idx, key) or "").strip()
            except Exception:
                row[key] = ""
        rows.append(row)
    return rows


# =============================================================================
# Relatório
# =============================================================================

def _print_query_header(cfg: dict[str, Any]) -> None:
    print("\n" + "=" * 100)
    print(f"🔎 {cfg.get('nome') or cfg.get('tabela')}")
    print(f"Tabela: {cfg.get('tabela')}")
    filters = cfg.get("filtros") or []
    if filters:
        print("Filtros:")
        for item in filters:
            print(
                f"  - {item.get('campo')} {item.get('opcao', 'EQ')} "
                f"{item.get('valor', '')}"
            )
    else:
        print("Filtros: (sem filtros)")
    print("=" * 100)


def _print_rows(rows: list[dict[str, str]]) -> None:
    if not rows:
        print("📭 Nenhum registo devolvido.")
        return

    headers = list(rows[0].keys())
    print(f"✅ {len(rows)} registo(s) devolvido(s).")
    print(" | ".join(headers))
    print("-" * min(180, max(80, len(" | ".join(headers)))))
    for idx, row in enumerate(rows, start=1):
        values = [str(row.get(header, "")) for header in headers]
        print(f"{idx:>4} | " + " | ".join(values))


def _save_json(payload: dict[str, Any]) -> Path:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = CACHE_DIR / f"analise_configuracao_sap_{timestamp}.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def _save_csvs(results: list[dict[str, Any]]) -> list[Path]:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    paths: list[Path] = []

    for idx, result in enumerate(results, start=1):
        rows = result.get("rows") or []
        if not rows:
            continue
        table = str(result.get("table") or f"TABLE_{idx}").strip().upper()
        output = CACHE_DIR / f"analise_{table}_{timestamp}.csv"
        with output.open("w", newline="", encoding="utf-8-sig") as file_obj:
            writer = csv.DictWriter(file_obj, fieldnames=list(rows[0].keys()), delimiter=";")
            writer.writeheader()
            writer.writerows(rows)
        paths.append(output)

    return paths


# =============================================================================
# Execução
# =============================================================================

def _run_single_query(session, cfg: dict[str, Any]) -> dict[str, Any]:
    table = str(cfg.get("tabela") or "").strip().upper()
    if not table:
        raise ValueError("Existe uma consulta sem nome de tabela configurado.")

    _print_query_header(cfg)
    _open_transaction(session, TRANSACTION)
    _set_table_name(session, table)
    _set_max_rows(session, MAX_ROWS)
    _apply_filters(session, cfg.get("filtros") or [])
    _execute_query(session)
    rows = _read_result_grid(
        session,
        MAX_ROWS,
        output_fields=cfg.get("campos_saida") or [],
    )
    _print_rows(rows)

    return {
        "name": cfg.get("nome") or table,
        "table": table,
        "filters": cfg.get("filtros") or [],
        "requested_fields": cfg.get("campos_saida") or [],
        "row_count": len(rows),
        "rows": rows,
    }


def main() -> int:
    print("SAP - Análise Dinâmica de Configuração")
    print(f"Transação de leitura : {TRANSACTION}")
    print(f"SAP KEY             : {SAP_KEY}")
    print(f"Máximo de linhas    : {MAX_ROWS}")
    print(f"Consultas           : {len(CONSULTAS)}")

    if not CONSULTAS:
        print("❌ A lista CONSULTAS está vazia. Configure pelo menos uma tabela.")
        return 2

    load_dotenv_manual()
    base_session = ensure_sap_access_from_env(
        key=SAP_KEY,
        timeout_s=40,
        load_env=True,
    )
    info = session_info(base_session)
    print(
        "✅ SAP ligado | "
        f"Sistema={info['system_name']} | Cliente={info['client']} | User={info['user']}"
    )

    analysis_session = _open_analysis_session(base_session)
    results: list[dict[str, Any]] = []

    try:
        for cfg in CONSULTAS:
            try:
                results.append(_run_single_query(analysis_session, cfg))
            except Exception as exc:
                table = str(cfg.get("tabela") or "").upper()
                print(f"❌ Erro na consulta {table}: {exc}")
                results.append(
                    {
                        "name": cfg.get("nome") or table,
                        "table": table,
                        "filters": cfg.get("filtros") or [],
                        "requested_fields": cfg.get("campos_saida") or [],
                        "row_count": 0,
                        "rows": [],
                        "error": str(exc),
                    }
                )
    finally:
        if FECHAR_MODO_NO_FIM and analysis_session is not base_session:
            _close_session_window(analysis_session)

    payload = {
        "meta": {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "system": info.get("system_name", ""),
            "client": info.get("client", ""),
            "user": info.get("user", ""),
            "transaction": TRANSACTION,
            "max_rows": MAX_ROWS,
        },
        "results": results,
    }

    if GERAR_JSON:
        json_path = _save_json(payload)
        print(f"\n💾 JSON: {json_path}")

    if GERAR_CSV:
        for csv_path in _save_csvs(results):
            print(f"💾 CSV : {csv_path}")

    errors = [item for item in results if item.get("error")]
    print("\n" + "=" * 100)
    if errors:
        print(f"⚠️  Análise concluída com {len(errors)} erro(s).")
        return 1
    print("✅ Análise concluída sem erros.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

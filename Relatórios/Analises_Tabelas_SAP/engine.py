# -*- coding: utf-8 -*-
"""Motor genérico e SOMENTE DE LEITURA para análise de tabelas SAP.

Este ficheiro não deve conter parâmetros de um processo específico.
Cada análise fica em ``processos/<nome>.py`` e apenas descreve tabelas,
filtros, campos de saída e parâmetros de execução.
"""
from __future__ import annotations

import csv
import json
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT_DIR = Path(__file__).resolve().parents[2]
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


@dataclass(frozen=True)
class RuntimeConfig:
    processo_id: str
    titulo: str
    sap_key: str
    transaction: str
    abrir_novo_modo: bool
    fechar_modo_no_fim: bool
    max_rows: int
    gerar_json: bool
    gerar_csv: bool
    consultas: list[dict[str, Any]]
    metodo: str = "GUI"


def _to_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on", "sim", "s"}


def normalizar_processo(config: dict[str, Any]) -> RuntimeConfig:
    processo_id = str(config.get("id") or "analise_sap").strip()
    titulo = str(config.get("titulo") or processo_id).strip()
    sap_key = str(config.get("sap_key") or "S4DCLNT100").strip().upper()
    transaction = str(config.get("transaction") or "SE16H").strip().upper()
    consultas = list(config.get("consultas") or [])

    metodo_raw = str(config.get("metodo") or "").strip().upper()
    if not metodo_raw:
        metodo_raw = "RFC" if _to_bool(config.get("use_rfc")) else "GUI"
    metodo = "RFC" if metodo_raw == "RFC" else "GUI"

    if transaction not in {"SE16H", "SE16N"}:
        raise ValueError("A transação de leitura suportada deve ser SE16H ou SE16N.")
    if not consultas:
        raise ValueError("O processo não possui consultas configuradas.")

    for idx, consulta in enumerate(consultas, start=1):
        if not str(consulta.get("tabela") or "").strip():
            raise ValueError(f"Consulta #{idx} sem tabela configurada.")

    return RuntimeConfig(
        processo_id=processo_id,
        titulo=titulo,
        sap_key=sap_key,
        transaction=transaction,
        abrir_novo_modo=_to_bool(config.get("abrir_novo_modo"), True),
        fechar_modo_no_fim=_to_bool(config.get("fechar_modo_no_fim"), False),
        max_rows=max(1, int(config.get("max_rows") or 200)),
        gerar_json=_to_bool(config.get("gerar_json"), True),
        gerar_csv=_to_bool(config.get("gerar_csv"), False),
        consultas=consultas,
        metodo=metodo,
    )



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
        before_ids: set[str] = set()
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


def _open_analysis_session(base_session, runtime: RuntimeConfig):
    if runtime.abrir_novo_modo:
        return _open_transaction_new_mode(base_session, runtime.transaction)
    _open_transaction(base_session, runtime.transaction)
    return base_session


def _close_session_window(session) -> None:
    try:
        session.findById("wnd[0]").close()
        time.sleep(0.3)
    except Exception:
        return
    for button in ("wnd[1]/usr/btnSPOP-OPTION1", "wnd[1]/tbar[0]/btn[0]"):
        try:
            session.findById(button).press()
            return
        except Exception:
            continue


def _set_table_name(session, table: str, transaction: str) -> None:
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
            f"Não encontrei o campo de tabela na {transaction}. "
            "Confirme que a transação abriu corretamente."
        )

    session.findById("wnd[0]").sendVKey(0)
    _wait_not_busy(session)
    time.sleep(0.5)
    _raise_if_sap_error(session, f"Tabela {table} não pôde ser carregada")


def _set_max_rows(session, max_rows: int) -> None:
    _set_text(
        session,
        [
            "wnd[0]/usr/txtMAX_SEL",
            "wnd[0]/usr/txtGD-MAXROWS",
            "wnd[0]/usr/txtGD-MAX_LINES",
            "wnd[0]/usr/txtMAX_HITS",
        ],
        str(max_rows),
    )


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

    found.setdefault("FIELDNAME", (13, "txtGS_SELFIELDS-FIELDNAME"))
    found.setdefault("LOW", (2, "ctxtGS_SELFIELDS-LOW"))
    return found


def _find_field_visible_row(session, table_control, field_name: str) -> tuple[int, dict] | None:
    field_name = field_name.strip().upper()
    columns = _selection_columns(table_control)
    field_col, field_prefix = columns["FIELDNAME"]

    try:
        row_count = int(table_control.RowCount)
        visible_count = max(1, int(table_control.VisibleRowCount))
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler os campos de seleção: {exc}") from exc

    positions = list(range(0, max(row_count, 1), visible_count))
    last_position = max(0, row_count - visible_count)
    if positions and positions[-1] != last_position:
        positions.append(last_position)

    for position in sorted(set(positions or [0])):
        try:
            table_control.VerticalScrollbar.Position = position
            _wait_not_busy(session, timeout_s=5)
            time.sleep(0.1)
        except Exception:
            pass

        for visible_row in range(min(visible_count, row_count)):
            field_id = f"{table_control.Id}/{field_prefix}[{field_col},{visible_row}]"
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

    if option and option != "EQ" and "OPTION" in columns:
        option_col, option_prefix = columns["OPTION"]
        if not option_prefix.lower().startswith("btn"):
            if not _set_selection_cell(session, table_control, option_prefix, option_col, visible_row, option):
                raise RuntimeError(f"Não foi possível definir a opção {option} para {field}.")

    if high and "HIGH" in columns:
        high_col, high_prefix = columns["HIGH"]
        _set_selection_cell(session, table_control, high_prefix, high_col, visible_row, high)


def _apply_filters(session, filters: list[dict[str, Any]]) -> None:
    if not filters:
        return
    table_control = _find_selection_table(session)
    if table_control is None:
        raise RuntimeError("Não foi encontrado o controlo de seleção da SE16H/SE16N.")
    for filter_cfg in filters:
        _apply_filter(session, table_control, filter_cfg)


def _execute_query(session) -> None:
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
    score = 10 + (10 if rows > 0 else 0)
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
    for element_id in (
        "wnd[0]/usr/cntlRESULT/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlALV_GRID/shellcont/shell",
        "wnd[0]/usr/shellcont/shell",
    ):
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


def _read_result_grid(session, max_rows: int, output_fields: list[str] | None = None) -> list[dict[str, str]]:
    grid = _find_best_result_grid(session)
    if grid is None:
        msg_type, msg = _status_bar(session)
        if msg:
            print(f"ℹ️  STATUS SAP: {msg_type} {msg}")
        return []

    keys = _column_keys(grid)
    requested = [str(x).strip().upper() for x in (output_fields or []) if str(x).strip()]
    if requested:
        available_upper = {key.upper() for key in keys}
        selected_keys = [key for key in keys if key.upper() in requested]
        missing = [field for field in requested if field not in available_upper]
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


def _print_query_header(cfg: dict[str, Any]) -> None:
    print("\n" + "=" * 100)
    print(f"🔎 {cfg.get('nome') or cfg.get('tabela')}")
    print(f"Tabela: {cfg.get('tabela')}")
    filters = cfg.get("filtros") or []
    if filters:
        print("Filtros:")
        for item in filters:
            print(f"  - {item.get('campo')} {item.get('opcao', 'EQ')} {item.get('valor', '')}")
    else:
        print("Filtros: (sem filtros)")
    campos = cfg.get("campos_saida") or []
    print("Campos de saída: " + (", ".join(campos) if campos else "todos os campos do ALV"))
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
        print(f"{idx:>4} | " + " | ".join(str(row.get(header, "")) for header in headers))


def _cache_dir(runtime: RuntimeConfig) -> Path:
    return ROOT_DIR / "cache" / "analises_tabelas_sap" / runtime.processo_id


def _save_json(payload: dict[str, Any], runtime: RuntimeConfig) -> Path:
    output_dir = _cache_dir(runtime)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = output_dir / f"{runtime.processo_id}_{timestamp}.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output


def _save_csvs(results: list[dict[str, Any]], runtime: RuntimeConfig) -> list[Path]:
    output_dir = _cache_dir(runtime)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    paths: list[Path] = []
    for idx, result in enumerate(results, start=1):
        rows = result.get("rows") or []
        if not rows:
            continue
        table = str(result.get("table") or f"TABLE_{idx}").strip().upper()
        output = output_dir / f"{runtime.processo_id}_{table}_{timestamp}.csv"
        with output.open("w", newline="", encoding="utf-8-sig") as file_obj:
            writer = csv.DictWriter(file_obj, fieldnames=list(rows[0].keys()), delimiter=";")
            writer.writeheader()
            writer.writerows(rows)
        paths.append(output)
    return paths


def _run_single_query(session, cfg: dict[str, Any], runtime: RuntimeConfig) -> dict[str, Any]:
    table = str(cfg.get("tabela") or "").strip().upper()
    _print_query_header(cfg)
    _open_transaction(session, runtime.transaction)
    _set_table_name(session, table, runtime.transaction)
    _set_max_rows(session, int(cfg.get("max_rows") or runtime.max_rows))
    _apply_filters(session, cfg.get("filtros") or [])
    _execute_query(session)
    rows = _read_result_grid(
        session,
        int(cfg.get("max_rows") or runtime.max_rows),
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


def _obter_conexao_rfc(sap_key: str) -> tuple[Any, dict[str, str]]:
    import os
    try:
        from pyrfc import Connection
    except ImportError as exc:
        raise RuntimeError("A biblioteca 'pyrfc' não está instalada neste ambiente Python.") from exc

    key_upper = str(sap_key or "S4DCLNT100").upper().strip()

    alias = "DEV"
    if "QAD" in key_upper or "S4Q" in key_upper:
        alias = "QAD"
    elif "PRD" in key_upper or "S4P" in key_upper:
        alias = "PRD"

    ashost = os.getenv(f"SAP_ASHOST_{alias}") or os.getenv("SAP_ASHOST", "")
    sysnr = os.getenv(f"SAP_SYSNR_{alias}") or os.getenv("SAP_SYSNR", "00")
    client = os.getenv(f"SAP_CLIENT_{key_upper}") or os.getenv(f"SAP_CLIENT_{alias}") or os.getenv("SAP_CLIENT", "100")
    user = os.getenv(f"SAP_USER_{alias}") or os.getenv("SAP_USER", "")
    passwd = os.getenv(f"SAP_PASSWORD_{key_upper}") or os.getenv(f"SAP_PASSWORD_{alias}") or os.getenv("SAP_PASSWORD", "")
    lang = os.getenv(f"SAP_LANGUAGE_{alias}") or os.getenv("SAP_LANGUAGE", "PT")

    if not ashost:
        raise RuntimeError(f"Falta definir SAP_ASHOST_{alias} no ficheiro .env.")
    if not passwd:
        raise RuntimeError(f"Falta definir a password RFC para {key_upper} / {alias} no .env.")

    conn = Connection(
        ashost=ashost,
        sysnr=sysnr,
        client=client,
        user=user,
        passwd=passwd,
        lang=lang,
    )
    info = {
        "system_name": alias,
        "client": client,
        "user": user,
    }
    return conn, info


def _run_single_query_rfc(conn: Any, cfg: dict[str, Any], runtime: RuntimeConfig) -> dict[str, Any]:
    table = str(cfg.get("tabela") or "").strip().upper()
    _print_query_header(cfg)

    options: list[dict[str, str]] = []
    filtros = cfg.get("filtros") or []
    for f in filtros:
        campo = str(f.get("campo") or "").strip().upper()
        valor = f.get("valor")
        opcao = str(f.get("opcao") or "EQ").strip().upper()
        if not campo or valor is None:
            continue
        if opcao == "EQ":
            clause = f"{campo} = '{valor}'"
        elif opcao == "NE":
            clause = f"{campo} <> '{valor}'"
        elif opcao == "LIKE":
            clause = f"{campo} LIKE '{valor}'"
        elif opcao == "IN":
            if isinstance(valor, (list, tuple, set)):
                vals_str = "', '".join(str(v) for v in valor)
                clause = f"{campo} IN ('{vals_str}')"
            else:
                clause = f"{campo} = '{valor}'"
        else:
            clause = f"{campo} {opcao} '{valor}'"

        options.append({"TEXT": clause})

    requested_fields = cfg.get("campos_saida") or []
    fields_input = [{"FIELDNAME": f.strip().upper()} for f in requested_fields if f.strip()]

    max_rows = int(cfg.get("max_rows") or runtime.max_rows)

    res = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE=table,
        OPTIONS=options,
        FIELDS=fields_input,
        DELIMITER="|",
        ROWCOUNT=max_rows,
    )

    fields_meta = res.get("FIELDS") or []
    headers = [f["FIELDNAME"].strip() for f in fields_meta]
    raw_data = res.get("DATA") or []

    rows: list[dict[str, str]] = []
    for item in raw_data:
        wa = str(item.get("WA") or "")
        parts = wa.split("|")
        row_dict: dict[str, str] = {}
        for idx_h, h in enumerate(headers):
            val = parts[idx_h].strip() if idx_h < len(parts) else ""
            row_dict[h] = val
        rows.append(row_dict)

    _print_rows(rows)

    return {
        "name": cfg.get("nome") or table,
        "table": table,
        "filters": cfg.get("filtros") or [],
        "requested_fields": requested_fields,
        "row_count": len(rows),
        "rows": rows,
    }


def executar_processo(config: dict[str, Any]) -> int:
    """Executa um processo declarativo definido em ``processos/*.py``."""
    runtime = normalizar_processo(config)
    print(f"SAP - {runtime.titulo}")
    print(f"Processo            : {runtime.processo_id}")
    print(f"Método              : {runtime.metodo}")
    if runtime.metodo == "GUI":
        print(f"Transação de leitura : {runtime.transaction}")
    print(f"SAP KEY             : {runtime.sap_key}")
    print(f"Máximo de linhas    : {runtime.max_rows}")
    print(f"Consultas           : {len(runtime.consultas)}")

    load_dotenv_manual()
    results: list[dict[str, Any]] = []

    if runtime.metodo == "RFC":
        try:
            conn, info = _obter_conexao_rfc(runtime.sap_key)
            print(
                "✅ SAP ligado (RFC) | "
                f"Sistema={info['system_name']} | Cliente={info['client']} | User={info['user']}"
            )
        except Exception as exc_rfc:
            print(f"❌ Erro ao ligar ao SAP via RFC: {exc_rfc}")
            return 1

        for cfg in runtime.consultas:
            try:
                results.append(_run_single_query_rfc(conn, cfg, runtime))
            except Exception as exc:
                table = str(cfg.get("tabela") or "").upper()
                print(f"❌ Erro na consulta RFC {table}: {exc}")
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
        try:
            conn.close()
        except Exception:
            pass
    else:
        base_session = ensure_sap_access_from_env(
            key=runtime.sap_key,
            timeout_s=40,
            load_env=True,
        )
        info = session_info(base_session)
        print(
            "✅ SAP ligado (GUI) | "
            f"Sistema={info['system_name']} | Cliente={info['client']} | User={info['user']}"
        )

        analysis_session = _open_analysis_session(base_session, runtime)
        try:
            for cfg in runtime.consultas:
                try:
                    results.append(_run_single_query(analysis_session, cfg, runtime))
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
            if runtime.fechar_modo_no_fim and analysis_session is not base_session:
                _close_session_window(analysis_session)

    payload = {
        "meta": {
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "processo": runtime.processo_id,
            "titulo": runtime.titulo,
            "metodo": runtime.metodo,
            "system": info.get("system_name", ""),
            "client": info.get("client", ""),
            "user": info.get("user", ""),
            "transaction": runtime.transaction if runtime.metodo == "GUI" else "RFC_READ_TABLE",
            "max_rows": runtime.max_rows,
        },
        "results": results,
    }

    if runtime.gerar_json:
        print(f"\n💾 JSON: {_save_json(payload, runtime)}")
    if runtime.gerar_csv:
        for csv_path in _save_csvs(results, runtime):
            print(f"💾 CSV : {csv_path}")

    errors = [item for item in results if item.get("error")]
    print("\n" + "=" * 100)
    if errors:
        print(f"⚠️  Análise concluída com {len(errors)} erro(s).")
        return 1
    print("✅ Análise concluída sem erros.")
    return 0


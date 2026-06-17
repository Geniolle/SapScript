# -*- coding: utf-8 -*-
"""
pesquisar_request.py

Objetivo:
- Abrir SE16H em NOVO modo (/ose16h)
- Minimizar a janela desse novo modo enquanto executa
- Ler resultados da E070 e imprimir lista NUMERADA:
    N | TRKORR | AS4TEXT
- Regra atualizada: APENAS listar as linhas cujo valor da coluna STRKORR for diferente de vazio.
- Guardar automaticamente a lista num ficheiro JSON para uso posterior (seleção por número da linha)

Execução:
& "C:/SAP Script/.venv/Scripts/python.exe" "C:/SAP Script/Processos/pesquisar_request.py" --system "S4Q" --max "5000"

Opcional:
--no-new-mode        (não usa /o; usa a sessão atual)
--no-minimize        (não minimiza)
--no-close           (não fecha a janela ao terminar)
"""

import sys
import time
import json
import os
import win32com.client

# Garantir codificação UTF-8 para evitar erros UnicodeEncodeError em consolas Windows
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

MSG_RZ11_SCRIPTING = 'Ativar na transação RZ11 o nome do parametro "sapgui/user_scripting" alterar para "TRUE"'


def _base_dir():
    # este ficheiro está em ...\SAP Script\Processos\pesquisar_request.py
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def _cache_dir():
    d = os.path.join(_base_dir(), "cache")
    os.makedirs(d, exist_ok=True)
    return d


def _cache_file_path():
    return os.path.join(_cache_dir(), "last_e070_list.json")


def _save_results(results, system_name, user):
    """
    results: list[tuple(trkorr, as4text)]
    Guarda:
      - índice (1..N)
      - trkorr
      - as4text
      - metadata
    """
    payload = {
        "meta": {
            "system": system_name,
            "user": user,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        },
        "items": [
            {"idx": i + 1, "TRKORR": trkorr, "AS4TEXT": as4text}
            for i, (trkorr, as4text) in enumerate(results)
        ],
    }

    path = _cache_file_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    os.environ["SAP_LAST_E070_LIST_FILE"] = path
    os.environ["SAP_LAST_E070_LIST_COUNT"] = str(len(results))

    print(f"\n💾 Lista guardada em: {path}")
    return path


def _log_alerta_rz11():
    print(f"⚠️  {MSG_RZ11_SCRIPTING}")


def _erro_scripting_inativo(e=None):
    msg = "O scripting do SAP GUI não está ativo ou não foi possível inicializar o objeto SAPGUI. Ativar na transação RZ11 o parâmetro 'sapgui/user_scripting' para 'TRUE'."
    print(f"❌ {msg}")
    _log_alerta_rz11()
    if e:
        print(f"🔧 Detalhes técnicos: {e}")
        msg += f" Detalhes técnicos: {e}"
    raise RuntimeError(msg)


def _get_application():
    try:
        sap = win32com.client.GetObject("SAPGUI")
        app = sap.GetScriptingEngine
        if not app:
            raise RuntimeError("GetScriptingEngine retornou vazio/None.")
        return app
    except Exception as e:
        _erro_scripting_inativo(e)


def _iter_sessions(application):
    try:
        for i in range(application.Children.Count):
            conn = application.Children(i)
            try:
                for j in range(conn.Children.Count):
                    yield conn.Children(j)
            except Exception:
                continue
    except Exception:
        return


def _pick_session(application, system_name=None):
    candidates = []
    for sess in _iter_sessions(application):
        try:
            sysname = (sess.Info.SystemName or "").upper()
        except Exception:
            sysname = ""
        try:
            user = (sess.Info.User or "").strip()
        except Exception:
            user = ""
        candidates.append((sysname, bool(user), sess))

    if not candidates:
        msg = "Nenhuma sessão SAP ativa encontrada. Abra o SAP Logon e faça login."
        print(f"❌ {msg}")
        raise RuntimeError(msg)

    if system_name:
        target = system_name.upper()
        in_sys = [c for c in candidates if c[0] == target]
        if in_sys:
            logged = [c for c in in_sys if c[1]]
            return logged[0][2] if logged else in_sys[0][2]

    logged_any = [c for c in candidates if c[1]]
    return logged_any[0][2] if logged_any else candidates[0][2]


def _wait_not_busy(session, timeout_s=12):
    t0 = time.time()
    while time.time() - t0 <= timeout_s:
        try:
            if not session.Busy:
                return True
        except Exception:
            return True
        time.sleep(0.1)
    return False


def _try_set_text(session, id_path, value):
    try:
        session.findById(id_path).text = value
        return True
    except Exception:
        return False


def _try_press(session, id_path):
    try:
        session.findById(id_path).press()
        return True
    except Exception:
        return False


def _iconify(session):
    try:
        session.findById("wnd[0]").iconify()
        return True
    except Exception:
        return False


def _close_window(session):
    try:
        session.findById("wnd[0]").close()
    except Exception:
        return

    time.sleep(0.3)

    # confirmação padrão de popup (se existir)
    try:
        if _try_press(session, "wnd[1]/usr/btnSPOP-OPTION1"):
            return
        if _try_press(session, "wnd[1]/tbar[0]/btn[0]"):
            return
    except Exception:
        pass


def _wait_for_table_input_field(session, timeout_s=5):
    candidates = [
        "wnd[0]/usr/ctxtGD-TAB",
        "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
        "wnd[0]/usr/ctxtTABNAME",
    ]
    t0 = time.time()
    while time.time() - t0 <= timeout_s:
        for cid in candidates:
            try:
                element = session.findById(cid)
                if element is not None:
                    return cid
            except Exception:
                pass
        time.sleep(0.1)
    return None


def _set_table_e070(session):
    cid = _wait_for_table_input_field(session, 5)
    if cid:
        if _try_set_text(session, cid, "E070"):
            try:
                session.findById("wnd[0]").sendVKey(0)
            except Exception:
                pass
            _wait_not_busy(session, 10)
            time.sleep(0.2)
            return True
    return False


def _set_max_ocorrencias(session, max_rows="5000"):
    candidates = [
        "wnd[0]/usr/txtMAX_SEL",
        "wnd[0]/usr/txtGD-MAXROWS",
        "wnd[0]/usr/txtMAX_HITS",
    ]
    for cid in candidates:
        if _try_set_text(session, cid, str(max_rows)):
            return True
    return False


def _find_table_control(session):
    try:
        root = session.findById("wnd[0]/usr")
    except Exception:
        return None
    stack = [root]
    while stack:
        obj = stack.pop()
        try:
            # Check both Name and ID to be fully bulletproof (Name is often without the "tbl" prefix)
            if obj.Name == "SAPLSE16NSELFIELDS_TC" or obj.Id.endswith("tblSAPLSE16NSELFIELDS_TC"):
                return obj
        except Exception:
            pass
        try:
            cnt = obj.Children.Count
            for i in range(cnt):
                stack.append(obj.Children(i))
        except Exception:
            pass
    return None


def _wait_for_table_control(session, timeout_s=8):
    t0 = time.time()
    while time.time() - t0 <= timeout_s:
        tbl = _find_table_control(session)
        if tbl is not None:
            return tbl
        time.sleep(0.1)
    return None


def _set_low_value(session, tbl_id, prefix_path, col_idx, row_idx, value):
    target_id = f"{tbl_id}/{prefix_path}[{col_idx},{row_idx}]"
    return _try_set_text(session, target_id, value)


def _detect_columns(tbl):
    col_info = {
        "technical_name_col": 13,
        "technical_name_prefix": "txtGS_SELFIELDS-FIELDNAME",
        "low_col": 2,
        "low_prefix": "ctxtGS_SELFIELDS-LOW",
        "option_col": 4,
        "option_prefix": "txtGS_SELFIELDS-OPTION"
    }
    try:
        for i in range(tbl.Children.Count):
            child = tbl.Children(i)
            id_str = child.Id
            if "[" in id_str and "]" in id_str:
                bracket_part = id_str.rsplit("[", 1)[-1].split("]")[0]
                parts = bracket_part.split(",")
                if len(parts) == 2:
                    col_idx = int(parts[0])
                    row_idx = int(parts[1])
                    if row_idx == 0:
                        prefix_path = id_str.rsplit("[", 1)[0]
                        prefix = prefix_path.split("/")[-1]
                        name = child.Name.upper()
                        
                        # Suffix match or exact match to avoid matching TOPLOW
                        if name.endswith("-LOW") or name == "GS_SELFIELDS-LOW":
                            col_info["low_col"] = col_idx
                            col_info["low_prefix"] = prefix
                        elif name.endswith("-OPTION") or name == "GS_SELFIELDS-OPTION" or name == "OPTION":
                            col_info["option_col"] = col_idx
                            col_info["option_prefix"] = prefix
                        elif "FIELDNAME" in name:
                            if col_idx > 10:
                                col_info["technical_name_col"] = col_idx
                                col_info["technical_name_prefix"] = prefix
    except Exception:
        pass
    return col_info


def _aplicar_filtros_base(session, user):
    tbl = _wait_for_table_control(session, 8)
    if not tbl:
        return False

    tbl_id = tbl.Id
    row_count = int(tbl.RowCount)
    visible_rows = int(tbl.VisibleRowCount)

    col_info = _detect_columns(tbl)
    col_fieldname_orig = col_info["technical_name_col"]
    col_fieldname_prefix = col_info["technical_name_prefix"]
    col_option = col_info["option_col"]
    col_option_prefix = col_info["option_prefix"]
    col_low = col_info["low_col"]
    col_low_prefix = col_info["low_prefix"]

    fields_to_find = {"TRSTATUS": None, "AS4USER": None, "STRKORR": None}
    
    # Read visible rows to find positions of the filter fields
    for r in range(min(row_count, visible_rows)):
        try:
            fname_id = f"{tbl_id}/{col_fieldname_prefix}[{col_fieldname_orig},{r}]"
            fieldname = session.findById(fname_id).text.strip().upper()
            if fieldname in fields_to_find:
                fields_to_find[fieldname] = r
        except Exception:
            continue

    status_set = False
    user_set = False

    # 1. Set TRSTATUS = 'D'
    r_status = fields_to_find.get("TRSTATUS")
    if r_status is not None:
        status_set = _set_low_value(session, tbl_id, col_low_prefix, col_low, r_status, "D")

    # 2. Set AS4USER = user
    r_user = fields_to_find.get("AS4USER")
    if r_user is not None:
        user_set = _set_low_value(session, tbl_id, col_low_prefix, col_low, r_user, user)

    # 3. Set STRKORR != "" (Option = "NE", Low = "")
    r_strkorr = fields_to_find.get("STRKORR")
    if r_strkorr is not None:
        low_id = f"{tbl_id}/{col_low_prefix}[{col_low},{r_strkorr}]"
        _try_set_text(session, low_id, "")
        
        # Only set option text if the option column is a text field, not a button
        if not col_option_prefix.lower().startswith("btn"):
            opt_id = f"{tbl_id}/{col_option_prefix}[{col_option},{r_strkorr}]"
            _try_set_text(session, opt_id, "NE")

    return status_set and user_set




def _press_execute(session):
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_not_busy(session, 12)
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"❌ Falha ao executar (F8) no SE16H: {e}")
        return False


def _walk_children(root, max_nodes=8000):
    stack = [root]
    seen = 0
    while stack and seen < max_nodes:
        obj = stack.pop()
        seen += 1
        yield obj
        try:
            cnt = obj.Children.Count
        except Exception:
            continue
        for i in range(cnt - 1, -1, -1):
            try:
                stack.append(obj.Children(i))
            except Exception:
                continue


def _score_grid_candidate(obj):
    try:
        rc = int(obj.RowCount)
        if rc < 0:
            return -1
    except Exception:
        return -1

    score = 0
    if rc > 0:
        score += 5

    # assinatura de E070
    for col in ("TRKORR", "STRKORR", "AS4TEXT"):
        try:
            _ = obj.GetCellValue(0, col)
            score += 10
        except Exception:
            pass

    return score


def _find_best_grid(session):
    # Otimização 1: Tentar caminhos diretos comuns do ALV Grid em SE16H/SE16N primeiro
    comuns = [
        "wnd[0]/usr/cntlRESULT/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/shellcont/shell",
    ]
    for c in comuns:
        try:
            obj = session.findById(c)
            # verifica se possui RowCount e GetCellValue (assinatura de GridView)
            _ = obj.RowCount
            _ = obj.GetCellValue(0, "TRKORR")
            return obj
        except Exception:
            continue

    # Otimização 2: Fallback com varredura genérica caso os caminhos diretos não existam
    roots = []
    try:
        roots.append(session.findById("wnd[0]/usr"))
    except Exception:
        pass
    try:
        roots.append(session.findById("wnd[0]"))
    except Exception:
        pass

    candidates = []
    for root in roots:
        for obj in _walk_children(root):
            s = _score_grid_candidate(obj)
            if s >= 0:
                candidates.append((s, obj))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def _get_cell(grid, row, col):
    try:
        return str(grid.GetCellValue(row, col)).strip()
    except Exception:
        return ""


def _open_se16h_new_mode(session):
    """
    Abre /ose16h em novo modo e devolve (new_session, created_flag).
    """
    before_ids = set()
    try:
        connection = session.Parent
        for i in range(connection.Children.Count):
            before_ids.add(connection.Children(i).Id)
    except Exception:
        connection = None

    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/ose16h"
        session.findById("wnd[0]").sendVKey(0)
    except Exception:
        return session, False

    if connection is not None:
        t0 = time.time()
        while time.time() - t0 <= 8:
            try:
                for i in range(connection.Children.Count):
                    c = connection.Children(i)
                    if c.Id not in before_ids:
                        _wait_not_busy(c, 12)
                        time.sleep(0.3)
                        return c, True
            except Exception:
                pass
            time.sleep(0.2)

    _wait_not_busy(session, 12)
    time.sleep(0.2)
    return session, False


def listar_requests(
    system_name=None,
    max_rows="5000",
    include_requests=False,
    use_new_mode=True,
    minimize=True,
    close_after=True,
):
    times = {}
    t_total_start = time.perf_counter()

    def print_profile():
        total_time = time.perf_counter() - t_total_start
        print("\n⏱️  PERFIL DE TEMPO DA EXECUÇÃO (Mapeamento de Gargalos):")
        print("=" * 65)
        for task_name, duration in times.items():
            percentage = (duration / total_time) * 100
            print(f"- {task_name:<45}: {duration:6.2f}s ({percentage:5.1f}%)")
        print("-" * 65)
        print(f"{'Tempo Total':<45}: {total_time:6.2f}s (100.0%)")
        print("=" * 65)

    # Mapear o sistema para a chave do ambiente de forma a usar o login automático se necessário
    key = None
    if system_name:
        sys_upper = system_name.upper()
        if sys_upper == "S4D":
            key = "S4DCLNT100"
        elif sys_upper == "S4Q":
            key = "S4QCLNT100"
        elif sys_upper == "S4P":
            key = "S4PCLNT100"
        elif sys_upper == "SPA":
            key = "SPACLNT001"

    t_start = time.perf_counter()
    base_session = None
    try:
        # Importar dinamicamente sap_session da raiz do projeto
        base_dir = _base_dir()
        if base_dir not in sys.path:
            sys.path.insert(0, base_dir)
        from sap_session import ensure_sap_access_from_env
        base_session = ensure_sap_access_from_env(key=key, timeout_s=45)
    except Exception as exc:
        print(f"⚠️ Erro ao tentar acesso automático via sap_session: {exc}")
        print("Tentando obter sessão existente manualmente...")

    if not base_session:
        app = _get_application()
        base_session = _pick_session(app, system_name=system_name)

    try:
        user = (base_session.Info.User or "").strip()
    except Exception:
        user = ""

    if not user:
        msg = "Sessão SAP encontrada, mas não está logada (Info.User vazio). Por favor, faça login no SAP GUI e tente novamente."
        print(f"❌ {msg}")
        raise RuntimeError(msg)

    times["Conexão / Acesso SAP GUI"] = time.perf_counter() - t_start

    work_session = base_session
    created_new = False

    t_start = time.perf_counter()
    if use_new_mode:
        work_session, created_new = _open_se16h_new_mode(base_session)
    else:
        try:
            base_session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16h"
            base_session.findById("wnd[0]").sendVKey(0)
            _wait_not_busy(base_session, 12)
            time.sleep(0.2)
        except Exception:
            pass
    times["Abertura do SE16H (Novo Modo ou Transição)"] = time.perf_counter() - t_start

    t_start = time.perf_counter()
    _set_table_e070(work_session)
    _set_max_ocorrencias(work_session, max_rows=max_rows)

    if not _aplicar_filtros_base(work_session, user):
        print("⚠️ Não consegui aplicar TRSTATUS/AS4USER/STRKORR de forma dinâmica. Vou executar mesmo assim.")

    if minimize:
        _iconify(work_session)
    times["Configuração dos Filtros (E070, User, Status)"] = time.perf_counter() - t_start

    t_start = time.perf_counter()
    if not _press_execute(work_session):
        if created_new and close_after:
            _close_window(work_session)
        raise RuntimeError("Falha ao executar a consulta no SE16H do SAP GUI.")

    # Verificação de mensagens de "nenhuma entrada" na barra de status
    try:
        sbar = work_session.findById("wnd[0]/sbar")
        sbar_text = str(sbar.Text).strip().lower()
        if sbar_text and any(term in sbar_text for term in ["nenhum", "no entries", "no values", "not found", "no matching"]):
            print(f"ℹ️ SAP Status Bar: {sbar.Text}")
            if created_new and close_after:
                _close_window(work_session)
            times["Execução da Consulta no SAP (F8)"] = time.perf_counter() - t_start
            print_profile()
            return []
    except Exception:
        pass
    times["Execução da Consulta no SAP (F8)"] = time.perf_counter() - t_start

    t_start = time.perf_counter()
    grid = _find_best_grid(work_session)
    if not grid:
        if created_new and close_after:
            _close_window(work_session)
        raise RuntimeError("Não foi possível encontrar a grelha de resultados do SE16H no SAP GUI.")

    try:
        row_count = int(grid.RowCount)
    except Exception as e:
        if created_new and close_after:
            _close_window(work_session)
        raise RuntimeError(f"Não foi possível obter RowCount da grelha de resultados: {e}")
    times["Localização do ALV Grid"] = time.perf_counter() - t_start

    t_start = time.perf_counter()
    results = []
    for r in range(row_count):
        # Micro-otimização: Ler a coluna STRKORR primeiro. Se estiver vazia, ignoramos imediatamente
        # a linha e poupamos 2 chamadas COM adicionais (TRKORR e AS4TEXT).
        strkorr = _get_cell(grid, r, "STRKORR")
        if not strkorr:
            continue

        trkorr = _get_cell(grid, r, "TRKORR")
        as4text = _get_cell(grid, r, "AS4TEXT") or _get_cell(grid, r, "TXT_BREVE") or _get_cell(grid, r, "TEXT")

        if trkorr or as4text:
            results.append((trkorr, as4text))
    times["Leitura dos Resultados (Loop COM no ALV)"] = time.perf_counter() - t_start

    # impressão NUMERADA
    try:
        sysname = work_session.Info.SystemName
    except Exception:
        sysname = ""

    print(f"\n✅ Resultados: {len(results)} | Sistema={sysname} | User={user}")
    print("N | TRKORR | AS4TEXT")
    print("-" * 90)
    for i, (trkorr, as4text) in enumerate(results, start=1):
        print(f"{i} | {trkorr} | {as4text}")

    # guarda para seleção futura por número
    _save_results(results, system_name=sysname, user=user)

    t_start = time.perf_counter()
    if created_new and close_after:
        _close_window(work_session)
    if created_new and close_after:
        times["Fecho da Janela / Sessão do SE16H"] = time.perf_counter() - t_start

    print_profile()

    return results


def _parse_args(argv):
    system = None
    max_rows = "5000"
    include_requests = False
    use_new_mode = True
    minimize = True
    close_after = True

    i = 0
    while i < len(argv):
        a = argv[i].strip()
        if a == "--system" and i + 1 < len(argv):
            system = argv[i + 1]
            i += 2
            continue
        if a == "--max" and i + 1 < len(argv):
            max_rows = argv[i + 1]
            i += 2
            continue
        if a == "--include-requests":
            include_requests = True
            i += 1
            continue
        if a == "--no-new-mode":
            use_new_mode = False
            i += 1
            continue
        if a == "--no-minimize":
            minimize = False
            i += 1
            continue
        if a == "--no-close":
            close_after = False
            i += 1
            continue
        i += 1

    return system, max_rows, include_requests, use_new_mode, minimize, close_after


if __name__ == "__main__":
    try:
        system, max_rows, include_requests, use_new_mode, minimize, close_after = _parse_args(sys.argv[1:])
        listar_requests(
            system_name=system,
            max_rows=max_rows,
            include_requests=include_requests,
            use_new_mode=use_new_mode,
            minimize=minimize,
            close_after=close_after,
        )
    except Exception as e:
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(f"[ERRO] Erro fatal: {e}", file=sys.stderr)
        sys.exit(1)
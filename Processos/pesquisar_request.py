# -*- coding: utf-8 -*-
"""
pesquisar_request.py

Objetivo:
- Abrir SE16H em NOVO modo (/ose16h) ou modo existente (/nse16h)
- Reutilizar sessões ativas da SE16H para evitar acumular janelas SAP abertas (se em modo CLI).
- Se chamado via Cockpit, aceita a sessão base ativa, abre /ose16h, executa e fecha a sessão temporária criada.
- Minimizar a janela desse novo modo enquanto executa (apenas se criada recentemente)
- Ler resultados da E070 e retornar/exibir lista estruturada
- APENAS listar as linhas cujo valor da coluna STRKORR for diferente de vazio.
- Guardar automaticamente a lista num ficheiro JSON e configurar variáveis de ambiente.

Refatorado para modularidade, reutilização de serviço e rastreabilidade de performance.
# DEBUG TEMPORÁRIO DE PERFORMANCE: remover quando a análise terminar
"""

import sys
import time
import json
import os
import re
import threading
import win32com.client
from dataclasses import dataclass, field
from typing import Optional, Any, Generator, Tuple

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

# Lock simples em memória para evitar pesquisas concorrentes no mesmo processo Python
_REQUEST_SEARCH_LOCK = threading.Lock()

# ─────────────────────────────────────────────────────────────────────────────
# 1. DATACLASSES
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class RequestSearchOptions:
    system_name: Optional[str] = None
    max_rows: str = "5000"
    include_requests: bool = False
    use_new_mode: bool = True
    minimize: bool = True
    close_after: bool = True
    debug_perf: bool = False
    save_cache: bool = True
    print_results: bool = True

@dataclass
class RequestItem:
    idx: int
    trkorr: str
    as4text: str

    def to_tuple(self) -> Tuple[str, str]:
        return (self.trkorr, self.as4text)

@dataclass
class RequestSearchResult:
    items: list[RequestItem]
    system: str
    user: str
    cache_path: Optional[str]
    timings: dict[str, float] = field(default_factory=dict)

@dataclass
class FilterApplyResult:
    ok: bool
    status_set: bool
    user_set: bool
    strkorr_set: bool
    row_count: int = 0
    visible_rows: int = 0

# ─────────────────────────────────────────────────────────────────────────────
# 2. PERFORMANCE TRACKER
# ─────────────────────────────────────────────────────────────────────────────

class PerfTracker:
    _THRESHOLDS = {
        "Conexão / Acesso SAP GUI": 10.0,
        "Abertura do SE16H": 8.0,
        "Configuração da consulta": 8.0,
        "Execução F8": 15.0,
        "Localização ALV Grid": 5.0,
        "Leitura dos resultados": 10.0,
        "Fecho da sessão": 5.0,
    }

    def __init__(self, enabled: bool = False):
        self.enabled = enabled
        self.start_time = time.perf_counter()
        self.last_time = self.start_time
        self.timings: dict[str, float] = {}

    def log(self, step: str, message: str = "", **data):
        if not self.enabled:
            return
        now = time.perf_counter()
        total = now - self.start_time
        delta = now - self.last_time
        self.last_time = now
        extra = ""
        if data:
            extra = " | " + " ".join(f"{k}={v}" for k, v in data.items())
        print(f"[REQ_PERF] +{total:08.3f}s \u0394{delta:06.3f}s | {step} | {message}{extra}")

    def mark(self, name: str, duration: float):
        self.timings[name] = duration

    class _TimeBlockContext:
        def __init__(self, tracker: "PerfTracker", name: str):
            self.tracker = tracker
            self.name = name
            self.t0 = 0.0

        def __enter__(self):
            self.t0 = time.perf_counter()
            self.tracker.log(f"{self.name.upper()}_START", f"Início de: {self.name}")
            return self

        def __exit__(self, exc_type, exc_val, exc_tb):
            dur = time.perf_counter() - self.t0
            self.tracker.mark(self.name, dur)
            self.tracker.log(f"{self.name.upper()}_END", f"Concluído: {self.name}", duration=f"{dur:.2f}s")

    def time_block(self, name: str) -> _TimeBlockContext:
        return self._TimeBlockContext(self, name)

    def summary(self):
        total_time = time.perf_counter() - self.start_time
        print("\n⏱️  PERFIL DE TEMPO DA EXECUÇÃO (Mapeamento de Gargalos):")
        print("=" * 65)
        slowest_name = ""
        slowest_dur = 0.0
        for task_name, duration in self.timings.items():
            percentage = (duration / total_time) * 100 if total_time > 0 else 0
            print(f"- {task_name:<45}: {duration:6.2f}s ({percentage:5.1f}%)")
            if duration > slowest_dur:
                slowest_dur = duration
                slowest_name = task_name
        print("-" * 65)
        print(f"{'Tempo Total':<45}: {total_time:6.2f}s (100.0%)")
        print("=" * 65)

        if self.enabled and slowest_name:
            pct = (slowest_dur / total_time * 100) if total_time > 0 else 0
            print(f"[REQ_PERF_WARN] Maior tempo: {slowest_name} = {slowest_dur:.2f}s ({pct:.0f}%)")
            if "Leitura" in slowest_name or "ALV" in slowest_name:
                print("[REQ_PERF_HINT] Possível causa: muitas chamadas COM GetCellValue. "
                      "Testar max_rows menor ou exportação ALV/clipboard.")
            elif "SE16H" in slowest_name or "Abertura" in slowest_name:
                print("[REQ_PERF_HINT] Possível causa: abertura lenta de sessão /o. "
                      "Testar --no-new-mode para usar a sessão existente.")
            elif "Acesso" in slowest_name or "Conexão" in slowest_name:
                print("[REQ_PERF_HINT] Possível causa: login automático demorou. "
                      "Verificar latência de rede ou SAP Logon.")

        if self.enabled:
            for stage, limit in self._THRESHOLDS.items():
                dur = self.timings.get(stage, 0.0)
                if dur > limit:
                    print(f"[REQ_PERF_WARN] Gargalo provável: {stage} demorou {dur:.2f}s (limite {limit}s).")

# ─────────────────────────────────────────────────────────────────────────────
# 3. BASE DIRECTORIES & UTILS
# ─────────────────────────────────────────────────────────────────────────────

def _base_dir() -> str:
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def _cache_dir() -> str:
    d = os.path.join(_base_dir(), "cache")
    os.makedirs(d, exist_ok=True)
    return d

def _cache_file_path() -> str:
    return os.path.join(_cache_dir(), "last_e070_list.json")

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

def _iter_sessions(application) -> Generator[Any, None, None]:
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

def _pick_session(application, system_name=None) -> Any:
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

def _wait_not_busy(session, timeout_s=12, perf: Optional[PerfTracker] = None, label="") -> bool:
    t0 = time.time()
    if perf:
        perf.log("WAIT_BUSY_START", "A aguardar SAP livre", label=label, timeout=timeout_s)
    while time.time() - t0 <= timeout_s:
        try:
            if not session.Busy:
                elapsed = time.time() - t0
                if perf:
                    perf.log("WAIT_BUSY_END", "SAP livre", label=label, elapsed=f"{elapsed:.2f}s", busy=False)
                return True
        except Exception:
            if perf:
                perf.log("WAIT_BUSY_END", "Exceção ao ler Busy — assume livre", label=label)
            return True
        time.sleep(0.1)
    elapsed = time.time() - t0
    if perf:
        perf.log("WAIT_BUSY_TIMEOUT", "Timeout atingido", label=label, elapsed=f"{elapsed:.2f}s")
    return False

def _try_set_text(session, id_path, value) -> bool:
    try:
        session.findById(id_path).text = value
        return True
    except Exception:
        return False

def _try_press(session, id_path) -> bool:
    try:
        session.findById(id_path).press()
        return True
    except Exception:
        return False

def _iconify(session) -> bool:
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
    try:
        if _try_press(session, "wnd[1]/usr/btnSPOP-OPTION1"):
            return
        if _try_press(session, "wnd[1]/tbar[0]/btn[0]"):
            return
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
# 4. MONITORIZAÇÃO DE SESSÕES & REUTILIZAÇÃO
# ─────────────────────────────────────────────────────────────────────────────

def _count_sap_sessions(application) -> int:
    count = 0
    try:
        for i in range(application.Children.Count):
            conn = application.Children(i)
            count += conn.Children.Count
    except Exception:
        pass
    return count

def listar_sessoes_sap(application) -> list[dict]:
    sessoes = []
    try:
        for i in range(application.Children.Count):
            conn = application.Children(i)
            for j in range(conn.Children.Count):
                sess = conn.Children(j)
                try:
                    sysname = (sess.Info.SystemName or "").upper()
                    user = (sess.Info.User or "").strip()
                    trans = (sess.Info.Transaction or "").strip().upper()
                except Exception:
                    sysname, user, trans = "", "", ""
                
                title = ""
                try:
                    title = sess.findById("wnd[0]").Text
                except Exception:
                    pass
                
                sessoes.append({
                    "conn_idx": i,
                    "sess_idx": j,
                    "system": sysname,
                    "user": user,
                    "transaction": trans,
                    "title": title,
                    "session": sess
                })
    except Exception:
        pass
    return sessoes

def encontrar_sessao_se16h_reutilizavel(application, system_name=None) -> Optional[Any]:
    sessoes = listar_sessoes_sap(application)
    target_sys = system_name.upper() if system_name else None
    
    for s in sessoes:
        sess = s["session"]
        if target_sys and s["system"] != target_sys:
            continue
        if not s["user"]:
            continue
        try:
            if sess.Busy:
                continue
        except Exception:
            continue
            
        trans = s["transaction"]
        title = s["title"].upper()
        
        if trans == "SE16H" or "SE16H" in title or "E070" in title:
            return sess
            
    return None

# ─────────────────────────────────────────────────────────────────────────────
# 5. ACESSO SAP
# ─────────────────────────────────────────────────────────────────────────────

def obter_sessao_sap(system_name: Optional[str] = None, perf: Optional[PerfTracker] = None) -> tuple[Any, str]:
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

    base_session = None
    if perf:
        perf.log("SAP_ACCESS_START", "A obter sessão SAP", system=system_name)

    try:
        base_dir = _base_dir()
        if base_dir not in sys.path:
            sys.path.insert(0, base_dir)
        from sap_session import ensure_sap_access_from_env
        base_session = ensure_sap_access_from_env(key=key, timeout_s=45)
    except Exception as exc:
        if perf:
            perf.log("SAP_ACCESS_ENV_FAIL", f"Acesso automático via sap_session indisponível: {exc}")

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

    return base_session, user

# ─────────────────────────────────────────────────────────────────────────────
# 6. NAVEGAÇÃO / ABERTURA SE16H
# ─────────────────────────────────────────────────────────────────────────────

def abrir_se16h_em_novo_modo(session_base: Any, perf: Optional[PerfTracker] = None) -> tuple[Any, bool]:
    """
    Abre /ose16h em novo modo a partir de uma sessão base já existente/logada.
    Identifica a nova sessão comparando as sessões antes e depois do /ose16h.
    Não tenta validar transação ou procurar SE16H já aberta (comportamento Cockpit).
    """
    connection = session_base.Parent
    before_ids = set()
    try:
        for i in range(connection.Children.Count):
            before_ids.add(connection.Children(i).Id)
    except Exception:
        pass

    if perf and perf.enabled:
        print(f"[REQ_PERF] A abrir /ose16h a partir da sessão base")
        print(f"[REQ_PERF] Sessões antes do /ose16h: {len(before_ids)}")

    try:
        session_base.findById("wnd[0]/tbar[0]/okcd").text = "/ose16h"
        session_base.findById("wnd[0]").sendVKey(0)
    except Exception as exc:
        raise RuntimeError(f"Falha ao enviar /ose16h na sessão base: {exc}")

    t0 = time.time()
    work_session = None
    while time.time() - t0 <= 10:
        try:
            for i in range(connection.Children.Count):
                c = connection.Children(i)
                if c.Id not in before_ids:
                    _wait_not_busy(c, 12, perf=perf, label="new_mode_load")
                    time.sleep(0.3)
                    work_session = c
                    break
        except Exception:
            pass
        if work_session is not None:
            break
        time.sleep(0.2)

    if work_session is None:
        raise RuntimeError("Não foi possível identificar a nova sessão SE16H criada após enviar /ose16h.")

    if perf and perf.enabled:
        print(f"[REQ_PERF] Sessões depois do /ose16h: {connection.Children.Count}")
        print("[REQ_PERF] Nova sessão SE16H identificada: sim")

    return work_session, True

def abrir_se16h(session, use_new_mode: bool = True, perf: Optional[PerfTracker] = None, system_name: Optional[str] = None) -> tuple[Any, bool]:
    """
    Usado apenas no modo CLI/Isolado. Abre a transação SE16H.
    Tenta reutilizar uma sessão se16h existente se disponível.
    """
    app = _get_application()
    
    if use_new_mode:
        reusable_sess = encontrar_sessao_se16h_reutilizavel(app, system_name)
        if reusable_sess is not None:
            if perf and perf.enabled:
                print("[REQ_PERF] Sessão SE16H reutilizável encontrada: sim")
                print("[REQ_PERF] Nova sessão criada com /ose16h: não")
            return reusable_sess, False
        
        if perf and perf.enabled:
            print("[REQ_PERF] Sessão SE16H reutilizável encontrada: não")
            print("[REQ_PERF] Nova sessão criada com /ose16h: sim")

        # Abre nova sessão
        before_ids = set()
        try:
            connection = session.Parent
            for i in range(connection.Children.Count):
                before_ids.add(connection.Children(i).Id)
        except Exception:
            connection = None

        if perf:
            perf.log("OPEN_SE16H_START", "A abrir /ose16h em novo modo")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/ose16h"
            session.findById("wnd[0]").sendVKey(0)
        except Exception as exc:
            if perf:
                perf.log("OPEN_SE16H_FAIL", f"Falha ao enviar /ose16h: {exc}")
            return session, False

        if connection is not None:
            t0 = time.time()
            while time.time() - t0 <= 8:
                try:
                    for i in range(connection.Children.Count):
                        c = connection.Children(i)
                        if c.Id not in before_ids:
                            _wait_not_busy(c, 12, perf=perf, label="new_mode_load")
                            time.sleep(0.3)
                            return c, True
                except Exception:
                    pass
                time.sleep(0.2)
        _wait_not_busy(session, 12, perf=perf, label="fallback_new_mode")
        time.sleep(0.2)
        return session, False
    else:
        if perf and perf.enabled:
            print("[REQ_PERF] Sessão SE16H reutilizável encontrada: não (new_mode desativado)")
            print("[REQ_PERF] Nova sessão criada com /ose16h: não")
        if perf:
            perf.log("OPEN_SE16H_START", "A transitar para /nse16h na sessão atual")
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nse16h"
            session.findById("wnd[0]").sendVKey(0)
            _wait_not_busy(session, 12, perf=perf, label="transition_se16h")
            time.sleep(0.2)
        except Exception as exc:
            if perf:
                perf.log("OPEN_SE16H_FAIL", f"Falha ao enviar /nse16h: {exc}")
        return session, False

# ─────────────────────────────────────────────────────────────────────────────
# 7. FILTROS E CONFIGURAÇÃO DA CONSULTA
# ─────────────────────────────────────────────────────────────────────────────

def _wait_for_table_control(session, timeout_s=8, perf: Optional[PerfTracker] = None) -> Any:
    t0 = time.time()
    while time.time() - t0 <= timeout_s:
        tbl = _find_table_control(session)
        if tbl is not None:
            return tbl
        time.sleep(0.1)
    return None

def configurar_consulta_e070(session, user: str, max_rows: str, perf: Optional[PerfTracker] = None) -> FilterApplyResult:
    # 1. Definir E070
    cid = _wait_for_table_input_field(session, 5, perf=perf)
    if not cid or not _try_set_text(session, cid, "E070"):
        return FilterApplyResult(ok=False, status_set=False, user_set=False, strkorr_set=False)

    try:
        session.findById("wnd[0]").sendVKey(0)
    except Exception:
        pass
    _wait_not_busy(session, 10, perf=perf, label="set_table_name")
    time.sleep(0.2)

    # 2. Definir max occurrences
    max_cids = ["wnd[0]/usr/txtMAX_SEL", "wnd[0]/usr/txtGD-MAXROWS", "wnd[0]/usr/txtMAX_HITS"]
    max_ok = False
    for mcid in max_cids:
        if _try_set_text(session, mcid, str(max_rows)):
            max_ok = True
            break

    # 3. Aplicar Filtros E070
    tbl = _wait_for_table_control(session, 8, perf=perf)
    if not tbl:
        return FilterApplyResult(ok=False, status_set=False, user_set=False, strkorr_set=False)

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
    strkorr_set = False

    # Set TRSTATUS = 'D'
    r_status = fields_to_find.get("TRSTATUS")
    if r_status is not None:
        status_set = _set_low_value(session, tbl_id, col_low_prefix, col_low, r_status, "D")

    # Set AS4USER = user
    r_user = fields_to_find.get("AS4USER")
    if r_user is not None:
        user_set = _set_low_value(session, tbl_id, col_low_prefix, col_low, r_user, user)

    # Set STRKORR != "" (Option = "NE", Low = "")
    r_strkorr = fields_to_find.get("STRKORR")
    if r_strkorr is not None:
        low_id = f"{tbl_id}/{col_low_prefix}[{col_low},{r_strkorr}]"
        _try_set_text(session, low_id, "")
        if not col_option_prefix.lower().startswith("btn"):
            opt_id = f"{tbl_id}/{col_option_prefix}[{col_option},{r_strkorr}]"
            strkorr_set = _try_set_text(session, opt_id, "NE")
        else:
            strkorr_set = True

    ok = status_set and user_set and max_ok
    return FilterApplyResult(
        ok=ok,
        status_set=status_set,
        user_set=user_set,
        strkorr_set=strkorr_set,
        row_count=row_count,
        visible_rows=visible_rows
    )

def executar_consulta_f8(session, perf: Optional[PerfTracker] = None) -> bool:
    try:
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _wait_not_busy(session, 12, perf=perf, label="press_execute")
        time.sleep(0.3)
        return True
    except Exception as e:
        print(f"❌ Falha ao executar (F8) no SE16H: {e}")
        return False

# ─────────────────────────────────────────────────────────────────────────────
# 8. LOCALIZAÇÃO E LEITURA DE ALV
# ─────────────────────────────────────────────────────────────────────────────

def localizar_alv_grid(session, perf: Optional[PerfTracker] = None) -> Any:
    comuns = [
        "wnd[0]/usr/cntlRESULT/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/shellcont/shell",
    ]
    for c in comuns:
        try:
            obj = session.findById(c)
            _ = obj.RowCount
            _ = obj.GetCellValue(0, "TRKORR")
            if perf:
                perf.log("GRID_DIRECT_HIT", "ALV Grid encontrado por caminho direto", path=c)
            return obj
        except Exception:
            continue

    if perf:
        perf.log("GRID_FALLBACK_START", "A procurar ALV Grid de forma recursiva")
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
    nodes_walked = 0
    for root in roots:
        for obj in _walk_children(root):
            nodes_walked += 1
            s = _score_grid_candidate(obj)
            if s >= 0:
                candidates.append((s, obj))

    if perf:
        perf.log("GRID_FALLBACK_END", "Pesquisa recursiva terminada", nodes_walked=nodes_walked, candidates=len(candidates))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

def ler_resultados_e070(grid, max_rows: Optional[int] = None, perf: Optional[PerfTracker] = None) -> list[RequestItem]:
    try:
        row_count = int(grid.RowCount)
    except Exception as e:
        raise RuntimeError(f"Não foi possível obter RowCount do ALV: {e}")

    limit = row_count
    if max_rows is not None and max_rows > 0:
        limit = min(row_count, max_rows)

    text_col = detectar_coluna_texto(grid)
    
    results = []
    counters = {"com_calls": 0}
    rows_seen = 0
    skipped_empty_strkorr = 0
    kept = 0
    _PROGRESS_EVERY = 500
    
    t_start = time.perf_counter()

    for r in range(limit):
        rows_seen += 1
        
        strkorr = _get_cell(grid, r, "STRKORR", counters)
        if not strkorr:
            skipped_empty_strkorr += 1
            continue

        trkorr = _get_cell(grid, r, "TRKORR", counters)
        
        as4text = ""
        if text_col:
            as4text = _get_cell(grid, r, text_col, counters)

        if trkorr or as4text:
            kept += 1
            results.append(RequestItem(idx=kept, trkorr=trkorr, as4text=as4text))

        if perf and perf.enabled and rows_seen % _PROGRESS_EVERY == 0:
            elapsed = time.perf_counter() - t_start
            avg_ms = (elapsed / rows_seen * 1000) if rows_seen else 0
            perf.log("ALV_READ_PROGRESS",
                     "Progresso da leitura do ALV",
                     row=f"{rows_seen}/{limit}",
                     kept=kept,
                     skipped_empty=skipped_empty_strkorr,
                     com_calls=counters["com_calls"],
                     avg_ms_per_row=f"{avg_ms:.1f}")

    return results

# ─────────────────────────────────────────────────────────────────────────────
# 9. CACHE JSON
# ─────────────────────────────────────────────────────────────────────────────

def guardar_cache_requests(items: list[RequestItem], system_name: str, user: str, expose_env: bool = True) -> str:
    payload = {
        "meta": {
            "system": system_name,
            "user": user,
            "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        },
        "items": [
            {"idx": item.idx, "TRKORR": item.trkorr, "AS4TEXT": item.as4text}
            for item in items
        ],
    }

    path = _cache_file_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    if expose_env:
        os.environ["SAP_LAST_E070_LIST_FILE"] = path
        os.environ["SAP_LAST_E070_LIST_COUNT"] = str(len(items))

    return path

# ─────────────────────────────────────────────────────────────────────────────
# 10. IMPRESSÃO / OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def imprimir_resultados(items: list[RequestItem], system: str, user: str):
    print(f"\n✅ Resultados: {len(items)} | Sistema={system} | User={user}")
    print("N | TRKORR | AS4TEXT")
    print("-" * 90)
    for item in items:
        print(f"{item.idx} | {item.trkorr} | {item.as4text}")

# ─────────────────────────────────────────────────────────────────────────────
# 11. API PÚBLICA / SERVIÇO COM SELEÇÃO DE MODO (COCKPIT VS CLI)
# ─────────────────────────────────────────────────────────────────────────────

def pesquisar_requests_service(options: RequestSearchOptions, session: Optional[Any] = None) -> RequestSearchResult:
    # Adquirir lock para evitar execuções concorrentes no mesmo processo Python
    if not _REQUEST_SEARCH_LOCK.acquire(blocking=False):
        if options.debug_perf:
            print("[REQ_PERF] Pesquisa de request já em execução; aguardando/liberando conforme configuração.")
        if not _REQUEST_SEARCH_LOCK.acquire(blocking=True, timeout=60):
            raise RuntimeError("Pesquisa de request já em execução; limite de aguardo excedido.")

    perf = PerfTracker(enabled=options.debug_perf)
    work_session = None
    created_new = False
    user = ""
    system_name = options.system_name

    try:
        if session is not None:
            # Modo Cockpit / Serviço
            if options.debug_perf:
                print("[REQ_PERF] Modo de execução: cockpit_session")
                print("[REQ_PERF] Sessão SAP recebida do processo principal: sim")
            
            session_base = session
            try:
                user = (session_base.Info.User or "").strip()
                system_name = (session_base.Info.SystemName or "").upper()
            except Exception:
                user = ""

            # Abre o novo modo (/ose16h) a partir da sessão recebida
            work_session, created_new = abrir_se16h_em_novo_modo(session_base, perf=perf)
        else:
            # Modo CLI / Execução isolada
            if options.debug_perf:
                print("[REQ_PERF] Modo de execução: cli_isolated")
                print("[REQ_PERF] Sessão SAP recebida do processo principal: não")
            
            app = _get_application()
            initial_sessions = _count_sap_sessions(app)
            if options.debug_perf:
                print(f"[REQ_PERF] Sessões SAP antes da pesquisa: {initial_sessions}")

            # 1. Conexão / Acesso SAP
            with perf.time_block("Conexão / Acesso SAP GUI"):
                session_base, user = obter_sessao_sap(options.system_name, perf=perf)
                if not system_name:
                    try:
                        system_name = (session_base.Info.SystemName or "").upper()
                    except Exception:
                        pass

            # 2. Abertura do SE16H (Reutilizando sessão existente se disponível)
            with perf.time_block("Abertura do SE16H"):
                work_session, created_new = abrir_se16h(session_base, options.use_new_mode, perf=perf, system_name=system_name)

            if options.debug_perf:
                print(f"[REQ_PERF] Sessões SAP depois de abrir SE16H: {_count_sap_sessions(app)}")

        # 3. Configuração dos Filtros
        with perf.time_block("Configuração da consulta"):
            cfg = configurar_consulta_e070(work_session, user, options.max_rows, perf=perf)
            if not cfg.ok:
                perf.log("WARN", "Configuração da consulta incompleta (TRSTATUS/AS4USER falhou)")
            
            # Só minimiza se a sessão foi recém-criada por este processo
            if options.minimize and created_new:
                _iconify(work_session)

        # 4. Execução F8
        with perf.time_block("Execução F8"):
            if not executar_consulta_f8(work_session, perf=perf):
                raise RuntimeError("Falha ao executar consulta no SE16H.")

            # Validar status bar por "nenhum registo"
            try:
                sbar = work_session.findById("wnd[0]/sbar")
                sbar_text = str(sbar.Text).strip().lower()
                if sbar_text and any(term in sbar_text for term in ["nenhum", "no entries", "no values", "not found", "no matching"]):
                    perf.log("INFO", f"SAP Status Bar indica zero resultados: {sbar.Text}")
                    
                    if created_new and options.close_after:
                        _close_window(work_session)
                        if options.debug_perf:
                            print("[REQ_PERF] A fechar sessão temporária SE16H: sim")
                    return RequestSearchResult(items=[], system=system_name or "", user=user, cache_path=None, timings=perf.timings)
            except Exception:
                pass

        # 5. Localização ALV Grid
        with perf.time_block("Localização ALV Grid"):
            grid = localizar_alv_grid(work_session, perf=perf)
            if not grid:
                raise RuntimeError("Não foi possível encontrar a grelha ALV do SE16H.")

        # 6. Leitura dos Resultados
        with perf.time_block("Leitura dos resultados"):
            max_limit = int(options.max_rows) if options.max_rows.isdigit() else None
            items = ler_resultados_e070(grid, max_limit, perf=perf)

    finally:
        # 7. Cleanup de sessões temporárias com try/finally
        if created_new and options.close_after and work_session:
            with perf.time_block("Fecho da sessão"):
                try:
                    _close_window(work_session)
                    if options.debug_perf:
                        print("[REQ_PERF] A fechar sessão temporária SE16H: sim")
                except Exception as exc:
                    if options.debug_perf:
                        print(f"[REQ_PERF] Erro ao fechar sessão temporária: {exc}")
        else:
            if options.debug_perf:
                print("[REQ_PERF] A fechar sessão temporária SE16H: não")

        if options.debug_perf:
            try:
                app = _get_application()
                print(f"[REQ_PERF] Sessões SAP no final: {_count_sap_sessions(app)}")
            except Exception:
                pass
            
        _REQUEST_SEARCH_LOCK.release()

    # 8. Gravar Cache
    cache_path = None
    if options.save_cache and items:
        cache_path = guardar_cache_requests(items, system_name or "", user)

    # 9. Imprimir se pedido
    if options.print_results:
        imprimir_resultados(items, system_name or "", user)

    # Mostrar tempos
    perf.summary()

    return RequestSearchResult(
        items=items,
        system=system_name or "",
        user=user,
        cache_path=cache_path,
        timings=perf.timings
    )

def listar_requests(
    system_name=None,
    max_rows="5000",
    include_requests=False,
    use_new_mode=True,
    minimize=True,
    close_after=True,
    debug_perf=False,
    session=None,
) -> list[tuple[str, str]]:
    options = RequestSearchOptions(
        system_name=system_name,
        max_rows=max_rows,
        include_requests=include_requests,
        use_new_mode=use_new_mode,
        minimize=minimize,
        close_after=close_after,
        debug_perf=debug_perf,
        save_cache=True,
        print_results=True
    )
    result = pesquisar_requests_service(options, session=session)
    return [item.to_tuple() for item in result.items]

# ─────────────────────────────────────────────────────────────────────────────
# 12. CLI / AUXILIARES INTERNOS REQUERIDOS
# ─────────────────────────────────────────────────────────────────────────────

def _parse_args(argv) -> RequestSearchOptions:
    opt = RequestSearchOptions()
    i = 0
    while i < len(argv):
        a = argv[i].strip()
        if a == "--system" and i + 1 < len(argv):
            opt.system_name = argv[i + 1]
            i += 2
            continue
        if a == "--max" and i + 1 < len(argv):
            opt.max_rows = argv[i + 1]
            i += 2
            continue
        if a == "--include-requests":
            opt.include_requests = True
            i += 1
            continue
        if a == "--no-new-mode":
            opt.use_new_mode = False
            i += 1
            continue
        if a == "--no-minimize":
            opt.minimize = False
            i += 1
            continue
        if a == "--no-close":
            opt.close_after = False
            i += 1
            continue
        if a == "--debug-perf":
            opt.debug_perf = True
            i += 1
            continue
        i += 1
    return opt

def _get_cell(grid, row: int, col: str, counters: Optional[dict] = None) -> str:
    if counters is not None:
        counters["com_calls"] += 1
    try:
        return str(grid.GetCellValue(row, col)).strip()
    except Exception:
        return ""

def detectar_coluna_texto(grid, candidates=("AS4TEXT", "TXT_BREVE", "TEXT")) -> Optional[str]:
    for col in candidates:
        try:
            _ = grid.GetCellValue(0, col)
            return col
        except Exception:
            pass
    return None

def _detect_columns(tbl) -> dict[str, Any]:
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

def _walk_children(root, max_nodes=8000) -> Generator[Any, None, None]:
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

def _score_grid_candidate(obj) -> int:
    try:
        rc = int(obj.RowCount)
        if rc < 0:
            return -1
    except Exception:
        return -1
    score = 0
    if rc > 0:
        score += 5
    for col in ("TRKORR", "STRKORR", "AS4TEXT"):
        try:
            _ = obj.GetCellValue(0, col)
            score += 10
        except Exception:
            pass
    return score

if __name__ == "__main__":
    if "--test" in sys.argv:
        print("🧪 A executar testes mínimos...")
        opt_test = RequestSearchOptions(system_name="S4D", debug_perf=True)
        assert opt_test.system_name == "S4D"
        assert opt_test.debug_perf is True
        
        p_test = PerfTracker(enabled=False)
        p_test.mark("Teste", 1.23)
        assert p_test.timings["Teste"] == 1.23
        
        item_test = RequestItem(idx=1, trkorr="S4DK900001", as4text="Teste")
        path_test = guardar_cache_requests([item_test], "S4D", "USER_TEST", expose_env=False)
        assert os.path.exists(path_test)
        with open(path_test, "r") as f:
            data = json.load(f)
            assert data["items"][0]["TRKORR"] == "S4DK900001"
            
        print("✅ Testes mínimos executados com sucesso!")
        sys.exit(0)

    try:
        options = _parse_args(sys.argv[1:])
        pesquisar_requests_service(options)
    except Exception as e:
        import traceback
        traceback.print_exc(file=sys.stderr)
        print(f"[ERRO] Erro fatal: {e}", file=sys.stderr)
        sys.exit(1)
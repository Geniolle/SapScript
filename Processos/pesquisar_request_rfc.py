# -*- coding: utf-8 -*-
"""
pesquisar_request_rfc.py

Objetivo:
- Pesquisa de requests/tarefas SAP ativas 100% via RFC (sem necessidade de SAP GUI aberta)
- Lê tabelas E070 (cabeçalho/tarefas) e E07T (descrições) usando a função RFC_READ_TABLE via pyrfc
- Totalmente compatível com a interface pública do pesquisar_request.py (dataclasses, cache JSON e função listar_requests)
- Oferece alto desempenho e fiabilidade para integração em scripts Python e Web APIs
"""

import sys
import time
import json
import os
import re
import functools
from dataclasses import dataclass, field
from typing import Optional, Any

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
# 1. DATACLASSES E ESTRUTURAS DE DADOS
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class RequestItem:
    idx: int
    trkorr: str
    as4text: str

    def to_tuple(self) -> tuple[str, str]:
        return (self.trkorr, self.as4text)


@dataclass
class RequestSearchOptions:
    system_name: Optional[str] = None
    max_rows: str = "5000"
    include_requests: bool = False
    debug_perf: bool = False
    save_cache: bool = True
    print_results: bool = True
    user_filter: Optional[str] = None


@dataclass
class RequestSearchResult:
    items: list[RequestItem] = field(default_factory=list)
    system: str = ""
    user: str = ""
    cache_path: str = ""
    timings: dict[str, float] = field(default_factory=dict)


# ─────────────────────────────────────────────────────────────────────────────
# 2. CREDENCIAIS E CONEXÃO RFC
# ─────────────────────────────────────────────────────────────────────────────

MAPA_SISTEMA = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}

def obter_conexao_rfc(system_name: Optional[str] = None) -> tuple[Any, str, str]:
    """
    Obtém uma conexão pyrfc.Connection com base no ficheiro .env.
    Retorna: (conn, system_name, user)
    """
    if not HAS_PYRFC:
        raise RuntimeError("A biblioteca 'pyrfc' não está instalada ou configurada neste ambiente.")

    system_up = (system_name or os.getenv("SAP_SYSTEM") or "DEV").upper().strip()
    
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

    REVERSE_MAPA = {"DEV": "S4D", "S4D": "S4D", "QAD": "S4Q", "S4Q": "S4Q", "PRD": "S4P", "S4P": "S4P"}
    system_code = REVERSE_MAPA.get(system_up, system_up)

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
# 3. LEITURA DE TABELAS VIA RFC_READ_TABLE
# ─────────────────────────────────────────────────────────────────────────────

def _rfc_read_table(conn: Any, table_name: str, fields: list[str], options: list[str], rowcount: int = 5000) -> list[dict[str, str]]:
    """
    Wrapper seguro para a RFC_READ_TABLE do SAP.
    """
    fields_payload = [{"FIELDNAME": f} for f in fields]
    options_payload = [{"TEXT": opt} for opt in options]

    res = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE=table_name,
        DELIMITER="|",
        FIELDS=fields_payload,
        OPTIONS=options_payload,
        ROWCOUNT=rowcount
    )

    sap_fields = [entry["FIELDNAME"].strip() for entry in res.get("FIELDS", [])]
    rows: list[dict[str, str]] = []
    for row in res.get("DATA", []):
        raw_wa = str(row.get("WA", ""))
        parts = raw_wa.split("|")
        row_dict = {}
        for idx, f_name in enumerate(sap_fields):
            row_dict[f_name] = parts[idx].strip() if idx < len(parts) else ""
        rows.append(row_dict)

    return rows


def pesquisar_requests_rfc_service(options: RequestSearchOptions) -> RequestSearchResult:
    """
    Serviço principal de pesquisa de requests via RFC.
    """
    t0 = time.time()
    timings = {}

    # 1. Conexão RFC
    t_conn_start = time.time()
    conn, system_code, user_conn = obter_conexao_rfc(options.system_name)
    user_target = (options.user_filter or user_conn).upper().strip()
    timings["conexao_rfc"] = time.time() - t_conn_start

    max_rows_int = int(options.max_rows) if str(options.max_rows).isdigit() else 5000

    # 2. Consulta E070
    t_e070_start = time.time()
    # TRSTATUS: D (Modifiable), O (Release started)
    opt_e070 = [
        f"AS4USER = '{user_target}' AND ( TRSTATUS = 'D' OR TRSTATUS = 'O' )"
    ]

    try:
        rows_e070 = _rfc_read_table(
            conn,
            table_name="E070",
            fields=["TRKORR", "STRKORR", "AS4USER", "TRFUNCTION", "TRSTATUS", "AS4DATE", "AS4TIME"],
            options=opt_e070,
            rowcount=max_rows_int
        )
    except Exception:
        # Fallback se query combinada falhar em algumas versões SAP
        opt_e070_simple = [f"AS4USER = '{user_target}'"]
        rows_e070 = _rfc_read_table(
            conn,
            table_name="E070",
            fields=["TRKORR", "STRKORR", "AS4USER", "TRFUNCTION", "TRSTATUS", "AS4DATE", "AS4TIME"],
            options=opt_e070_simple,
            rowcount=max_rows_int
        )
        rows_e070 = [r for r in rows_e070 if r.get("TRSTATUS") in ("D", "O")]

    timings["leitura_e070"] = time.time() - t_e070_start

    # 3. Filtrar e agrupar pelos IDs dos Cabeçalhos Principais (evitar retornar ID da sub-tarefa)
    header_trkorrs_ordered: list[str] = []
    header_trkorrs_set: set[str] = set()

    for r in rows_e070:
        trkorr = r.get("TRKORR", "").strip()
        strkorr = r.get("STRKORR", "").strip()
        if not trkorr:
            continue

        header_id = strkorr if strkorr else trkorr
        if header_id and header_id not in header_trkorrs_set:
            header_trkorrs_set.add(header_id)
            header_trkorrs_ordered.append(header_id)

    # 4. Consulta E07T (Descrições das Requests Principais)
    t_e07t_start = time.time()
    text_map: dict[str, str] = {}

    if header_trkorrs_ordered:
        CHUNK_SIZE = 25
        for i in range(0, len(header_trkorrs_ordered), CHUNK_SIZE):
            chunk = header_trkorrs_ordered[i:i + CHUNK_SIZE]
            opt_e07t = []
            for idx_c, trk in enumerate(chunk):
                if idx_c == 0:
                    opt_e07t.append(f"TRKORR = '{trk}'")
                else:
                    opt_e07t.append(f"OR TRKORR = '{trk}'")

            try:
                rows_e07t = _rfc_read_table(
                    conn,
                    table_name="E07T",
                    fields=["TRKORR", "AS4TEXT", "LANGU"],
                    options=opt_e07t,
                    rowcount=2000
                )
                for r_t in rows_e07t:
                    k = r_t.get("TRKORR", "").strip()
                    txt = r_t.get("AS4TEXT", "").strip()
                    lang = r_t.get("LANGU", "").strip()
                    if k not in text_map or lang in ("P", "E"):
                        text_map[k] = txt
            except Exception as exc:
                if options.debug_perf:
                    print(f"⚠️ Aviso ao ler E07T no chunk: {exc}")

    timings["leitura_e07t"] = time.time() - t_e07t_start

    # 5. Montar lista final de RequestItems com os IDs dos Cabeçalhos Principais
    items: list[RequestItem] = []
    for idx, header_id in enumerate(header_trkorrs_ordered, start=1):
        as4text = text_map.get(header_id, f"Request {header_id}")
        items.append(RequestItem(idx=idx, trkorr=header_id, as4text=as4text))

    # 5. Salvar Cache JSON se solicitado
    cache_path = ""
    if options.save_cache:
        try:
            cache_dir = os.path.dirname(os.path.abspath(__file__))
            cache_path = os.path.join(cache_dir, "cache_requests.json")
            payload = {
                "meta": {
                    "system": system_code,
                    "user": user_target,
                    "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                    "mode": "RFC"
                },
                "items": [
                    {"idx": item.idx, "TRKORR": item.trkorr, "AS4TEXT": item.as4text}
                    for item in items
                ],
            }
            with open(cache_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            os.environ["SAP_CACHE_REQUESTS_PATH"] = cache_path
            os.environ["SAP_ULTIMA_REQUEST"] = items[0].trkorr if items else ""
        except Exception as e:
            if options.debug_perf:
                print(f"⚠️ Não foi possível salvar o cache de requests: {e}")

    timings["total"] = time.time() - t0

    if options.print_results:
        imprimir_resultados(items, system_code, user_target)
        if options.debug_perf:
            print("\n⏱️ Performance RFC:")
            for step, duration in timings.items():
                print(f"   - {step}: {duration:.3f}s")

    return RequestSearchResult(
        items=items,
        system=system_code,
        user=user_target,
        cache_path=cache_path,
        timings=timings
    )


def imprimir_resultados(items: list[RequestItem], system: str, user: str):
    print(f"\n✅ Resultados RFC: {len(items)} | Sistema={system} | User={user}")
    print("N | TRKORR | AS4TEXT")
    print("-" * 90)
    for item in items:
        print(f"{item.idx} | {item.trkorr} | {item.as4text}")


def listar_requests(
    system_name=None,
    max_rows="5000",
    include_requests=False,
    debug_perf=False,
    user_filter=None,
) -> list[tuple[str, str]]:
    """
    Interface simplificada para utilização noutros scripts do projeto.
    Retorna uma lista de tuplos: [("S4QK900123", "Descrição da tarefa..."), ...]
    """
    options = RequestSearchOptions(
        system_name=system_name,
        max_rows=str(max_rows),
        include_requests=include_requests,
        debug_perf=debug_perf,
        save_cache=True,
        print_results=True,
        user_filter=user_filter
    )
    result = pesquisar_requests_rfc_service(options)
    return [item.to_tuple() for item in result.items]


# ─────────────────────────────────────────────────────────────────────────────
# 4. EXECUÇÃO VIA TERMINAL CLI
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Pesquisar Requests/Tarefas SAP via RFC.")
    parser.add_argument("--system", type=str, default=None, help="Sistema SAP (ex: DEV, QAD, S4D, S4Q)")
    parser.add_argument("--max", type=str, default="5000", help="Máximo de registos (padrão: 5000)")
    parser.add_argument("--include-requests", action="store_true", help="Incluir requests pai além das tarefas")
    parser.add_argument("--user", type=str, default=None, help="Filtrar por utilizador específico")
    parser.add_argument("--debug-perf", action="store_true", help="Exibir tempos de execução")

    args = parser.parse_args()

    try:
        listar_requests(
            system_name=args.system,
            max_rows=args.max,
            include_requests=args.include_requests,
            debug_perf=args.debug_perf,
            user_filter=args.user
        )
    except Exception as exc:
        print(f"\n❌ Erro na execução da pesquisa via RFC: {exc}")
        sys.exit(1)

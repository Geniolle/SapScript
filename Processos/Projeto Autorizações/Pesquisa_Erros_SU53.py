# -*- coding: utf-8 -*-
"""
Pesquisa_Erros_SU53.py
======================================================================
Pesquisa de Erros de Autorização SU53 via RFC e Fluxo de Atribuição:
  1. Consulta o buffer de memória partilhada da SU53 do utilizador via RFC
     (SUSR_USER_SU53_READ) em todos os servidores aplicacionais (PRD).
  2. Apresenta a análise integrada com:
     - Transações bloqueadas no objeto S_TCODE
     - Outros Objetos de Autorização com falha (S_USER_UID, S_SAL_LOG, etc.)
  3. Permite selecionar tanto Transações como Objetos de Autorização:
     - Para Transações (S_TCODE):
       * Localiza o utilizador na folha 'Proposta Ativa' (Departamento e Composite Role).
       * Valida e marca 'X' na folha departamental correspondente.
       * Localiza a Função Individual (Single Role Z_...) na folha 'Proposta'.
       * Adiciona a Função Individual na folha 'Proposta Ativa'.
       * Adiciona a relação na folha 'PFCG_COMPOSTA'.
       * Grava o Excel (Excel COM ou openpyxl, com backup).
       * Executa no SAP PRD e QAS (QAD) via RFC com Ajuste de Utilizadores.
     - Para Objetos de Autorização:
       * Apresenta diagnóstico técnico do objeto, campos testados, valores,
         autorizações atuais do utilizador e roles que contêm o objeto.
======================================================================
"""

import os
import sys
import re
import shutil
import unicodedata
from decimal import Decimal
from datetime import datetime, timezone
from pathlib import Path
from collections import defaultdict
from typing import Optional, Dict, Any, List, Tuple

# Garantir raiz do projeto no sys.path
PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Garantir codificação UTF-8 no Windows Console
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

# Auto-reexecução no ambiente virtual .venv-rfc se pyrfc ou pandas não estiverem presentes
try:
    import pyrfc
    import pandas as pd
    import openpyxl
except ImportError:
    venv_python = os.path.join(str(PROJECT_ROOT), ".venv-rfc", "Scripts", "python.exe")
    if os.path.exists(venv_python) and sys.executable.lower() != venv_python.lower():
        import subprocess
        res = subprocess.run([venv_python, os.path.abspath(__file__)] + sys.argv[1:])
        sys.exit(res.returncode)

# Suporte a cores no terminal
try:
    import colorama
    colorama.init(autoreset=True)
    CLR_G = colorama.Fore.GREEN + colorama.Style.BRIGHT
    CLR_R = colorama.Fore.RED + colorama.Style.BRIGHT
    CLR_Y = colorama.Fore.YELLOW + colorama.Style.BRIGHT
    CLR_C = colorama.Fore.CYAN + colorama.Style.BRIGHT
    CLR_M = colorama.Fore.MAGENTA + colorama.Style.BRIGHT
    CLR_W = colorama.Fore.WHITE + colorama.Style.BRIGHT
    CLR_RST = colorama.Style.RESET_ALL
except Exception:
    CLR_G = CLR_R = CLR_Y = CLR_C = CLR_M = CLR_W = CLR_RST = ""

from sap_rfc._rfc_common import (
    build_connection_params_for, find_project_root, load_project_env,
    make_write_guard, read_table
)
from pyrfc import Connection

# Caminhos padrão do Excel
CAMINHO_EXCEL_PADRAO = str(PROJECT_ROOT / "S4H_Perfis de autorização_v1.xlsx")
CAMINHO_EXCEL_BACKUP = str(PROJECT_ROOT / "output" / "S4H_Perfis de autorização_v1_backup_automatico.xlsx")

# Mapeamento de departamentos para folhas do Excel
MAPA_DEPARTAMENTOS_SHEET = {
    "CONSTRUCTION & MAINTENANCE": "Construction & Maintenance",
    "PURCHASE & SERVICES": "Purchase & Services",
    "CLIENT SERVICES": "Client Services",
    "INDUSTRY SERVICES": "Industry Services",
    "PEOPLE & TALENT": "People & Talent",
    "P&T": "People & Talent",
    "HEALTH & SAFETY": "Health & Safety",
    "H&S": "Health & Safety",
    "DIGITAL": "Digital",
    "IT": "IT",
    "TI": "IT",
    "LEGAL": "Legal",
    "PROPOSTA FI": "Proposta FI",
    "FINANCEIRO": "Proposta FI",
    "FI": "Proposta FI",
}


def normalizar_texto(valor: Any) -> str:
    """Remove acentos, converte para maiúsculas e remove espaços redundantes."""
    if valor is None:
        return ""
    txt = str(valor).strip()
    if txt.lower() in ("nan", "none", "<na>"):
        return ""
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(ch for ch in txt if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", txt).strip().upper()


def normalizar_chave(txt: Any) -> str:
    """Normaliza removendo todos os caracteres não alfanuméricos."""
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(txt))


def converter_utc_para_local(ts_dec: Any) -> str:
    """Converte carimbo de data/hora UTC do SAP para a timezone local da máquina."""
    ts_str = str(ts_dec).replace(".", "")
    if len(ts_str) >= 14:
        try:
            dt_utc = datetime.strptime(ts_str[:14], "%Y%m%d%H%M%S").replace(tzinfo=timezone.utc)
            dt_local = dt_utc.astimezone()
            return dt_local.strftime("%d.%m.%Y %H:%M:%S")
        except Exception:
            return str(ts_dec)
    return str(ts_dec)


# =====================================================================
# 1. CONSULTA RFC DO BUFFER DE ERROS SU53
# =====================================================================

def consultar_erros_su53_rfc(usuario: str, target_env: str = "PRD") -> List[Dict[str, Any]]:
    """
    Consulta todas as falhas de autorização registadas no buffer da SU53
    (SUSR_USER_SU53_READ) para o utilizador indicado em todos os servidores aplicacionais.
    """
    load_project_env(find_project_root())
    params = build_connection_params_for(target_env)
    u_norm = usuario.strip().upper()

    print(f"\n{CLR_C}>>> Ligando ao SAP {target_env} via RFC para ler buffer da SU53 do user '{u_norm}'...{CLR_RST}")
    conn = Connection(**params)
    try:
        res = conn.call(
            "SUSR_USER_SU53_READ",
            IV_BNAME=u_norm,
            IV_ALL_SERVERS="X",
            IV_FROM=Decimal("0")
        )
    finally:
        conn.close()

    entries = res.get("ET_USR07_EXT", [])
    erros = []
    for e in entries:
        rc = int(e.get("RC", 0))
        if rc > 0:  # Apenas verificações falhadas
            ts_dec = e.get("TIMESTAMP", 0)
            dt_formatada = converter_utc_para_local(ts_dec)

            campos_valores = []
            for i in range(1, 10):
                f_name = e.get(f"FIEL{i}", "").strip()
                v_val = e.get(f"VAL{i:02d}", "").strip()
                if f_name:
                    campos_valores.append((f_name, v_val))
            if e.get("FIEL0"):
                campos_valores.append((e.get("FIEL0", "").strip(), e.get("VAL10", "").strip()))

            tcode_f = next((v for f, v in campos_valores if f == "TCD"), "")
            if not tcode_f and e.get("OBJCT") == "S_TCODE":
                tcode_f = e.get("VAL01", "").strip()

            erros.append({
                "timestamp": ts_dec,
                "data_hora": dt_formatada,
                "servidor": e.get("INSTANCE", "").strip(),
                "objeto": e.get("OBJCT", "").strip(),
                "rc": rc,
                "programa": e.get("ABAPPROG", "").strip(),
                "linha": e.get("ABAPLINE", 0),
                "tcode_app": e.get("P_TCODE", "").strip(),
                "nome_servico": e.get("NAME", "").strip(),
                "campos": campos_valores,
                "tcode_falha": tcode_f
            })

    erros.sort(key=lambda x: x["timestamp"], reverse=True)
    return erros


def obter_descricao_objeto_sap(objeto: str) -> str:
    """Obtém a descrição técnica do objeto de autorização na tabela TOBJT."""
    try:
        load_project_env(find_project_root())
        conn = Connection(**build_connection_params_for("PRD"))
        res = conn.call(
            "RFC_READ_TABLE",
            QUERY_TABLE="TOBJT",
            DELIMITER="|",
            OPTIONS=[{"TEXT": f"OBJECT = '{objeto.strip().upper()}' AND ( LANGU = 'P' OR LANGU = 'E' )"}],
            FIELDS=[{"FIELDNAME": "OBJECT"}, {"FIELDNAME": "LANGU"}, {"FIELDNAME": "TTEXT"}],
            ROWCOUNT=2
        )
        conn.close()
        for r in res.get("DATA", []):
            parts = r["WA"].split("|")
            if len(parts) >= 3 and parts[2].strip():
                return parts[2].strip()
    except Exception:
        pass
    return ""


def exibir_tabela_erros(erros: List[Dict[str, Any]], usuario: str, target_env: str = "PRD"):
    """Exibe na consola o extrato cronológico recente de falhas recolhidas da SU53."""
    env_label = "QAS (QAD)" if target_env.upper() in ("QAD", "QAS") else target_env.upper()
    print("\n" + "=" * 95)
    print(f"  {CLR_W}HISTÓRICO RECENTE DE FALHAS REGISTADAS NA SU53 - UTILIZADOR: {CLR_Y}{usuario.upper()}{CLR_W} [AMBIENTE: {CLR_M}{env_label}{CLR_W}]{CLR_RST}")
    print("=" * 95)

    if not erros:
        print(f"  {CLR_Y}[AVISO]{CLR_RST} Não foram encontrados registos de erro (RC > 0) no buffer da SU53 para '{usuario}' no ambiente {env_label}.")
        print(f"  {CLR_C}-> Nota Operacional:{CLR_RST} No SAP GUI, após ocorrer o bloqueio de autorização na transação,")
        print(f"                     é necessário executar {CLR_W}/nSU53{CLR_RST} nessa mesma sessão para que o SAP")
        print("                     descarregue a verificação falhada para o buffer acessível via RFC.")
        print("=" * 95)
        return

    print(f"{'#':<3} | {'DATA/HORA (LOCAL)':<19} | {'OBJETO':<12} | {'RC':<3} | {'DETALHES / VALORES VERIFICADOS':<48}")
    print("-" * 95)

    for idx, err in enumerate(erros[:20], 1):
        campos_txt = ", ".join([f"{f}={v}" for f, v in err["campos"]])
        if err["objeto"] == "S_TCODE" and err["tcode_falha"]:
            detalhes = f"{CLR_R}TRANSAÇÃO: {err['tcode_falha']}{CLR_RST}"
        else:
            detalhes = campos_txt[:48]

        cor_obj = CLR_R if err["objeto"] == "S_TCODE" else CLR_Y
        print(f"{idx:<3} | {err['data_hora']:<19} | {cor_obj}{err['objeto']:<12}{CLR_RST} | {err['rc']:<3} | {detalhes}")

    if len(erros) > 20:
        print(f"... e mais {len(erros) - 20} registos anteriores no buffer.")
    print("-" * 95)


# =====================================================================
# 2. SÍNTESE E AGRUPAMENTO: TRANSAÇÕES VS OBJETOS DE AUTORIZAÇÃO
# =====================================================================

def agrupar_erros_para_analise(erros: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Agrupa os erros em duas categorias distintas:
      1. Transações com erro de arranque (Objeto S_TCODE)
      2. Outros Objetos de Autorização com falha (S_USER_UID, S_SAL_LOG, S_TABU_DIS, etc.)
    """
    transacoes_map = {}
    objetos_map = defaultdict(list)

    for e in erros:
        obj = e["objeto"]
        if obj == "S_TCODE":
            tc = e["tcode_falha"]
            if tc and tc not in transacoes_map:
                transacoes_map[tc] = e
        else:
            objetos_map[obj].append(e)

    lista_transacoes = list(transacoes_map.values())

    lista_objetos = []
    for obj_name, ocorrencias in objetos_map.items():
        mais_recente = ocorrencias[0]
        # Consolidar campos testados distintos
        campos_relevantes = []
        for o in ocorrencias:
            for f, v in o["campos"]:
                if f and f != "dummyfield" and (f, v) not in campos_relevantes:
                    campos_relevantes.append((f, v))

        lista_objetos.append({
            "objeto": obj_name,
            "descricao": obter_descricao_objeto_sap(obj_name),
            "rc_mais_recente": mais_recente["rc"],
            "data_hora_recente": mais_recente["data_hora"],
            "programa": mais_recente["programa"],
            "linha": mais_recente["linha"],
            "campos_relevantes": campos_relevantes,
            "total_ocorrencias": len(ocorrencias)
        })

    return lista_transacoes, lista_objetos


def exibir_analise_sintese(
    transacoes: List[Dict[str, Any]],
    objetos: List[Dict[str, Any]],
    caminho_excel: str,
    usuario: str
) -> List[Dict[str, Any]]:
    """
    Apresenta o painel de análise com a lista combinada e numerada
    para permitir a seleção de Transações ou Objetos, exibindo imediatamente
    a pesquisa no Excel e sugestão de Função Individual (Z_...) para cada transação bloqueada.
    """
    print("\n" + "=" * 95)
    print(f"  {CLR_W}PAINEL DE ANÁLISE E SÍNTESE DE FALHAS (TRANSAÇÕES & OBJETOS DETETADOS){CLR_RST}")
    print("=" * 95)

    opcoes_selecao = []
    num_item = 1

    # SECÇÃO 1: TRANSAÇÕES COM SUGESTÃO DE FUNÇÃO INDIVIDUAL
    print(f"\n{CLR_C}[A] TRANSAÇÕES BLOQUEADAS NO ARRANQUE (OBJETO S_TCODE):{CLR_RST}")
    if transacoes:
        for t in transacoes:
            tc = t["tcode_falha"]
            print(f"  [{CLR_Y}{num_item}{CLR_RST}] {CLR_W}Transação: {CLR_Y}{tc:<20}{CLR_RST} -> RC={t['rc']} | Última: {t['data_hora']} ({t['programa']}:{t['linha']})")

            # Pesquisa imediata de sugestão de Função Individual no Excel / Catálogo
            roles_sug = buscar_funcao_individual_para_transacao(caminho_excel, tc, usuario)
            sug_role_padrao = None

            if roles_sug:
                exatos = [r for r in roles_sug if not r.get("candidata")]
                candidatos = [r for r in roles_sug if r.get("candidata")]

                if exatos:
                    r_top = exatos[0]
                    sug_role_padrao = r_top["role"]
                    outras = [r["role"] for r in exatos[1:3]]
                    outras_txt = f" (ou: {', '.join(outras)})" if outras else ""
                    print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_G}{r_top['role']}{CLR_RST} - {r_top['descricao']} [{r_top['fonte']}]{outras_txt}")
                elif candidatos:
                    r_top = candidatos[0]
                    sug_role_padrao = r_top["role"]
                    cands_txt = " / ".join([c["role"] for c in candidatos[:2]])
                    print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_Y}Não catalogada na folha 'Proposta'{CLR_RST} (Candidatas sugeridas: {CLR_W}{cands_txt}{CLR_RST} ou introduzir manual)")
            else:
                print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_Y}Não encontrada no catálogo 'Proposta'{CLR_RST} (requer indicação manual de Z_...)")

            opcoes_selecao.append({
                "tipo": "TRANSACAO",
                "valor": tc,
                "detalhes": t,
                "sugestao_role": sug_role_padrao,
                "roles_encontradas": roles_sug
            })
            num_item += 1
    else:
        print("  (Nenhuma transação S_TCODE identificada com falha no buffer recente)")

    # SECÇÃO 2: OBJETOS DE AUTORIZAÇÃO
    print(f"\n{CLR_C}[B] OUTROS OBJETOS DE AUTORIZAÇÃO COM FALHA:{CLR_RST}")
    if objetos:
        for o in objetos:
            desc = f" ({o['descricao']})" if o["descricao"] else ""
            campos_str = ", ".join([f"{f}={v}" for f, v in o["campos_relevantes"]]) if o["campos_relevantes"] else "Verificação de existência"
            print(f"  [{CLR_Y}{num_item}{CLR_RST}] {CLR_W}Objeto:    {CLR_M}{o['objeto']:<20}{CLR_RST}{desc}")
            print(f"       -> RC={o['rc_mais_recente']} | Falhas: {o['total_ocorrencias']}x | Campos: {campos_str} | Prog: {o['programa']}")
            opcoes_selecao.append({
                "tipo": "OBJETO",
                "valor": o["objeto"],
                "detalhes": o
            })
            num_item += 1
    else:
        print("  (Nenhum outro objeto de autorização com falha identificado)")

    print("-" * 95)
    print("  [O] Introduzir manualmente outra Transação ou Objeto")
    print("  [0] Cancelar e sair")
    print("-" * 95)

    return opcoes_selecao



# =====================================================================
# 3. DIAGNÓSTICO E ANÁLISE APROFUNDADA DE OBJETO DE AUTORIZAÇÃO
# =====================================================================

def diagnosticar_objeto_autorizacao(usuario: str, obj_info: Dict[str, Any], caminho_excel: str):
    """
    Executa diagnóstico do Objeto de Autorização:
      - Autorizações atuais do utilizador para esse objeto via RFC (SUSR_USER_AUTH_FOR_OBJ_GET).
      - Funções PFCG que contêm o objeto no SAP PRD (AGR_1251) ou na folha PFCG_AUTHORITY.
    """
    obj_nome = obj_info["objeto"]
    u_norm = usuario.strip().upper()

    print("\n" + "=" * 95)
    print(f"  {CLR_W}DIAGNÓSTICO TÉCNICO DO OBJETO DE AUTORIZAÇÃO: {CLR_M}{obj_nome}{CLR_RST}")
    print("=" * 95)

    desc = obj_info.get("descricao") or obter_descricao_objeto_sap(obj_nome)
    if desc:
        print(f"  Descrição SAP:  {CLR_W}{desc}{CLR_RST}")
    print(f"  Código Retorno: {CLR_R}RC = {obj_info['rc_mais_recente']}{CLR_RST} (Sem autorização suficiente)")
    print(f"  Última Falha:   {obj_info['data_hora_recente']} no programa {obj_info['programa']} (linha {obj_info['linha']})")

    if obj_info.get("campos_relevantes"):
        campos_txt = ", ".join([f"{f}={v}" for f, v in obj_info["campos_relevantes"]])
        print(f"  Campos/Valores: {CLR_Y}{campos_txt}{CLR_RST}")

    # 1. Consultar se o utilizador tem autorizações atuais para este objeto
    print(f"\n{CLR_C}1. A verificar buffer atual do utilizador '{u_norm}' para {obj_nome}...{CLR_RST}")
    try:
        load_project_env(find_project_root())
        conn = Connection(**build_connection_params_for("PRD"))
        res_auth = conn.call("SUSR_USER_AUTH_FOR_OBJ_GET", USER_NAME=u_norm, SEL_OBJECT=obj_nome)
        conn.close()
        valores = res_auth.get("VALUES", [])
        if valores:
            print(f"  {CLR_G}[OK] O utilizador possui {len(valores)} registo(s) de autorização no buffer:{CLR_RST}")
            for v in valores[:5]:
                print(f"     - Perfil: {v.get('AUTH')} | Campo: {v.get('FIELD')} | De: {v.get('VON')} Até: {v.get('BIS')}")
        else:
            print(f"  {CLR_R}[ATENÇÃO] O utilizador '{u_norm}' tem 0 autorizações atribuídas para o objeto '{obj_nome}'.{CLR_RST}")
    except Exception as exc:
        print(f"  {CLR_Y}[AVISO] Objeto não configurado no perfil atual: {exc}{CLR_RST}")

    # 2. Pesquisar roles no SAP PRD (AGR_1251) que contêm o objeto
    print(f"\n{CLR_C}2. A pesquisar funções PFCG no SAP PRD que concedem o objeto {obj_nome}...{CLR_RST}")
    try:
        load_project_env(find_project_root())
        conn = Connection(**build_connection_params_for("PRD"))
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_1251", "AGR_FLAGS"))
        rows_agr = read_table(
            conn, guard, table_name="AGR_1251", fields=["AGR_NAME", "FIELD", "LOW"],
            options=[{"TEXT": f"OBJECT = '{obj_nome}'"}],
            rowcount=10
        )
        conn.close()
        roles_distintas = sorted(list({r[0].strip().upper() for r in rows_agr if r and r[0].strip().startswith("Z_")}))
        if roles_distintas:
            print(f"  {CLR_G}Funções Z encontradas no SAP PRD com o objeto {obj_nome}:{CLR_RST}")
            for r in roles_distintas[:8]:
                print(f"    - {CLR_Y}{r}{CLR_RST}")
        else:
            print(f"  (Nenhuma função Z direta identificada no PRD com filtros básicos)")
    except Exception:
        pass

    # 3. Pesquisar na folha PFCG_AUTHORITY do Excel
    print(f"\n{CLR_C}3. A consultar folha 'PFCG_AUTHORITY' no Excel...{CLR_RST}")
    try:
        df_auth = pd.read_excel(caminho_excel, sheet_name="PFCG_AUTHORITY")
        col_obj = next((c for c in df_auth.columns if "OBJETO" in str(c).upper()), None)
        if col_obj:
            m_auth = df_auth[df_auth[col_obj].astype(str).str.strip().str.upper() == obj_nome]
            if not m_auth.empty:
                print(f"  {CLR_G}Encontrados {len(m_auth)} registo(s) na folha 'PFCG_AUTHORITY':{CLR_RST}")
                for _, row_a in m_auth.head(5).iterrows():
                    comp = row_a.get("AGR_NAME_COMPOSTA", "")
                    sgl = row_a.get("AGR_NAME", "")
                    actvt = row_a.get("ACTVT", "")
                    print(f"    - Composta: {comp} | Individual: {sgl} | ACTVT: {actvt}")
            else:
                print(f"  (O objeto '{obj_nome}' não se encontra atualmente mapeado na folha PFCG_AUTHORITY)")
    except Exception as exc:
        print(f"  (Erro ao ler folha PFCG_AUTHORITY: {exc})")

    print("\n" + "-" * 95)
    print(f"  {CLR_Y}[NOTA INFORMATIVA]{CLR_RST} A análise técnica do objeto {CLR_M}{obj_nome}{CLR_RST} foi concluída.")
    print("      A lógica de atribuição automatizada de campos/objetos de autorização adicionais")
    print("      em 'PFCG_AUTHORITY' será estruturada e ativada na próxima etapa.")
    print("-" * 95)


# =====================================================================
# 4. ANÁLISE NO FICHEIRO EXCEL (PROPOSTA ATIVA, DEPARTAMENTO, PROPOSTA)
# =====================================================================

def obter_dados_utilizador_proposta_ativa(caminho_excel: str, usuario: str) -> Optional[Dict[str, Any]]:
    """Localiza o utilizador na folha 'Proposta Ativa' e devolve os seus metadados."""
    u_norm = usuario.strip().upper()
    df_ativa = pd.read_excel(caminho_excel, sheet_name="Proposta Ativa")

    col_user = [c for c in df_ativa.columns if any(k in str(c).upper() for k in ("USER", "UTILIZADOR", "USU"))][0]
    col_comp = [c for c in df_ativa.columns if "COMPOSITE" in str(c).upper()][0]
    col_dep = [c for c in df_ativa.columns if str(c).strip().upper() == "DEPARTAMENTO"][0]
    col_dep_dir = [c for c in df_ativa.columns if "DIRE" in str(c).upper()][0]
    col_nome = [c for c in df_ativa.columns if any(k in str(c).upper() for k in ("NOME", "COLABORADOR"))][0]

    for idx, row in df_ativa.iterrows():
        u_val = str(row.get(col_user, "")).strip().upper()
        if u_val == u_norm:
            d_val = str(row.get(col_dep, "")).strip() if pd.notna(row.get(col_dep)) else ""
            dd_val = str(row.get(col_dep_dir, "")).strip() if pd.notna(row.get(col_dep_dir)) else ""
            c_val = str(row.get(col_comp, "")).strip() if pd.notna(row.get(col_comp)) else ""
            nome_val = str(row.get(col_nome, "")).strip() if pd.notna(row.get(col_nome)) else ""

            dep = d_val if d_val and d_val.lower() != "nan" else dd_val
            comp = c_val if c_val and c_val.lower() != "nan" else ""

            funcoes_existentes = []
            for col_idx in range(10, len(row)):
                cell_val = row.iloc[col_idx]
                if pd.notna(cell_val):
                    v_str = str(cell_val).strip().upper()
                    if v_str.startswith("Z") and v_str != comp and v_str not in funcoes_existentes:
                        funcoes_existentes.append(v_str)

            return {
                "linha_excel": idx + 2,
                "usuario": u_norm,
                "nome": nome_val,
                "departamento": dep,
                "composite_role": comp,
                "funcoes_existentes": funcoes_existentes
            }

    return None


def buscar_funcao_individual_para_transacao(
    caminho_excel: str,
    tcode: str,
    usuario: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Pesquisa a Função Individual (Single Role Z_...) correspondente à transação:
      1. Folha 'Proposta' (catálogo principal, coluna 'FUNÇÃO')
      2. Folha 'Single Roles (S4H)'
      3. Folha 'PFCG_CREATE'
      4. SAP PRD via RFC (tabelas AGR_1251 e AGR_TCODES)
      5. Se não existir no catálogo, identifica funções candidatas no PRD por prefixo ou funções do utilizador.
    """
    tc_norm = tcode.strip().upper()
    funcoes_encontradas = []

    # 1. Pesquisa na folha 'Proposta' (fonte primordial do catálogo)
    try:
        df_prop = pd.read_excel(caminho_excel, sheet_name="Proposta")
        current_role = None
        current_desc = ""

        for _, row in df_prop.iterrows():
            f_val = str(row.get("FUNÇÃO", "")).strip() if pd.notna(row.get("FUNÇÃO")) else ""
            d_val = str(row.get("DESCRIÇÃO", "")).strip() if pd.notna(row.get("DESCRIÇÃO")) else ""

            if f_val.startswith("Z_") and d_val and "TRANSACAO NAO EXISTE" not in d_val.upper():
                current_role = f_val
                current_desc = d_val
            elif f_val and current_role:
                for t in f_val.replace(";", " ").split():
                    if t.strip().upper() == tc_norm:
                        if not any(f["role"] == current_role for f in funcoes_encontradas):
                            funcoes_encontradas.append({
                                "role": current_role,
                                "descricao": current_desc,
                                "fonte": "Folha 'Proposta'",
                                "candidata": False
                            })
    except Exception:
        pass

    # 2. Pesquisa na folha 'Single Roles (S4H)'
    try:
        df_sgl = pd.read_excel(caminho_excel, sheet_name="Single Roles (S4H)")
        cur_sgl_role = None
        cur_sgl_desc = ""
        for _, row in df_sgl.iterrows():
            c1 = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            c2 = str(row.iloc[2]).strip().upper() if pd.notna(row.iloc[2]) else ""
            c3 = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
            if c1.startswith("Z"):
                cur_sgl_role = c1
                cur_sgl_desc = c3
            elif c2 == tc_norm and cur_sgl_role:
                if not any(f["role"] == cur_sgl_role for f in funcoes_encontradas):
                    funcoes_encontradas.append({
                        "role": cur_sgl_role,
                        "descricao": c3 or cur_sgl_desc or f"Função Single {cur_sgl_role}",
                        "fonte": "Folha 'Single Roles (S4H)'",
                        "candidata": False
                    })
    except Exception:
        pass

    # 3. Pesquisa na folha 'PFCG_CREATE'
    try:
        df_create = pd.read_excel(caminho_excel, sheet_name="PFCG_CREATE")
        for _, row in df_create.iterrows():
            agr = str(row.get("AGR_NAME", "")).strip().upper()
            tc_col = str(row.get("TCODE", "")).strip().upper()
            desc_col = str(row.get("TEXT", "")).strip()
            if agr.startswith("Z_") and tc_norm in [t.strip().upper() for t in tc_col.replace(";", " ").split()]:
                if not any(f["role"] == agr for f in funcoes_encontradas):
                    funcoes_encontradas.append({
                        "role": agr,
                        "descricao": desc_col,
                        "fonte": "Folha 'PFCG_CREATE'",
                        "candidata": False
                    })
    except Exception:
        pass

    # 4. Pesquisa no SAP PRD via RFC (tabelas AGR_1251 e AGR_TCODES)
    if not funcoes_encontradas:
        try:
            load_project_env(find_project_root())
            conn = Connection(**build_connection_params_for("PRD"))
            guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_1251", "AGR_FLAGS", "AGR_USERS", "AGR_TCODES"))

            # AGR_1251
            rows = read_table(
                conn, guard, table_name="AGR_1251", fields=["AGR_NAME"],
                options=[{"TEXT": f"OBJECT = 'S_TCODE' AND LOW = '{tc_norm}'"}],
                rowcount=10
            )
            for r in rows:
                r_name = r[0].strip().upper()
                if r_name.startswith("Z_") and not any(f["role"] == r_name for f in funcoes_encontradas):
                    funcoes_encontradas.append({
                        "role": r_name,
                        "descricao": f"Função SAP PRD com {tc_norm}",
                        "fonte": "SAP PRD (AGR_1251)",
                        "candidata": False
                    })

            # AGR_TCODES
            if not funcoes_encontradas:
                rows_tc = read_table(
                    conn, guard, table_name="AGR_TCODES", fields=["AGR_NAME"],
                    options=[{"TEXT": f"TCODE = '{tc_norm}'"}],
                    rowcount=10
                )
                for r in rows_tc:
                    r_name = r[0].strip().upper()
                    if r_name.startswith("Z_") and not any(f["role"] == r_name for f in funcoes_encontradas):
                        funcoes_encontradas.append({
                            "role": r_name,
                            "descricao": f"Função SAP PRD com {tc_norm}",
                            "fonte": "SAP PRD (AGR_TCODES)",
                            "candidata": False
                        })

            # 5. Se não encontrado no catálogo, identificar funções candidatas (por utilizador ou prefixo)
            if not funcoes_encontradas:
                candidatos = []
                if usuario:
                    u_rows = read_table(
                        conn, guard, table_name="AGR_USERS", fields=["AGR_NAME"],
                        options=[{"TEXT": f"UNAME = '{usuario.strip().upper()}'"}],
                        rowcount=50
                    )
                    u_roles = [ur[0].strip().upper() for ur in u_rows if ur and ur[0].strip().startswith("Z")]
                    for ur in u_roles:
                        prefix_tc = tc_norm[:4]
                        if prefix_tc in ur or ("BASIS" in ur and "PFCG" in tc_norm) or ("IT" in ur and "PFCG" in tc_norm):
                            candidatos.append({
                                "role": ur,
                                "descricao": f"Função existente do utilizador {usuario.strip().upper()}",
                                "fonte": "SAP PRD (Atribuída ao Utilizador)",
                                "candidata": True
                            })

                if not candidatos and len(tc_norm) >= 4:
                    pref = tc_norm[:4]
                    p_rows = read_table(
                        conn, guard, table_name="AGR_1251", fields=["AGR_NAME"],
                        options=[{"TEXT": f"OBJECT = 'S_TCODE' AND LOW LIKE '{pref}%'"}],
                        rowcount=10
                    )
                    for pr in p_rows:
                        pr_name = pr[0].strip().upper()
                        if pr_name.startswith("Z_") and not any(c["role"] == pr_name for c in candidatos):
                            candidatos.append({
                                "role": pr_name,
                                "descricao": f"Função com transações do prefixo {pref}*",
                                "fonte": "SAP PRD (Candidata)",
                                "candidata": True
                            })

                # Ordenar candidatos priorizando roles cujo nome contenha o prefixo da transação
                pref_tc = tc_norm[:4]
                candidatos.sort(key=lambda c: (0 if pref_tc in c["role"] else (1 if "BASIS" in c["role"] else 2), c["role"]))
                funcoes_encontradas.extend(candidatos)

            conn.close()
        except Exception:
            pass

    return funcoes_encontradas


def resolver_nome_sheet_departamento(caminho_excel: str, departamento: str) -> Optional[str]:
    """Identifica o nome real da folha do departamento no ficheiro Excel."""
    dep_norm = normalizar_texto(departamento)
    if dep_norm in MAPA_DEPARTAMENTOS_SHEET:
        return MAPA_DEPARTAMENTOS_SHEET[dep_norm]

    wb = openpyxl.load_workbook(caminho_excel, read_only=True)
    sheet_names = wb.sheetnames
    wb.close()

    dep_clean = normalizar_chave(departamento)
    for s in sheet_names:
        if normalizar_chave(s) == dep_clean:
            return s
        if dep_clean and dep_clean in normalizar_chave(s):
            return s

    return None


# =====================================================================
# 5. ATUALIZAÇÃO INTEGRADA DAS FOLHAS EXCEL
# =====================================================================

def aplicar_alteracoes_excel(
    caminho_excel: str,
    dados_user: Dict[str, Any],
    tcode: str,
    single_role: str,
    nome_sheet_dep: str
) -> Dict[str, Any]:
    """Aplica as alterações no Excel (folha departamental, Proposta Ativa e PFCG_COMPOSTA)."""
    comp_role = dados_user["composite_role"]
    u_norm = dados_user["usuario"]
    ts_agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    sucesso_com = False
    wb_aberto_com = None

    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            for w in xl.Workbooks:
                if os.path.basename(caminho_excel).lower() in w.Name.lower():
                    wb_aberto_com = w
                    break
        except Exception:
            wb_aberto_com = None

    # --- MÉTODO 1: EXCEL COM (Trabalha diretamente na folha aberta) ---
    if wb_aberto_com is not None:
        try:
            print(f"  {CLR_C}[Excel COM] A aplicar alterações no livro aberto...{CLR_RST}")
            wb = wb_aberto_com

            # ETAPA 1: Folha Departamental
            ws_dep = wb.Worksheets(nome_sheet_dep)
            max_r = ws_dep.UsedRange.Rows.Count
            max_c = ws_dep.UsedRange.Columns.Count

            col_user_idx = None
            hdr_row = 1
            for r in range(1, 4):
                for c in range(1, max_c + 1):
                    val = str(ws_dep.Cells(r, c).Value or "").strip().upper()
                    if u_norm in val:
                        col_user_idx = c
                        hdr_row = r
                        break
                if col_user_idx is not None:
                    break

            if col_user_idx is None:
                col_user_idx = max_c + 1
                header_text = f"{u_norm}\n{dados_user.get('nome', '')}".strip()
                ws_dep.Cells(hdr_row, col_user_idx).Value = header_text
                print(f"    + Adicionada coluna do utilizador {u_norm} na folha '{nome_sheet_dep}' (Col {col_user_idx})")

            row_tcode_idx = None
            for r in range(hdr_row + 1, max_r + 2):
                val_tc = str(ws_dep.Cells(r, 1).Value or "").strip().upper()
                if val_tc == tcode.strip().upper():
                    row_tcode_idx = r
                    break

            if row_tcode_idx is None:
                row_tcode_idx = max_r + 1
                ws_dep.Cells(row_tcode_idx, 1).Value = tcode.strip().upper()
                print(f"    + Adicionada nova linha para transação {tcode} na folha '{nome_sheet_dep}' (Linha {row_tcode_idx})")

            ws_dep.Cells(row_tcode_idx, col_user_idx).Value = "X"
            print(f"    + Marcado 'X' para {u_norm} na transação {tcode} ({nome_sheet_dep})")

            # ETAPA 2: Proposta Ativa
            ws_ativa = wb.Worksheets("Proposta Ativa")
            row_user = dados_user["linha_excel"]
            if not ws_ativa.Cells(row_user, 1).Value:
                ws_ativa.Cells(row_user, 1).Value = u_norm
                ws_ativa.Cells(row_user, 3).Value = dados_user.get("nome", "")
                ws_ativa.Cells(row_user, 6).Value = dados_user.get("departamento", "")
                ws_ativa.Cells(row_user, 8).Value = dados_user.get("departamento", "")
                ws_ativa.Cells(row_user, 9).Value = comp_role

            col_target = 11
            while ws_ativa.Cells(row_user, col_target).Value:
                curr_val = str(ws_ativa.Cells(row_user, col_target).Value).strip().upper()
                if curr_val == single_role.strip().upper():
                    col_target = None
                    break
                col_target += 1

            if col_target is not None:
                ws_ativa.Cells(row_user, col_target).Value = single_role.strip().upper()
                print(f"    + Adicionada função '{single_role}' ao utilizador na 'Proposta Ativa' (Col {col_target})")
            else:
                print(f"    . Função '{single_role}' já constava na linha do utilizador na 'Proposta Ativa'.")

            # ETAPA 3: PFCG_COMPOSTA
            ws_comp = wb.Worksheets("PFCG_COMPOSTA")
            last_r_comp = ws_comp.UsedRange.Rows.Count
            while ws_comp.Cells(last_r_comp, 2).Value:
                last_r_comp += 1

            max_id = 0
            desc_composta = ""
            par_ja_existe = False

            for r in range(2, last_r_comp):
                rid_val = ws_comp.Cells(r, 1).Value
                try:
                    rid = int(rid_val)
                    if rid > max_id:
                        max_id = rid
                except Exception:
                    pass
                comp_c = str(ws_comp.Cells(r, 2).Value or "").strip().upper()
                role_c = str(ws_comp.Cells(r, 4).Value or "").strip().upper()
                txt_c = str(ws_comp.Cells(r, 3).Value or "").strip()
                if comp_c == comp_role.strip().upper():
                    if txt_c and not desc_composta:
                        desc_composta = txt_c
                    if role_c == single_role.strip().upper():
                        par_ja_existe = True

            if not par_ja_existe:
                new_r = last_r_comp
                ws_comp.Cells(new_r, 1).Value = max_id + 1
                ws_comp.Cells(new_r, 2).Value = comp_role.strip().upper()
                ws_comp.Cells(new_r, 3).Value = desc_composta
                ws_comp.Cells(new_r, 4).Value = single_role.strip().upper()
                ws_comp.Cells(new_r, 5).Value = "Criado"
                ws_comp.Cells(new_r, 6).Value = "Atribuído em SAP DEV, PRD e QAD"
                ws_comp.Cells(new_r, 7).Value = ts_agora
                ws_comp.Cells(new_r, 8).Value = "Validado"
                ws_comp.Cells(new_r, 9).Value = "Validado"
                print(f"    + Atribuída relação {comp_role} -> {single_role} na folha 'PFCG_COMPOSTA' (ID {max_id + 1})")
            else:
                print(f"    . Relação {comp_role} -> {single_role} já constava na folha 'PFCG_COMPOSTA'.")

            wb.Save()
            sucesso_com = True
        except Exception as exc:
            print(f"  {CLR_Y}[AVISO] Erro na gravação via COM: {exc}. A tentar via openpyxl...{CLR_RST}")
            sucesso_com = False

    # --- MÉTODO 2: OPENPYXL (Se o Excel estiver fechado) ---
    if not sucesso_com:
        try:
            print(f"  {CLR_C}[openpyxl] A gravar alterações no ficheiro Excel...{CLR_RST}")
            wb = openpyxl.load_workbook(caminho_excel)

            # ETAPA 1: Folha Departamental
            ws_dep = wb[nome_sheet_dep]
            col_user_idx = None
            hdr_row = 1
            for r in range(1, 4):
                for c in range(1, ws_dep.max_column + 1):
                    val = str(ws_dep.cell(row=r, column=c).value or "").strip().upper()
                    if u_norm in val:
                        col_user_idx = c
                        hdr_row = r
                        break
                if col_user_idx is not None:
                    break

            if col_user_idx is None:
                col_user_idx = ws_dep.max_column + 1
                ws_dep.cell(row=hdr_row, column=col_user_idx, value=f"{u_norm}\n{dados_user.get('nome', '')}".strip())
                print(f"    + Adicionada coluna do utilizador {u_norm} na folha '{nome_sheet_dep}' (Col {col_user_idx})")

            row_tcode_idx = None
            for r in range(hdr_row + 1, ws_dep.max_row + 2):
                val_tc = str(ws_dep.cell(row=r, column=1).value or "").strip().upper()
                if val_tc == tcode.strip().upper():
                    row_tcode_idx = r
                    break

            if row_tcode_idx is None:
                row_tcode_idx = ws_dep.max_row + 1
                ws_dep.cell(row=row_tcode_idx, column=1, value=tcode.strip().upper())
                print(f"    + Adicionada nova linha para transação {tcode} na folha '{nome_sheet_dep}' (Linha {row_tcode_idx})")

            ws_dep.cell(row=row_tcode_idx, column=col_user_idx, value="X")
            print(f"    + Marcado 'X' para {u_norm} na transação {tcode} ({nome_sheet_dep})")

            # ETAPA 2: Proposta Ativa
            ws_ativa = wb["Proposta Ativa"]
            row_user = dados_user["linha_excel"]
            if not ws_ativa.cell(row=row_user, column=1).value:
                ws_ativa.cell(row=row_user, column=1, value=u_norm)
                ws_ativa.cell(row=row_user, column=3, value=dados_user.get("nome", ""))
                ws_ativa.cell(row=row_user, column=6, value=dados_user.get("departamento", ""))
                ws_ativa.cell(row=row_user, column=8, value=dados_user.get("departamento", ""))
                ws_ativa.cell(row=row_user, column=9, value=comp_role)

            col_target = 11
            while ws_ativa.cell(row=row_user, column=col_target).value:
                curr_val = str(ws_ativa.cell(row=row_user, column=col_target).value).strip().upper()
                if curr_val == single_role.strip().upper():
                    col_target = None
                    break
                col_target += 1

            if col_target is not None:
                ws_ativa.cell(row=row_user, column=col_target, value=single_role.strip().upper())
                print(f"    + Adicionada função '{single_role}' ao utilizador na 'Proposta Ativa' (Col {col_target})")
            else:
                print(f"    . Função '{single_role}' já constava na linha do utilizador na 'Proposta Ativa'.")

            # ETAPA 3: PFCG_COMPOSTA
            ws_comp = wb["PFCG_COMPOSTA"]
            max_id = 0
            desc_composta = ""
            par_ja_existe = False

            for r in range(2, ws_comp.max_row + 1):
                rid_val = ws_comp.cell(row=r, column=1).value
                try:
                    rid = int(rid_val)
                    if rid > max_id:
                        max_id = rid
                except Exception:
                    pass
                comp_c = str(ws_comp.cell(row=r, column=2).value or "").strip().upper()
                role_c = str(ws_comp.cell(row=r, column=4).value or "").strip().upper()
                txt_c = str(ws_comp.cell(row=r, column=3).value or "").strip()
                if comp_c == comp_role.strip().upper():
                    if txt_c and not desc_composta:
                        desc_composta = txt_c
                    if role_c == single_role.strip().upper():
                        par_ja_existe = True

            if not par_ja_existe:
                ws_comp.append([
                    max_id + 1, comp_role.strip().upper(), desc_composta,
                    single_role.strip().upper(), "Criado", "Atribuído em SAP DEV, PRD e QAD",
                    ts_agora, "Validado", "Validado", ""
                ])
                print(f"    + Atribuída relação {comp_role} -> {single_role} na folha 'PFCG_COMPOSTA' (ID {max_id + 1})")
            else:
                print(f"    . Relação {comp_role} -> {single_role} já constava na folha 'PFCG_COMPOSTA'.")

            wb.save(caminho_excel)
            wb.close()
            try:
                shutil.copy2(caminho_excel, CAMINHO_EXCEL_BACKUP)
            except Exception:
                pass
        except Exception as exc:
            return {"ok": False, "erro": f"Falha na gravação do Excel: {exc}"}

    return {"ok": True}


# =====================================================================
# 6. ATUALIZAÇÃO VIA RFC NO SAP (PRD E QAS / QAD)
# =====================================================================
# 6. ATUALIZAÇÃO VIA RFC NO SAP (PRD E QAS / QAD)
# =====================================================================

def sincronizar_sap_rfc(
    env: str,
    composite_role: str,
    single_role: str,
    usuario: Optional[str] = None
) -> Dict[str, Any]:
    """
    Atribui a Função Individual à Composta via RFC e executa o ajuste de utilizadores.
    Se a função composta ainda não existir no SAP, atribui a função individual diretamente
    ao utilizador no SAP e executa a comparação de utilizadores (User Comparison).
    """
    env_norm = env.strip().upper()
    comp_norm = composite_role.strip().upper()
    sgl_norm = single_role.strip().upper()
    u_norm = usuario.strip().upper() if usuario else ""

    print(f"\n{CLR_C}>>> [SAP {env_norm}] A sincronizar autorizações via RFC...{CLR_RST}")

    try:
        load_project_env(find_project_root())
        params = build_connection_params_for(env_norm)
        conn = Connection(**params)
    except Exception as exc:
        return {"ok": False, "ambiente": env_norm, "erro": f"Não foi possível conectar ao SAP {env_norm}: {exc}"}

    try:
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_DEFINE", "AGR_USERS"))

        # 1. Verificar se a função individual existe no SAP
        try:
            r_sgl = read_table(conn, guard, table_name="AGR_DEFINE", fields=["AGR_NAME"], options=[{"TEXT": f"AGR_NAME = '{sgl_norm}'"}], rowcount=1)
            if not r_sgl:
                print(f"    {CLR_Y}[AVISO] A função individual '{sgl_norm}' não se encontra na tabela AGR_DEFINE do SAP {env_norm}.{CLR_RST}")
        except Exception:
            pass

        # 2. Verificar se a função composta existe no SAP
        comp_existe = False
        if comp_norm:
            try:
                r_comp = read_table(conn, guard, table_name="AGR_DEFINE", fields=["AGR_NAME"], options=[{"TEXT": f"AGR_NAME = '{comp_norm}'"}], rowcount=1)
                comp_existe = bool(r_comp)
            except Exception:
                comp_existe = False

        sucesso_comp = False
        if comp_existe:
            print(f"    1. Composite Role '{comp_norm}' identificada no SAP {env_norm}.")
            print(f"       A chamar PRGN_RFC_ADD_AGRS_TO_COLL_AGR para associar {sgl_norm}...")
            try:
                res_add = conn.call(
                    "PRGN_RFC_ADD_AGRS_TO_COLL_AGR",
                    ACTIVITY_GROUP=comp_norm,
                    CHECK_NAMESPACE="X",
                    ENQUEUE="X",
                    NO_DIALOG="X",
                    PROFILE_COMPARISON="X",
                    REC_PERS_DATA="X",
                    REC_PROF_DATA="X",
                    REC_SINGLE_ROLES="X",
                    REQUEST="",
                    ACTIVITY_GROUPS=[{"AGR_NAME": sgl_norm, "TEXT": ""}]
                )
                ret_add = res_add.get("RETURN", []) or []
                erros_add = [r for r in ret_add if str(r.get("TYPE", "")).upper() in ("E", "A")]
                if erros_add:
                    msg_e = "; ".join([str(e.get("MESSAGE", "")) for e in erros_add])
                    print(f"       {CLR_R}[ERRO] PRGN_RFC_ADD_AGRS_TO_COLL_AGR: {msg_e}{CLR_RST}")
                else:
                    sucesso_comp = True
                    print(f"       {CLR_G}[OK] Função {sgl_norm} associada à composta {comp_norm}!{CLR_RST}")

                print(f"    2. A executar Comparação de Utilizadores (User Comparison) no {env_norm}...")
                try:
                    conn.call("PRGN_GEN_PROFILES_FOR_ROLES", IV_USERCOMPARE="X", IT_ROLES=[{"AGR_NAME": comp_norm}, {"AGR_NAME": sgl_norm}])
                    print(f"       {CLR_G}[OK] Comparação de utilizadores (Ajustar Utilizadores) concluída com sucesso!{CLR_RST}")
                except Exception as exc_gen:
                    print(f"       {CLR_Y}[AVISO] Comparação de utilizadores: {exc_gen}{CLR_RST}")

            except Exception as exc_add:
                print(f"       {CLR_Y}[INFO] {comp_norm} no SAP não é uma Collective Role ({exc_add}). Prosseguindo com atribuição direta ao utilizador...{CLR_RST}")
                sucesso_comp = False

        else:
            if comp_norm:
                print(f"    {CLR_Y}[INFO] A Composite Role '{comp_norm}' ainda não existe no SAP {env_norm} (está registada no Excel para criação em PFCG).{CLR_RST}")

        if not sucesso_comp and u_norm:
            print(f"    2. A atribuir a função individual '{sgl_norm}' diretamente ao utilizador '{u_norm}' no SAP {env_norm}...")
            try:
                rows_u = read_table(conn, guard, table_name="AGR_USERS", fields=["AGR_NAME", "FROM_DAT", "TO_DAT"], options=[{"TEXT": f"UNAME = '{u_norm}'"}], rowcount=50)
                actgroups = []
                for ru in rows_u:
                    actgroups.append({
                        "AGR_NAME": ru[0].strip(),
                        "FROM_DAT": ru[1].strip() or "20260101",
                        "TO_DAT": ru[2].strip() or "99991231"
                    })
                if not any(a["AGR_NAME"] == sgl_norm for a in actgroups):
                    actgroups.append({
                        "AGR_NAME": sgl_norm,
                        "FROM_DAT": datetime.now().strftime("%Y%m%d"),
                        "TO_DAT": "99991231"
                    })
                res_bapi = conn.call("BAPI_USER_ACTGROUPS_ASSIGN", USERNAME=u_norm, ACTIVITYGROUPS=actgroups)
                ret_bapi = res_bapi.get("RETURN", []) or []
                erros_bapi = [r for r in ret_bapi if str(r.get("TYPE", "")).upper() in ("E", "A")]
                if erros_bapi:
                    msg_eb = "; ".join([str(e.get("MESSAGE", "")) for e in erros_bapi])
                    print(f"       {CLR_R}[ERRO BAPI] {msg_eb}{CLR_RST}")
                else:
                    print(f"       {CLR_G}[OK] Função {sgl_norm} atribuída com sucesso ao utilizador {u_norm} no SAP {env_norm}!{CLR_RST}")
                    try:
                        conn.call("PRGN_GEN_PROFILES_FOR_ROLES", IV_USERCOMPARE="X", IT_ROLES=[{"AGR_NAME": sgl_norm}])
                        print(f"       {CLR_G}[OK] Ajustar Utilizadores (User Comparison) concluído com sucesso!{CLR_RST}")
                    except Exception:
                        pass
            except Exception as exc_bapi:
                print(f"       {CLR_Y}[AVISO] Atribuição direta ao utilizador: {exc_bapi}{CLR_RST}")

    finally:
        conn.close()

    return {"ok": True, "ambiente": env_norm}


# =====================================================================
# 7. FLUXO DE ATRIBUIÇÃO PARA UMA TRANSAÇÃO
# =====================================================================

def executar_atribuicao_transacao(
    caminho_excel: str,
    usuario: str,
    tcode_escolhido: str,
    sugestao_role: Optional[str] = None,
    roles_previamente_encontradas: Optional[List[Dict[str, Any]]] = None,
    target_env: str = "PRD"
):
    """Executa o ciclo completo de atribuição para uma transação S_TCODE."""
    print(f"\n{CLR_W}>>> Transação selecionada para tratamento: {CLR_Y}{tcode_escolhido}{CLR_RST}")

    dados_user = obter_dados_utilizador_proposta_ativa(caminho_excel, usuario)
    if not dados_user:
        print(f"\n{CLR_Y}[AVISO] O utilizador '{usuario}' não se encontra registado na folha 'Proposta Ativa'.{CLR_RST}")
        nome_sap = ""
        try:
            load_project_env(find_project_root())
            conn_u = Connection(**build_connection_params_for("QAD"))
            res_u = conn_u.call("BAPI_USER_GET_DETAIL", USERNAME=usuario.strip().upper())
            conn_u.close()
            addr_u = res_u.get("ADDRESS", {})
            nome_sap = addr_u.get("FULLNAME") or f"{addr_u.get('NAME_FIRST', '')} {addr_u.get('NAME_LAST', '')}".strip()
        except Exception:
            pass

        if not nome_sap:
            try:
                conn_u2 = Connection(**build_connection_params_for("PRD"))
                res_u2 = conn_u2.call("BAPI_USER_GET_DETAIL", USERNAME=usuario.strip().upper())
                conn_u2.close()
                addr_u2 = res_u2.get("ADDRESS", {})
                nome_sap = addr_u2.get("FULLNAME") or f"{addr_u2.get('NAME_FIRST', '')} {addr_u2.get('NAME_LAST', '')}".strip()
            except Exception:
                pass

        if not nome_sap:
            nome_sap = usuario.strip().upper()

        print(f"  Nome identificado no SAP: {CLR_W}{nome_sap}{CLR_RST}")
        resp_add = input(f"Deseja registar o utilizador '{usuario}' na folha 'Proposta Ativa'? [S/N] [S]: ").strip().upper()
        if resp_add in ("", "S", "SIM", "Y", "YES"):
            print("Departamentos disponíveis:")
            deps_lista = sorted(list(set(MAPA_DEPARTAMENTOS_SHEET.values())))
            for i, d in enumerate(deps_lista, 1):
                print(f"  [{i}] {d}")
            sel_d = input(f"Selecione o departamento correspondente [Padrão: IT]: ").strip()
            if sel_d.isdigit() and 1 <= int(sel_d) <= len(deps_lista):
                departamento_escolhido = deps_lista[int(sel_d) - 1]
            else:
                departamento_escolhido = "IT"

            comp_input = input("Introduza o nome da Composite Role [ENTER para Z_BR_IT_SPECIALIST]: ").strip().upper() or "Z_BR_IT_SPECIALIST"

            df_ativa = pd.read_excel(caminho_excel, sheet_name="Proposta Ativa")
            dados_user = {
                "linha_excel": len(df_ativa) + 2,
                "usuario": usuario.strip().upper(),
                "nome": nome_sap,
                "departamento": departamento_escolhido,
                "composite_role": comp_input,
                "funcoes_existentes": []
            }
        else:
            print("Operação cancelada. Não é possível prosseguir sem os metadados do utilizador.")
            return

    print("\n" + "-" * 75)
    print(f"  {CLR_W}DADOS DO UTILIZADOR IDENTIFICADOS NA 'PROPOSTA ATIVA':{CLR_RST}")
    print(f"  Utilizador:     {CLR_Y}{dados_user['usuario']}{CLR_RST} ({dados_user['nome']})")
    print(f"  Departamento:   {CLR_Y}{dados_user['departamento'] or 'Não definido'}{CLR_RST}")
    print(f"  Composite Role: {CLR_Y}{dados_user['composite_role'] or 'Não atribuída'}{CLR_RST}")
    print(f"  Linha no Excel: Linha {dados_user['linha_excel']}")
    print("-" * 75)

    # Validação do Departamento
    departamento = dados_user["departamento"]
    if not departamento:
        print(f"{CLR_Y}[AVISO] O utilizador não tem Departamento definido na 'Proposta Ativa'.{CLR_RST}")
        print("Departamentos disponíveis:")
        deps_lista = sorted(list(set(MAPA_DEPARTAMENTOS_SHEET.values())))
        for i, d in enumerate(deps_lista, 1):
            print(f"  [{i}] {d}")
        sel_d = input(f"Selecione o departamento correspondente [1-{len(deps_lista)}]: ").strip()
        try:
            departamento = deps_lista[int(sel_d) - 1]
            dados_user["departamento"] = departamento
        except Exception:
            print("Departamento inválido. Operação abortada.")
            return

    nome_sheet_dep = resolver_nome_sheet_departamento(caminho_excel, departamento)
    if not nome_sheet_dep:
        print(f"{CLR_R}[ERRO] Folha departamental correspondente a '{departamento}' não encontrada no Excel.{CLR_RST}")
        return

    # Validação da Composite Role
    comp_role = dados_user["composite_role"]
    if not comp_role:
        print(f"{CLR_Y}[AVISO] O utilizador não tem Composite Role definida na 'Proposta Ativa'.{CLR_RST}")
        comp_role = input("Introduza o nome técnico da Composite Role a atualizar (ex: Z_BR_IT_SPECIALIST): ").strip().upper()
        if not comp_role:
            print("Composite Role não indicada. Operação abortada.")
            return
        dados_user["composite_role"] = comp_role

    # Obtenção da Função Individual
    single_role_escolhida = None

    if sugestao_role:
        single_role_escolhida = sugestao_role
        print(f"  {CLR_G}Função Individual recomendada:{CLR_RST} {CLR_Y}{single_role_escolhida}{CLR_RST}")
        sel_r = input(f"  Pressione [ENTER] para aceitar '{CLR_Y}{single_role_escolhida}{CLR_RST}' ou [M] para alterar: ").strip().upper()
        if sel_r == "M":
            single_role_escolhida = input("  Introduza o nome da Função Individual (Z_...): ").strip().upper()
    else:
        if roles_previamente_encontradas is not None:
            roles_proposta = roles_previamente_encontradas
        else:
            print(f"\n{CLR_C}>>> A pesquisar Função Individual para a transação '{tcode_escolhido}' na folha 'Proposta'...{CLR_RST}")
            roles_proposta = buscar_funcao_individual_para_transacao(caminho_excel, tcode_escolhido, usuario)

        if roles_proposta:
            exatos = [r for r in roles_proposta if not r.get("candidata")]
            candidatos = [r for r in roles_proposta if r.get("candidata")]

            if exatos:
                print(f"{CLR_G}Função(ões) individual(ais) localizada(s) no catálogo:{CLR_RST}")
                for idx, r_info in enumerate(exatos, 1):
                    print(f"  [{idx}] {CLR_Y}{r_info['role']}{CLR_RST} - {r_info['descricao']} ({r_info['fonte']})")
                print("  [M] Introduzir manualmente outro nome de Função Individual")

                prompt_r = f"Pressione [ENTER] para aceitar '{CLR_Y}{exatos[0]['role']}{CLR_RST}' ou escolha [1-{len(exatos)}, M] [Padrão: 1]: "
                sel_r = input(prompt_r).strip()
                if not sel_r or sel_r == "1":
                    single_role_escolhida = exatos[0]["role"]
                elif sel_r.upper() == "M":
                    single_role_escolhida = input("Introduza o nome da Função Individual (Z_...): ").strip().upper()
                else:
                    try:
                        single_role_escolhida = exatos[int(sel_r) - 1]["role"]
                    except Exception:
                        single_role_escolhida = exatos[0]["role"]

            elif candidatos:
                print(f"{CLR_Y}A transação '{tcode_escolhido}' não consta na folha 'Proposta'. Funções sugeridas:{CLR_RST}")
                for idx, r_info in enumerate(candidatos[:3], 1):
                    print(f"  [{idx}] {CLR_Y}{r_info['role']}{CLR_RST} - {r_info['descricao']} ({r_info['fonte']})")
                print("  [M] Introduzir manualmente outro nome de Função Individual")

                prompt_c = f"Pressione [ENTER] para aceitar '{CLR_Y}{candidatos[0]['role']}{CLR_RST}' ou escolha [1-{len(candidatos[:3])}, M] [Padrão: 1]: "
                sel_c = input(prompt_c).strip()
                if not sel_c or sel_c == "1":
                    single_role_escolhida = candidatos[0]["role"]
                elif sel_c.upper() == "M":
                    single_role_escolhida = input("Introduza o nome da Função Individual (Z_...): ").strip().upper()
                else:
                    try:
                        single_role_escolhida = candidatos[int(sel_c) - 1]["role"]
                    except Exception:
                        single_role_escolhida = candidatos[0]["role"]
        else:
            print(f"{CLR_Y}[AVISO] A transação '{tcode_escolhido}' não foi encontrada no catálogo da folha 'Proposta'.{CLR_RST}")
            single_role_escolhida = input("Introduza manualmente o nome da Função Individual (ex: Z_IT_EQUIPA_INTERNA): ").strip().upper()

    if not single_role_escolhida:
        print("Nenhuma Função Individual selecionada. Operação abortada.")
        return

    print(f"\n{CLR_W}>>> Função Individual selecionada: {CLR_Y}{single_role_escolhida}{CLR_RST}")

    # Confirmação do Plano de Ação
    env_primary = "QAD" if target_env in ("QAD", "QAS") else "PRD"
    env_secondary = "PRD" if env_primary == "QAD" else "QAD"

    print("\n" + "=" * 95)
    print(f"  {CLR_W}PLANO DE EXECUÇÃO E ATRIBUIÇÃO AO PERFIL DO UTILIZADOR {CLR_Y}{usuario}{CLR_RST}:")
    print("=" * 95)
    print(f"  1. Folha Departamental [{CLR_Y}{nome_sheet_dep}{CLR_RST}]:")
    print(f"     -> Marcar 'X' para {CLR_Y}{usuario}{CLR_RST} na transação {CLR_Y}{tcode_escolhido}{CLR_RST}")
    print(f"  2. Folha 'Proposta Ativa':")
    print(f"     -> Adicionar {CLR_Y}{single_role_escolhida}{CLR_RST} na linha do utilizador (próxima coluna disponível)")
    print(f"  3. Folha 'PFCG_COMPOSTA':")
    print(f"     -> Associar {CLR_Y}{comp_role}{CLR_RST} -> {CLR_Y}{single_role_escolhida}{CLR_RST}")
    print(f"  4. SAP {env_primary} via RFC:")
    print(f"     -> Atribuir Função + Efetuar Ajuste de Utilizadores (User Comparison)")
    print("=" * 95)

    confirmacao = input(f"\nConfirmar e aplicar alterações no Excel e no SAP? [{CLR_G}S{CLR_W}/N] [{CLR_G}S{CLR_W}]: ").strip().upper()
    if confirmacao not in ("", "S", "SIM", "Y", "YES", "1"):
        print("Operação cancelada pelo utilizador. Nenhuma alteração foi efetuada.")
        return

    # Executar no Excel
    print(f"\n{CLR_C}1/2. Atualizando livro Excel...{CLR_RST}")
    res_excel = aplicar_alteracoes_excel(
        caminho_excel=caminho_excel,
        dados_user=dados_user,
        tcode=tcode_escolhido,
        single_role=single_role_escolhida,
        nome_sheet_dep=nome_sheet_dep
    )
    if not res_excel.get("ok"):
        print(f"{CLR_R}[ERRO EXCEL] {res_excel.get('erro')}{CLR_RST}")
        return
    print(f"{CLR_G}[OK] Livro Excel atualizado com sucesso!{CLR_RST}")

    # Executar no SAP principal
    print(f"\n{CLR_C}2/2. Aplicando alterações no SAP {env_primary} via RFC...{CLR_RST}")
    res_prim = sincronizar_sap_rfc(env_primary, comp_role, single_role_escolhida, usuario=usuario)
    if res_prim.get("ok"):
        print(f"{CLR_G}[OK] Alterações concluídas no SAP {env_primary}!{CLR_RST}")
    else:
        print(f"{CLR_R}[AVISO SAP {env_primary}] {res_prim.get('erro')}{CLR_RST}")

    # Sincronização secundária opcional
    if env_primary == "QAD":
        resp_sec = input(f"\nDeseja também aplicar as alterações no SAP {env_secondary} (Produção)? [S/N] [N]: ").strip().upper()
        if resp_sec in ("S", "SIM", "Y", "YES", "1"):
            print(f"\n{CLR_C}>>> Aplicando alterações no SAP {env_secondary} via RFC...{CLR_RST}")
            sincronizar_sap_rfc(env_secondary, comp_role, single_role_escolhida, usuario=usuario)
    else:
        resp_sec = input(f"\nDeseja também aplicar as alterações no SAP {env_secondary} (QAS / Teste)? [S/N] [S]: ").strip().upper()
        if resp_sec in ("", "S", "SIM", "Y", "YES", "1"):
            print(f"\n{CLR_C}>>> Aplicando alterações no SAP {env_secondary} via RFC...{CLR_RST}")
            sincronizar_sap_rfc(env_secondary, comp_role, single_role_escolhida, usuario=usuario)

    print("\n" + "=" * 95)
    print(f"  {CLR_G}PROCESSO DE RESOLUÇÃO SU53 PARA {tcode_escolhido} CONCLUÍDO COM SUCESSO!{CLR_RST}")
    print("=" * 95)


# =====================================================================
# 8. FLUXO PRINCIPAL INTERATIVO
# =====================================================================

def main():
    print("\n" + "=" * 95)
    print(f"  {CLR_W}SISTEMA INTEGRADO DE PESQUISA SU53 & ATRIBUIÇÃO DE AUTORIZAÇÕES (SAP / EXCEL){CLR_RST}")
    print("=" * 95)

    caminho_excel = CAMINHO_EXCEL_PADRAO
    if not os.path.exists(caminho_excel):
        print(f"{CLR_R}[ERRO] Ficheiro mestre Excel não encontrado: {caminho_excel}{CLR_RST}")
        return

    # 1. Solicitação do utilizador e ambiente
    cli_user = None
    cli_env = None
    for arg in sys.argv[1:]:
        a_up = arg.strip().upper()
        if a_up in ("PRD", "QAS", "QAD", "DEV"):
            cli_env = "QAD" if a_up == "QAS" else a_up
        elif not cli_user:
            cli_user = a_up

    default_user = "ITSALSA" if any("ITSALSA" in str(a).upper() for a in sys.argv) else "CSILVA"
    try:
        user_input = cli_user or input(f"\n{CLR_W}Informe o utilizador SAP a analisar [{CLR_Y}{default_user}{CLR_W}]: {CLR_RST}").strip().upper()
    except (KeyboardInterrupt, EOFError):
        print("\nOperação cancelada.")
        return

    usuario = user_input or default_user

    default_env_opt = "2" if any(k in str(a).upper() for a in sys.argv for k in ("QAS", "QAD", "ITSALSA")) else "1"
    try:
        env_input = cli_env or input(f"{CLR_W}Informe o ambiente SAP ([1] PRD | [2] QAS) [{CLR_Y}{default_env_opt}{CLR_W}]: {CLR_RST}").strip().upper()
    except (KeyboardInterrupt, EOFError):
        print("\nOperação cancelada.")
        return

    if env_input in ("2", "QAS", "QAD"):
        target_env = "QAD"
    else:
        target_env = "PRD"

    # 2. Leitura do buffer SU53 via RFC
    try:
        erros = consultar_erros_su53_rfc(usuario, target_env=target_env)
    except Exception as exc:
        print(f"\n{CLR_R}[ERRO RFC] Falha ao consultar o buffer da SU53: {exc}{CLR_RST}")
        return

    # Exibir tabela recente
    exibir_tabela_erros(erros, usuario, target_env=target_env)

    # 3. Agrupamento em Transações (S_TCODE) e Objetos de Autorização
    transacoes, objetos = agrupar_erros_para_analise(erros)

    while True:
        # Apresentar painel de análise com opções combinadas
        opcoes = exibir_analise_sintese(transacoes, objetos, caminho_excel=caminho_excel, usuario=usuario)

        # SE EXISTIR EXATAMENTE 1 TRANSAÇÃO BLOQUEADA:
        if len(transacoes) == 1:
            t_item = transacoes[0]
            tc = t_item["tcode_falha"]
            opc_item = next((op for op in opcoes if op["tipo"] == "TRANSACAO" and op["valor"] == tc), {})
            sug_role = opc_item.get("sugestao_role")
            sug_txt = f" (Função: {CLR_G}{sug_role}{CLR_W})" if sug_role else ""

            print("\n" + "=" * 95)
            print(f"  {CLR_W}AÇÃO PROPOSTA PARA RESOLUÇÃO DA FALHA SU53 DO UTILIZADOR {CLR_Y}{usuario}{CLR_RST}")
            print("=" * 95)
            print(f"  Transação com Falha:   {CLR_Y}{tc}{CLR_RST} [RC={t_item['rc']}]")
            if sug_role:
                print(f"  Função Recomendada:    {CLR_G}{sug_role}{CLR_RST}")
            print(f"  Utilizador Alvo:       {CLR_Y}{usuario}{CLR_RST} [Ambiente: {target_env}]")
            print("-" * 95)

            prompt_q = (
                f"{CLR_W}Deseja adicionar a transação {CLR_Y}{tc}{CLR_W}{sug_txt} "
                f"ao perfil do utilizador {CLR_Y}{usuario}{CLR_W}? "
                f"[{CLR_G}S{CLR_W}/N/Outra/0] [{CLR_G}S{CLR_W}]: {CLR_RST}"
            )
            escolha = input(prompt_q).strip().upper()

            if escolha in ("", "S", "SIM", "Y", "YES", "1"):
                executar_atribuicao_transacao(
                    caminho_excel=caminho_excel,
                    usuario=usuario,
                    tcode_escolhido=tc,
                    sugestao_role=sug_role,
                    roles_previamente_encontradas=opc_item.get("roles_encontradas"),
                    target_env=target_env
                )
                return

            elif escolha in ("N", "NAO", "NO"):
                if objetos:
                    resp_obj = input(f"\nDeseja diagnosticar um dos outros {len(objetos)} objetos de autorização com falha? [S/N] [S]: ").strip().upper()
                    if resp_obj in ("", "S", "SIM", "Y", "YES", "1"):
                        # Tratar objeto
                        diagnosticar_objeto_autorizacao(
                            usuario=usuario,
                            obj_info=objetos[0],
                            caminho_excel=caminho_excel
                        )
                        return
                print("Operação concluída sem efetuar alterações.")
                return

            elif escolha == "0":
                print("Saindo sem efetuar alterações.")
                return

            elif escolha in ("O", "OUTRA", "M", "MANUAL"):
                tipo_manual = input("Deseja introduzir uma [T]ransação ou um [O]bjeto? [T/O]: ").strip().upper()
                if tipo_manual.startswith("O"):
                    obj_manual = input("Introduza o nome do Objeto de Autorização (ex: S_USER_UID): ").strip().upper()
                    if obj_manual:
                        diagnosticar_objeto_autorizacao(
                            usuario=usuario,
                            obj_info={
                                "objeto": obj_manual,
                                "descricao": obter_descricao_objeto_sap(obj_manual),
                                "rc_mais_recente": 12,
                                "data_hora_recente": "Consulta Manual",
                                "programa": "MANUAL",
                                "linha": 0,
                                "campos_relevantes": []
                            },
                            caminho_excel=caminho_excel
                        )
                else:
                    tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                    if tc_manual:
                        executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)
                break

            else:
                # Tenta índice se selecionou da lista
                try:
                    idx_sel = int(escolha) - 1
                    if 0 <= idx_sel < len(opcoes):
                        item_sel = opcoes[idx_sel]
                        if item_sel["tipo"] == "TRANSACAO":
                            executar_atribuicao_transacao(
                                caminho_excel=caminho_excel,
                                usuario=usuario,
                                tcode_escolhido=item_sel["valor"],
                                sugestao_role=item_sel.get("sugestao_role"),
                                roles_previamente_encontradas=item_sel.get("roles_encontradas"),
                                target_env=target_env
                            )
                            return
                        elif item_sel["tipo"] == "OBJETO":
                            diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=item_sel["detalhes"], caminho_excel=caminho_excel)
                            return
                except ValueError:
                    print("[AVISO] Opção não reconhecida.")

        # SE EXISTIREM MÚLTIPLAS TRANSAÇÕES BLOQUEADAS:
        elif len(transacoes) > 1:
            print("\n" + "=" * 95)
            print(f"  {CLR_W}AÇÃO PROPOSTA PARA RESOLUÇÃO DAS FALHAS SU53 - UTILIZADOR {CLR_Y}{usuario}{CLR_RST}")
            print("=" * 95)
            print(f"Foram detetadas {len(transacoes)} transações bloqueadas na SU53:")
            for i, t in enumerate(transacoes, 1):
                tc = t["tcode_falha"]
                opc_item = next((op for op in opcoes if op["tipo"] == "TRANSACAO" and op["valor"] == tc), {})
                sug = opc_item.get("sugestao_role") or "Manual"
                print(f"  [{CLR_Y}{i}{CLR_RST}] Adicionar {CLR_Y}{tc:<20}{CLR_RST} (Função: {CLR_G}{sug}{CLR_RST}) ao perfil de {CLR_Y}{usuario}{CLR_RST}")
            print(f"  [{CLR_Y}T{CLR_RST}] Adicionar TODAS as {len(transacoes)} transações acima ao perfil de {CLR_Y}{usuario}{CLR_RST}")
            if objetos:
                print(f"  [{CLR_Y}B{CLR_RST}] Diagnosticar outros Objetos de Autorização detetados")
            print(f"  [{CLR_Y}O{CLR_RST}] Introduzir manualmente outra Transação ou Objeto")
            print(f"  [{CLR_Y}0{CLR_RST}] Sair sem efetuar alterações")
            print("-" * 95)

            escolha = input(f"Selecione a opção pretendida [1-{len(transacoes)}, T, B, O, 0] [Padrão: 1]: ").strip().upper()

            if escolha in ("", "1"):
                t_item = transacoes[0]
                tc = t_item["tcode_falha"]
                opc_item = next((op for op in opcoes if op["tipo"] == "TRANSACAO" and op["valor"] == tc), {})
                executar_atribuicao_transacao(
                    caminho_excel=caminho_excel,
                    usuario=usuario,
                    tcode_escolhido=tc,
                    sugestao_role=opc_item.get("sugestao_role"),
                    roles_previamente_encontradas=opc_item.get("roles_encontradas"),
                    target_env=target_env
                )
                return

            elif escolha == "T":
                for t in transacoes:
                    tc = t["tcode_falha"]
                    opc_item = next((op for op in opcoes if op["tipo"] == "TRANSACAO" and op["valor"] == tc), {})
                    executar_atribuicao_transacao(
                        caminho_excel=caminho_excel,
                        usuario=usuario,
                        tcode_escolhido=tc,
                        sugestao_role=opc_item.get("sugestao_role"),
                        roles_previamente_encontradas=opc_item.get("roles_encontradas"),
                        target_env=target_env
                    )
                return

            elif escolha == "0":
                print("Saindo sem efetuar alterações.")
                return

            elif escolha in ("B", "OBJETOS"):
                if objetos:
                    diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=objetos[0], caminho_excel=caminho_excel)
                    continue

            elif escolha in ("O", "OUTRA", "M", "MANUAL"):
                tipo_manual = input("Deseja introduzir uma [T]ransação ou um [O]bjeto? [T/O]: ").strip().upper()
                if tipo_manual.startswith("O"):
                    obj_manual = input("Introduza o nome do Objeto de Autorização (ex: S_USER_UID): ").strip().upper()
                    if obj_manual:
                        diagnosticar_objeto_autorizacao(
                            usuario=usuario,
                            obj_info={
                                "objeto": obj_manual,
                                "descricao": obter_descricao_objeto_sap(obj_manual),
                                "rc_mais_recente": 12,
                                "data_hora_recente": "Consulta Manual",
                                "programa": "MANUAL",
                                "linha": 0,
                                "campos_relevantes": []
                            },
                            caminho_excel=caminho_excel
                        )
                else:
                    tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                    if tc_manual:
                        executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)
                break

            else:
                try:
                    idx_sel = int(escolha) - 1
                    if 0 <= idx_sel < len(transacoes):
                        t_item = transacoes[idx_sel]
                        tc = t_item["tcode_falha"]
                        opc_item = next((op for op in opcoes if op["tipo"] == "TRANSACAO" and op["valor"] == tc), {})
                        executar_atribuicao_transacao(
                            caminho_excel=caminho_excel,
                            usuario=usuario,
                            tcode_escolhido=tc,
                            sugestao_role=opc_item.get("sugestao_role"),
                            roles_previamente_encontradas=opc_item.get("roles_encontradas"),
                            target_env=target_env
                        )
                        return
                    else:
                        print("[AVISO] Opção fora do intervalo válido.")
                except ValueError:
                    print("[AVISO] Opção inválida.")

        # SE NÃO HOUVER TRANSAÇÕES BLOQUEADAS (APENAS OBJETOS):
        else:
            if objetos:
                print("\n" + "=" * 95)
                print(f"  {CLR_W}DIAGNÓSTICO DE OBJETOS DE AUTORIZAÇÃO - UTILIZADOR {CLR_Y}{usuario}{CLR_RST}")
                print("=" * 95)
                for i, o in enumerate(objetos, 1):
                    desc = f" ({o['descricao']})" if o["descricao"] else ""
                    print(f"  [{CLR_Y}{i}{CLR_RST}] Objeto: {CLR_M}{o['objeto']:<20}{CLR_RST}{desc}")
                print("  [O] Introduzir manualmente outra Transação ou Objeto")
                print("  [0] Sair sem efetuar alterações")
                print("-" * 95)

                escolha = input(f"Selecione o objeto a diagnosticar [1-{len(objetos)}, O, 0] [Padrão: 1]: ").strip().upper()
                if escolha in ("", "1"):
                    diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=objetos[0], caminho_excel=caminho_excel)
                    return
                elif escolha == "0":
                    print("Saindo sem efetuar alterações.")
                    return
                elif escolha in ("O", "OUTRA", "M", "MANUAL"):
                    tipo_manual = input("Deseja introduzir uma [T]ransação ou um [O]bjeto? [T/O]: ").strip().upper()
                    if tipo_manual.startswith("O"):
                        obj_manual = input("Introduza o nome do Objeto de Autorização (ex: S_USER_UID): ").strip().upper()
                        if obj_manual:
                            diagnosticar_objeto_autorizacao(
                                usuario=usuario,
                                obj_info={
                                    "objeto": obj_manual,
                                    "descricao": obter_descricao_objeto_sap(obj_manual),
                                    "rc_mais_recente": 12,
                                    "data_hora_recente": "Consulta Manual",
                                    "programa": "MANUAL",
                                    "linha": 0,
                                    "campos_relevantes": []
                                },
                                caminho_excel=caminho_excel
                            )
                    else:
                        tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                        if tc_manual:
                            executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)
                    break
                else:
                    try:
                        idx_sel = int(escolha) - 1
                        if 0 <= idx_sel < len(objetos):
                            diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=objetos[idx_sel], caminho_excel=caminho_excel)
                            return
                    except ValueError:
                        print("[AVISO] Opção inválida.")
            else:
                print(f"\n{CLR_G}[INFO] Não foram detetados erros na SU53 do utilizador '{usuario}'.{CLR_RST}")
                return


if __name__ == "__main__":
    main()


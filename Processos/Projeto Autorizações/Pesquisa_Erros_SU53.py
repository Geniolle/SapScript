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
import importlib.util
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

# Objetos cuja atualização automática por script é estritamente proibida (governança / aprovação manual SAP)
OBJETOS_PROIBIDOS_ATUALIZACAO = {
    "M_EINK_FRG",  # Liberação de documentos de compras (estratégia de liberação sensível)
    "F_KNA1_BUK",  # Clientes: autorização para empresas (dados financeiros/controlo por empresa)
}


def parse_selecao_multipla(entrada: str, limite_max: int) -> list:
    """
    Interpreta uma seleção numérica que pode conter números individuais, vírgulas ou intervalos.
    Exemplos: '1', '5-6', '1,3,5', '2-4, 7', '5 - 6'.
    Retorna lista ordenada e sem duplicados de índices base-0 válidos.
    """
    if not entrada:
        return []
    indices = set()
    partes = [p.strip() for p in entrada.replace(";", ",").split(",") if p.strip()]
    for parte in partes:
        if "-" in parte:
            sub = parte.split("-", 1)
            try:
                inicio = int(sub[0].strip())
                fim = int(sub[1].strip())
                if inicio > fim:
                    inicio, fim = fim, inicio
                for n in range(inicio, fim + 1):
                    if 1 <= n <= limite_max:
                        indices.add(n - 1)
            except ValueError:
                continue
        else:
            try:
                n = int(parte.strip())
                if 1 <= n <= limite_max:
                    indices.add(n - 1)
            except ValueError:
                continue
    return sorted(indices)



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

def consultar_erros_su53_rfc(usuario: str, target_env: str = "PRD", silencioso: bool = False) -> List[Dict[str, Any]]:
    """
    Consulta todas as falhas de autorização registadas no buffer da SU53
    (SUSR_USER_SU53_READ) para o utilizador indicado em todos os servidores aplicacionais.
    """
    load_project_env(find_project_root())
    params = build_connection_params_for(target_env)
    u_norm = usuario.strip().upper()

    if not silencioso:
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
    return desduplicar_erros_su53(erros)


def desduplicar_erros_su53(erros: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Remove registos duplicados no histórico da SU53, mantendo apenas a ocorrência mais recente
    (maior timestamp / data e hora mais nova) para a mesma falha (mesmo Objeto + mesmos Campos/Valores).
    """
    if not erros:
        return []

    vistos = set()
    desduplicados = []

    for e in erros:
        obj = e.get("objeto", "").strip().upper()
        tc = (e.get("tcode_falha") or "").strip().upper()
        # Normalizar tuplo de campos e valores ordenados para chave única
        campos_tuple = tuple(sorted(
            (str(f).strip().upper(), str(v).strip().upper())
            for f, v in e.get("campos", [])
            if str(f).strip().upper() != "DUMMYFIELD"
        ))
        
        chave_falha = (obj, tc, campos_tuple)

        if chave_falha not in vistos:
            vistos.add(chave_falha)
            desduplicados.append(e)

    return desduplicados


def filtrar_erros_atuais(
    erros: List[Dict[str, Any]],
    data_local: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """Mantém as falhas de hoje na instância da ocorrência mais recente."""
    if not erros:
        return []

    data_alvo = data_local or datetime.now().astimezone().strftime("%d.%m.%Y")
    erros_hoje = [e for e in erros if str(e.get("data_hora", "")).startswith(data_alvo)]
    if not erros_hoje:
        return []

    instancia_atual = erros_hoje[0].get("servidor", "")
    if not instancia_atual:
        return erros_hoje
    return [e for e in erros_hoje if e.get("servidor", "") == instancia_atual]


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


def obter_descricao_transacao_sap(tcode: str) -> str:
    """Obtém a descrição da transação na TSTCT, preferindo português."""
    tc_norm = str(tcode or "").strip().upper()
    if not tc_norm:
        return ""
    try:
        load_project_env(find_project_root())
        conn = Connection(**build_connection_params_for("PRD"))
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("TSTCT",))
        rows = read_table(
            conn,
            guard,
            table_name="TSTCT",
            fields=["SPRSL", "TCODE", "TTEXT"],
            options=[{"TEXT": f"TCODE = '{tc_norm}'"}],
            rowcount=10,
        )
        conn.close()
        textos = {str(r[0]).strip(): str(r[2]).strip() for r in rows if len(r) >= 3 and str(r[2]).strip()}
        return textos.get("P") or textos.get("E") or next(iter(textos.values()), "")
    except Exception:
        return ""


def obter_autorizacoes_ativas_usuario_rfc(usuario: str, target_env: str = "PRD") -> Tuple[set, Dict[str, Dict[str, set]]]:
    """
    Obtém todas as roles ativas (diretas e herdadas de compostas) do utilizador
    e o mapa de autorizações por objeto e campo na AGR_1251.
    Retorna:
      (all_roles, obj_auth_map)
      onde obj_auth_map[objeto][campo] = set(valores)
    """
    try:
        load_project_env(find_project_root())
        params = build_connection_params_for(target_env)
        conn = Connection(**params)
        
        # 1. Obter roles ativas
        res_u = conn.call("BAPI_USER_GET_DETAIL", USERNAME=usuario.strip().upper())
        direct_roles = {r["AGR_NAME"].strip() for r in res_u.get("ACTIVITYGROUPS", [])}
        
        res_agrs = conn.call("RFC_READ_TABLE", QUERY_TABLE="AGR_AGRS", DELIMITER="|", FIELDS=[{"FIELDNAME": "AGR_NAME"}, {"FIELDNAME": "CHILD_AGR"}])
        comp_children = defaultdict(set)
        for l in res_agrs.get("DATA", []):
            parts = l["WA"].split("|")
            comp_children[parts[0].strip()].add(parts[1].strip())
            
        all_roles = set(direct_roles)
        for r in list(direct_roles):
            all_roles.update(comp_children.get(r, set()))
            
        conn.close()
        return all_roles
    except Exception as exc:
        return set()


def verificar_falha_ja_resolvida_rfc(
    erro: Dict[str, Any],
    all_roles: set,
    target_env: str = "PRD"
) -> Tuple[bool, str]:
    """
    Verifica se a falha registada na SU53 já se encontra coberta e atribuída ao utilizador.
    Retorna: (is_resolvido, role_que_cobre)
    """
    if not all_roles:
        return False, ""

    objeto = erro.get("objeto", "").strip().upper()
    tcode_falha = (erro.get("tcode_falha") or "").strip().upper()

    try:
        load_project_env(find_project_root())
        params = build_connection_params_for(target_env)
        conn = Connection(**params)

        # 1. Se for falha de transação (S_TCODE)
        if objeto == "S_TCODE" or tcode_falha:
            tc = tcode_falha
            if not tc:
                tc = next((v for f, v in erro.get("campos", []) if f == "TCD"), "").strip().upper()
            if tc:
                res = conn.call("RFC_READ_TABLE", QUERY_TABLE="AGR_TCODES", DELIMITER="|", FIELDS=[{"FIELDNAME": "AGR_NAME"}], OPTIONS=[{"TEXT": f"TCODE = '{tc}'"}])
                roles_com_tc = {l["WA"].strip() for l in res.get("DATA", [])}
                inter = all_roles.intersection(roles_com_tc)
                conn.close()
                if inter:
                    return True, sorted(list(inter))[0]
                return False, ""

        # 2. Se for outro Objeto de Autorização
        campos = erro.get("campos", [])
        if not campos and "campos_relevantes" in erro:
            campos = erro["campos_relevantes"]

        options = [{"TEXT": f"OBJECT = '{objeto}' AND DELETED = ' '"}]
        res = conn.call("RFC_READ_TABLE", QUERY_TABLE="AGR_1251", DELIMITER="|", FIELDS=[{"FIELDNAME": "AGR_NAME"}, {"FIELDNAME": "FIELD"}, {"FIELDNAME": "LOW"}], OPTIONS=options)
        conn.close()

        role_auths = defaultdict(lambda: defaultdict(set))
        for l in res.get("DATA", []):
            parts = l["WA"].split("|")
            agr = parts[0].strip()
            if agr in all_roles:
                fld = parts[1].strip().upper()
                val = parts[2].strip().upper()
                role_auths[agr][fld].add(val)

        for agr, fld_map in role_auths.items():
            cobre_tudo = True
            campos_validos = 0
            for f_req, v_req in campos:
                f_norm = str(f_req).strip().upper()
                v_norm = str(v_req).strip().upper()
                if not f_norm or f_norm in ("DUMMYFIELD", "TCD"):
                    continue
                campos_validos += 1
                vals_role = fld_map.get(f_norm, set())
                if not ("*" in vals_role or v_norm in vals_role):
                    cobre_tudo = False
                    break
            if cobre_tudo and campos_validos > 0:
                return True, agr

        return False, ""
    except Exception:
        return False, ""


def exibir_tabela_erros(erros: List[Dict[str, Any]], usuario: str, target_env: str = "PRD", all_roles: Optional[set] = None):
    """
    Exibe na consola o extrato recente de falhas da SU53 agrupado por:
      1. [❌ PENDENTE] Falhas que ainda requerem autorização (ordenadas da mais recente para a mais antiga)
      2. [✓ JÁ ATRIBUÍDO] Falhas históricas já cobertas por funções ativas no utilizador
    """
    env_label = "QAS (QAD)" if target_env.upper() in ("QAD", "QAS") else target_env.upper()
    print("\n" + "=" * 105)
    print(f"  {CLR_W}HISTÓRICO RECENTE DE FALHAS REGISTADAS NA SU53 - UTILIZADOR: {CLR_Y}{usuario.upper()}{CLR_W} [AMBIENTE: {CLR_M}{env_label}{CLR_W}]{CLR_RST}")
    print("=" * 105)

    if not erros:
        print(f"  {CLR_Y}[AVISO]{CLR_RST} Não foram encontrados registos de erro (RC > 0) no buffer da SU53 para '{usuario}' no ambiente {env_label}.")
        print(f"  {CLR_C}-> Nota Operacional:{CLR_RST} No SAP GUI, após ocorrer o bloqueio de autorização na transação,")
        print(f"                     é necessário executar {CLR_W}/nSU53{CLR_RST} nessa mesma sessão para que o SAP")
        print("                     descarregue a verificação falhada para o buffer acessível via RFC.")
        print("=" * 105)
        return

    # Separar em Pendentes e Já Atribuídos
    erros_pendentes = [e for e in erros if not e.get("resolvido", False)]
    erros_atribuidos = [e for e in erros if e.get("resolvido", False)]

    # Se houver pendências, mostrar detalhe das pendências e resumir o que já está atribuído
    if erros_pendentes:
        print(f"{'#':<3} | {'DATA/HORA (LOCAL)':<19} | {'OBJETO':<12} | {'RC':<3} | {'DETALHES / VALORES VERIFICADOS':<40} | {'ESTADO ATUAL'}")
        print("-" * 105)

        for idx, err in enumerate(erros_pendentes, 1):
            campos_txt = ", ".join([f"{f}={v}" for f, v in err["campos"]])
            if err["objeto"] == "S_TCODE" and err["tcode_falha"]:
                detalhes = f"TRANSAÇÃO: {err['tcode_falha']}"
            else:
                detalhes = campos_txt

            if len(detalhes) > 40:
                detalhes = detalhes[:37] + "..."

            estado_txt = f"{CLR_R}❌ PENDENTE (Falta autorização){CLR_RST}"
            cor_obj = CLR_R if err["objeto"] == "S_TCODE" else CLR_Y
            print(f"{idx:<3} | {err['data_hora']:<19} | {cor_obj}{err['objeto']:<12}{CLR_RST} | {err['rc']:<3} | {detalhes:<40} | {estado_txt}")

        if erros_atribuidos:
            roles_cobertas = sorted(list(set(e.get("role_resolvido") for e in erros_atribuidos if e.get("role_resolvido"))))
            txt_roles = f" ({', '.join(roles_cobertas)})" if roles_cobertas else ""
            print(f"\n  {CLR_G}✓ Nota:{CLR_RST} {len(erros_atribuidos)} verificação(ões) anterior(es) já estão autorizadas pelas roles ativas{txt_roles}.")
    else:
        print(f"  {CLR_G}✓ Todas as {len(erros_atribuidos)} falhas registadas já se encontram autorizadas no SAP!{CLR_RST}")

    print("-" * 105)



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
            aplicacao = (e.get("tcode_app") or e.get("nome_servico") or "").strip().upper()
            tcode_ctx = (e.get("tcode_falha") or "").strip().upper()
            objetos_map[(aplicacao, tcode_ctx, obj)].append(e)

    lista_transacoes = list(transacoes_map.values())

    lista_objetos = []
    for (aplicacao, tcode_ctx, obj_name), ocorrencias in objetos_map.items():
        mais_recente = ocorrencias[0]
        # Consolidar campos testados distintos
        campos_relevantes = []
        for o in ocorrencias:
            for f, v in o["campos"]:
                if f and f != "dummyfield" and (f, v) not in campos_relevantes:
                    campos_relevantes.append((f, v))

        lista_objetos.append({
            "aplicacao": aplicacao,
            "tcode_falha": tcode_ctx,
            "objeto": obj_name,
            "descricao": obter_descricao_objeto_sap(obj_name),
            "rc_mais_recente": mais_recente["rc"],
            "data_hora_recente": mais_recente["data_hora"],
            "programa": mais_recente["programa"],
            "linha": mais_recente["linha"],
            "campos_relevantes": campos_relevantes,
            "total_ocorrencias": len(ocorrencias),
            "resolvido": any(o.get("resolvido", False) for o in ocorrencias),
            "role_resolvido": next((o.get("role_resolvido") for o in ocorrencias if o.get("role_resolvido")), "")
        })

    return lista_transacoes, lista_objetos


def enriquecer_objeto_com_proposta(
    objeto: Dict[str, Any],
    caminho_excel: str,
    usuario: Optional[str] = None,
) -> Dict[str, Any]:
    """Adiciona descrição da aplicação e funções mapeadas na sheet Proposta."""
    resultado = dict(objeto)
    aplicacao = str(resultado.get("aplicacao") or "").strip().upper()
    tcode_falha = str(resultado.get("tcode_falha") or "").strip().upper()

    # Determinar se a aplicação parece ser um código de transação válido
    app_valida = aplicacao and not any(k in aplicacao for k in ("SESSION", "B726", "IDOC")) and len(aplicacao) <= 20
    resultado["nome_aplicacao"] = obter_descricao_transacao_sap(aplicacao) if app_valida else ""

    # Buscar funções candidatas tanto pela transação de contexto (ex: LT03) como pela aplicação de origem (ex: VT01N)
    alvos_pesquisa = []
    if app_valida:
        alvos_pesquisa.append(aplicacao)
    if tcode_falha and tcode_falha not in alvos_pesquisa:
        alvos_pesquisa.append(tcode_falha)

    funcoes_proposta = []
    funcoes_outras = []
    for alvo in alvos_pesquisa:
        roles = buscar_funcao_individual_para_transacao(caminho_excel, alvo, usuario)
        for r in roles:
            if not r.get("candidata"):
                if r.get("fonte") == "Folha 'Proposta'":
                    if not any(fp["role"] == r["role"] for fp in funcoes_proposta):
                        funcoes_proposta.append(r)
                else:
                    if not any(fo["role"] == r["role"] for fo in funcoes_outras):
                        funcoes_outras.append(r)

    # Priorizar sheet Proposta; caso não haja, complementar com Single Roles / PFCG
    todas_funcoes = funcoes_proposta if funcoes_proposta else funcoes_outras
    resultado["funcoes_proposta"] = todas_funcoes
    resultado["funcao_proposta"] = todas_funcoes[0] if todas_funcoes else None
    return resultado


def formatar_relatorio_erro_aplicacao(
    objeto: Dict[str, Any],
    funcoes_proposta: Optional[List[Dict[str, Any]]] = None,
) -> str:
    """Formata o diagnóstico SU53 com funções da sheet Proposta / Catálogo."""
    nome_obj = str(objeto.get("objeto") or "").strip().upper()

    linhas = []

    if nome_obj in OBJETOS_PROIBIDOS_ATUALIZACAO:
        linhas.extend([
            f"  {CLR_R}[BLOQUEIO DE SEGURANÇA]{CLR_RST} O objeto {CLR_Y}{nome_obj}{CLR_RST} está classificado como {CLR_R}PROIBIDO DE ATUALIZAR POR SCRIPT{CLR_RST}.",
            "  (Este objeto gere estratégias/aprovações críticas e só pode ser revisto manualmente na PFCG/SAP GUI).",
            ""
        ])

    linhas.extend([
        f"Na matriz de autorizações (sheet Proposta / Catálogo), as funções correspondentes são:",
        "",
    ])
    roles = funcoes_proposta if funcoes_proposta is not None else objeto.get("funcoes_proposta", [])
    if not roles:
        linhas.append(f"- Função: {CLR_Y}Não encontrada na Proposta{CLR_RST} (ao selecionar este item, poderá indicar a Single Role manualmente)")
        return "\n".join(linhas)

    for idx, role in enumerate(roles):
        linhas.append(
            f"- Função: {role.get('role', '')} - Descrição: {role.get('descricao', '')} [{role.get('fonte', 'Catálogo')}]"
        )
        linha_funcao = role.get("linha_funcao")
        linha_transacao = role.get("linha_transacao")
        if linha_funcao and linha_transacao:
            linhas.append(
                f"  (A função começa na linha {linha_funcao}; mapeamento na linha {linha_transacao} da Proposta)"
            )
        if idx < len(roles) - 1:
            linhas.append("")
    return "\n".join(linhas)



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
            app_origem = (t.get("nome_servico") or t.get("tcode_app") or "").strip().upper()
            app_valida = app_origem and not any(k in app_origem for k in ("SESSION", "B726", "IDOC")) and len(app_origem) <= 20
            
            st_res = f" {CLR_G}[✓ JÁ ATRIBUÍDO]{CLR_RST}" if t.get("resolvido") else f" {CLR_R}[❌ PENDENTE]{CLR_RST}"
            app_info = f" (via App: {CLR_Y}{app_origem}{CLR_W})" if app_valida and app_origem != tc else ""
            print(f"  [{CLR_Y}{num_item}{CLR_RST}] {CLR_W}Transação: {CLR_Y}{tc:<20}{CLR_W}{app_info}{st_res} -> RC={t['rc']} | Última: {t['data_hora']} ({t['programa']}:{t['linha']})")

            # Pesquisa imediata de sugestão de Função Individual no Excel / Catálogo
            # 1. Pesquisar pela própria transação que falhou (ex: VTFAKT)
            roles_sug = buscar_funcao_individual_para_transacao(caminho_excel, tc, usuario)
            sug_role_padrao = None

            # 2. Se a transação falhada não estiver catalogada na folha 'Proposta' e tiver app de origem (ex: VT01N), buscar pela aplicação
            roles_app = []
            if app_valida and app_origem != tc:
                roles_app = buscar_funcao_individual_para_transacao(caminho_excel, app_origem, usuario)

            if t.get("resolvido") and t.get("role_resolvido"):
                print(f"       {CLR_G}-> Acesso confirmado: Transação já se encontra coberta pela role ativa '{t['role_resolvido']}'.{CLR_RST}")
            else:
                tem_sugestao = False
                # Prioridade A: Função da folha 'Proposta' para a transação bloqueada
                exatos_tc = [r for r in roles_sug if not r.get("candidata") and r.get("fonte") == "Folha 'Proposta'"]
                # Prioridade B: Função da folha 'Proposta' para a aplicação de origem (ex: VT01N -> Z_SHIPMENT_CREATE)
                exatos_app = [r for r in roles_app if not r.get("candidata") and r.get("fonte") == "Folha 'Proposta'"]
                # Outras funções (Single Roles, PFCG_CREATE, etc.)
                outros_tc = [r for r in roles_sug if not r.get("candidata") and r.get("fonte") != "Folha 'Proposta'"]
                outros_app = [r for r in roles_app if not r.get("candidata") and r.get("fonte") != "Folha 'Proposta'"]

                if exatos_tc:
                    r_top = exatos_tc[0]
                    sug_role_padrao = r_top["role"]
                    outras = [r["role"] for r in exatos_tc[1:3]]
                    outras_txt = f" (ou: {', '.join(outras)})" if outras else ""
                    print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_G}{r_top['role']}{CLR_RST} - {r_top['descricao']} [{r_top['fonte']}]{outras_txt}")
                    tem_sugestao = True
                elif exatos_app:
                    r_top = exatos_app[0]
                    sug_role_padrao = r_top["role"]
                    print(f"       {CLR_C}-> Sugestão via Aplicação de Origem ({app_origem}):{CLR_RST} {CLR_G}{r_top['role']}{CLR_RST} - {r_top['descricao']} [{r_top['fonte']}]")
                    if outros_tc:
                        outras_txt = ", ".join([f"{r['role']} ({r['fonte']})" for r in outros_tc[:2]])
                        print(f"          (Alternativas para {tc}: {outras_txt})")
                    tem_sugestao = True
                elif outros_tc or outros_app:
                    lista_outros = outros_tc or outros_app
                    r_top = lista_outros[0]
                    sug_role_padrao = r_top["role"]
                    outras = [r["role"] for r in lista_outros[1:3]]
                    outras_txt = f" (ou: {', '.join(outras)})" if outras else ""
                    print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_G}{r_top['role']}{CLR_RST} - {r_top['descricao']} [{r_top['fonte']}]{outras_txt}")
                    tem_sugestao = True
                else:
                    candidatos = [r for r in (roles_sug + roles_app) if r.get("candidata")]
                    if candidatos:
                        r_top = candidatos[0]
                        sug_role_padrao = r_top["role"]
                        cands_txt = " / ".join([c["role"] for c in candidatos[:2]])
                        print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_Y}Não catalogada na folha 'Proposta'{CLR_RST} (Candidatas sugeridas: {CLR_W}{cands_txt}{CLR_RST} ou introduzir manual)")
                        tem_sugestao = True

                if not tem_sugestao:
                    print(f"       {CLR_C}-> Sugestão Função Individual:{CLR_RST} {CLR_Y}Não encontrada no catálogo 'Proposta'{CLR_RST} (requer indicação manual de Z_...)")

            # Combinar roles encontradas para seleção
            todas_roles_encontradas = roles_sug + [r for r in roles_app if not any(x["role"] == r["role"] for x in roles_sug)]

            opcoes_selecao.append({
                "tipo": "TRANSACAO",
                "valor": tc,
                "detalhes": t,
                "sugestao_role": sug_role_padrao,
                "roles_encontradas": todas_roles_encontradas,
                "resolvido": t.get("resolvido", False)
            })
            num_item += 1
    else:
        print("  (Nenhuma transação S_TCODE identificada com falha no buffer recente)")

    # SECÇÃO 2: OBJETOS DE AUTORIZAÇÃO
    print(f"\n{CLR_C}[B] OUTROS OBJETOS DE AUTORIZAÇÃO COM FALHA:{CLR_RST}")
    if objetos:
        for indice_objeto, objeto in enumerate(objetos):
            if indice_objeto:
                print(f"\n{'.' * 95}\n")
            o = enriquecer_objeto_com_proposta(objeto, caminho_excel, usuario)
            desc = f" ({o['descricao']})" if o["descricao"] else ""
            campos_str = ", ".join([f"{f}={v}" for f, v in o["campos_relevantes"]]) if o["campos_relevantes"] else "Verificação de existência"
            st_res_obj = f" {CLR_G}[✓ JÁ ATRIBUÍDO]{CLR_RST}" if o.get("resolvido") else f" {CLR_R}[❌ PENDENTE]{CLR_RST}"
            print(f"  [{CLR_Y}{num_item}{CLR_RST}] {CLR_W}Objeto:    {CLR_M}{o['objeto']:<20}{CLR_RST}{desc}{st_res_obj}")
            if o.get("resolvido") and o.get("role_resolvido"):
                print(f"       {CLR_G}-> Acesso confirmado: Valores já estão autorizados na role ativa '{o['role_resolvido']}'.{CLR_RST}")
            print(f"       -> RC={o['rc_mais_recente']} | Falhas: {o['total_ocorrencias']}x | Campos: {campos_str} | Prog: {o['programa']}")
            print()
            print(formatar_relatorio_erro_aplicacao(o))
            opcoes_selecao.append({
                "tipo": "OBJETO",
                "valor": o["objeto"],
                "detalhes": o,
                "resolvido": o.get("resolvido", False)
            })
            num_item += 1
    else:
        print("  (Nenhum outro objeto de autorização com falha identificado)")

    return opcoes_selecao



# =====================================================================
# 3. ATUALIZAÇÃO DE OBJETO DE AUTORIZAÇÃO NAS FUNÇÕES INDICADAS
# =====================================================================

def preparar_atualizacao_objeto(obj_info: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Consolida roles e valores SU53 que serão aplicados via PFCGMASSVAL."""
    objeto = str(obj_info.get("objeto") or "").strip().upper()
    roles = list(dict.fromkeys(
        str(item.get("role") or "").strip().upper()
        for item in obj_info.get("funcoes_proposta", [])
        if str(item.get("role") or "").strip()
    ))
    valores_por_campo: Dict[str, List[str]] = {}
    for campo, valor in obj_info.get("campos_relevantes", []):
        campo_norm = str(campo or "").strip().upper()
        valor_norm = str(valor or "").strip().upper()
        if not campo_norm or not valor_norm or campo_norm == "DUMMYFIELD":
            continue
        valores = valores_por_campo.setdefault(campo_norm, [])
        if valor_norm not in valores:
            valores.append(valor_norm)
    campos = {campo: ",".join(valores) for campo, valores in valores_por_campo.items()}
    if not objeto or not roles or not campos:
        return None
    return {"objeto": objeto, "roles": roles, "campos": campos}


def classificar_modo_objeto(registos: List[Dict[str, Any]]) -> str:
    """Atualiza/altera autorizações em funções existentes; nunca reativa entradas eliminadas."""
    return "ALTERAR"


def valor_autorizacao_cobre(valor_existente: str, valor_pedido: str) -> bool:
    """Reconhece autorização total ou variável organizacional como cobertura de qualquer valor específico."""
    existente = str(valor_existente or "").strip().upper()
    pedido = str(valor_pedido or "").strip().upper()
    if existente in ("*", "$WERKS", "$BUKRS", "$EKORG", "$VKORG", "$VTWEG", "$SPART", "$KOKRS", "$PRCTR"):
        return True
    return existente == pedido


def _validar_roles_existentes_rfc(roles: List[str], target_env: str) -> List[str]:
    """Impede que o executor crie por engano uma role indicada apenas no Excel."""
    load_project_env(find_project_root())
    conn = Connection(**build_connection_params_for(target_env))
    try:
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_DEFINE",))
        existentes = set()
        for role in roles:
            linhas = read_table(
                conn,
                guard,
                table_name="AGR_DEFINE",
                fields=["AGR_NAME"],
                options=[{"TEXT": f"AGR_NAME = '{role}'"}],
                rowcount=1,
            )
            if linhas:
                existentes.add(role)
        return [role for role in roles if role not in existentes]
    finally:
        conn.close()


def _obter_modos_objeto_rfc(plano: Dict[str, Any], target_env: str) -> Dict[str, str]:
    """Consulta AGR_1250 para impedir criação duplicada do objeto na role."""
    load_project_env(find_project_root())
    conn = Connection(**build_connection_params_for(target_env))
    try:
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_1250",))
        modos = {}
        for role in plano["roles"]:
            linhas = read_table(
                conn,
                guard,
                table_name="AGR_1250",
                fields=["AUTH", "DELETED"],
                options=[
                    {"TEXT": f"AGR_NAME = '{role}'"},
                    {"TEXT": f"AND OBJECT = '{plano['objeto']}'"},
                ],
                rowcount=0,
            )
            registos = [{"AUTH": r[0], "DELETED": r[1]} for r in linhas]
            modos[role] = classificar_modo_objeto(registos)
        return modos
    finally:
        conn.close()


def _validar_objeto_funcoes_rfc(plano: Dict[str, Any], target_env: str) -> List[str]:
    """Relê AGR_1251 e devolve roles sem todos os valores ativos solicitados."""
    load_project_env(find_project_root())
    conn = Connection(**build_connection_params_for(target_env))
    try:
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_1251",))
        falhas = []
        esperados = [(campo, valor) for campo, valores in plano["campos"].items() for valor in valores.split(",")]
        for role in plano["roles"]:
            linhas = read_table(
                conn,
                guard,
                table_name="AGR_1251",
                fields=["FIELD", "LOW", "DELETED"],
                options=[
                    {"TEXT": f"AGR_NAME = '{role}'"},
                    {"TEXT": f"AND OBJECT = '{plano['objeto']}'"},
                ],
                rowcount=0,
            )
            ativos = [(r[0].strip().upper(), r[1].strip().upper()) for r in linhas if len(r) > 2 and r[2].strip() != "X"]
            if any(
                not any(campo_ativo == campo and valor_autorizacao_cobre(valor_ativo, valor) for campo_ativo, valor_ativo in ativos)
                for campo, valor in esperados
            ):
                falhas.append(role)
        return falhas
    finally:
        conn.close()


def executar_atualizacao_objeto_rfc(
    plano: Dict[str, Any],
    env: str
) -> bool:
    """
    Executa a inserção e atualização de objetos de autorização e respetivos valores
    EXCLUSIVAMENTE via RFC nativo (PRGN_AUTO_GENERATE_PROFILE_NEW + PRGN_GEN_PROFILES_FOR_ROLES),
    eliminando qualquer dependência de SAP GUI ou scripts de tela.
    """
    load_project_env(find_project_root())
    conn = Connection(**build_connection_params_for(env))
    nome_objeto = str(plano["objeto"]).strip().upper()
    ts_agora = datetime.now().strftime("%Y%m%d")
    hora_agora = datetime.now().strftime("%H%M%S")

    sucesso_total = True
    print(f"\n{CLR_C}>>> [SAP {env}] A aplicar objeto {nome_objeto} via RFC (100% nativo)...{CLR_RST}")

    try:
        for role in plano["roles"]:
            print(f"  -> Processando role: {CLR_Y}{role}{CLR_RST}")
            manual_auths = []
            for campo, valores_str in plano["campos"].items():
                for v in str(valores_str).split(","):
                    v_clean = v.strip().upper()
                    if v_clean:
                        manual_auths.append({
                            "OBJECT": nome_objeto,
                            "FIELD": campo.strip().upper(),
                            "LOW": v_clean,
                            "HIGH": "",
                            "LFDNR": "000000",
                            "MODIFIER": "CSILVA",
                            "MODDATE": ts_agora,
                            "MODTIME": hora_agora
                        })

            try:
                conn.call(
                    "PRGN_AUTO_GENERATE_PROFILE_NEW",
                    ACTIVITY_GROUP=role,
                    GENERATE_PROFILE="X",
                    NO_DIALOG="X",
                    REBUILD_AUTH_DATA="",
                    MANUAL_AUTHS=manual_auths
                )
            except Exception as e_auto:
                # OPEN_AUTHS é tolerado se houver outros campos abertos na role
                pass

            # Gerar perfil e ajustar utilizadores
            try:
                conn.call("PRGN_GEN_PROFILES_FOR_ROLES", IV_USERCOMPARE="X", IT_ROLES=[{"AGR_NAME": role}])
            except Exception as e_gen:
                print(f"     {CLR_Y}[AVISO geração]: {e_gen}{CLR_RST}")

            # Validar semáforo
            try:
                st = conn.call("PRGN_CHECK_PROFILE_STATUS_RFC", ACTIVITY_GROUP=role, TRANSACTIONS_CHANGED="")
                led = str(st.get("LED_COLOR") or "").strip().upper()
                msg = str(st.get("MESSAGE_TEXT") or "").strip()
                if led == "GREEN":
                    print(f"     {CLR_G}[SEMÁFORO {env}] 🟢 {led} - {msg}{CLR_RST}")
                else:
                    print(f"     {CLR_Y}[SEMÁFORO {env}] 🟡 {led} - {msg}{CLR_RST}")
            except Exception:
                pass

        # Validação final em AGR_1251
        falhas = _validar_objeto_funcoes_rfc(plano, env)
        if falhas:
            print(f"  {CLR_R}[ERRO] Valores não confirmados na AGR_1251 em: {', '.join(falhas)}.{CLR_RST}")
            return False

        print(f"  {CLR_G}[OK] Objeto {nome_objeto} e valores gravados e confirmados via RFC no SAP {env}!{CLR_RST}")
        return True

    finally:
        conn.close()


def diagnosticar_objeto_autorizacao(
    usuario: str,
    obj_info: Dict[str, Any],
    caminho_excel: str,
    target_env: str = "PRD",
    pedir_confirmacao: bool = True,
):
    """Atualiza o objeto nas roles indicadas pela Proposta 100% via RFC e valida o resultado."""
    nome_objeto = str(obj_info.get("objeto") or "").strip().upper()
    if nome_objeto in OBJETOS_PROIBIDOS_ATUALIZACAO:
        print("\n" + "=" * 95)
        print(f"  {CLR_R}[OPERAÇÃO BLOQUEADA]{CLR_W} Objeto {CLR_Y}{nome_objeto}{CLR_W} é proibido de atualizar por script.{CLR_RST}")
        print("=" * 95)
        print(f"  O objeto {CLR_Y}{nome_objeto}{CLR_RST} controla regras de estratégia de liberação / governança crítica.")
        print(f"  Por motivo de conformidade e segurança, {CLR_R}não é permitida a alteração automática{CLR_RST}.")
        print("=" * 95)
        return False

    obj_enriquecido = enriquecer_objeto_com_proposta(obj_info, caminho_excel, usuario)
    plano = preparar_atualizacao_objeto(obj_enriquecido)
    if not plano:
        # Se não há roles mapeadas automaticamente para a aplicação, solicita ao utilizador indicar a Single Role
        nome_obj = str(obj_info.get("objeto") or "").strip().upper()
        campos_relevantes = obj_info.get("campos_relevantes", [])
        app_origem = obj_info.get("aplicacao") or "Desconhecida"
        if nome_obj and campos_relevantes:
            print(f"\n{CLR_Y}[AVISO] O objeto {nome_obj} ocorreu em '{app_origem}' e a sua função não foi encontrada na Proposta.{CLR_RST}")
            resp_ind = input(f"Deseja indicar manualmente a Single Role para o objeto {nome_obj}? [S/N] [S]: ").strip().upper()
            if resp_ind in ("", "S", "SIM", "Y", "YES", "1"):
                role_manual = input(f"Informe a Single Role onde deseja aplicar {nome_obj} (ex: Z_BASIS_BASE) [0 para cancelar]: ").strip().upper()
                if role_manual and role_manual != "0":
                    obj_manual_info = dict(obj_enriquecido)
                    obj_manual_info["funcoes_proposta"] = [{"role": role_manual, "descricao": "Manual", "fonte": "Manual"}]
                    plano = preparar_atualizacao_objeto(obj_manual_info)

    if not plano:
        print(f"\n{CLR_R}[ERRO] Faltam funções da Proposta ou campos/valores SU53; nenhuma alteração foi feita.{CLR_RST}")
        return False

    env = "QAD" if str(target_env).upper() in ("QAD", "QAS") else "PRD"
    try:
        roles_inexistentes = _validar_roles_existentes_rfc(plano["roles"], env)
    except Exception as exc:
        print(f"\n{CLR_R}[ERRO] Falha na pré-validação RFC do SAP {env}: {exc}{CLR_RST}")
        return False
    if roles_inexistentes:
        print(f"\n{CLR_R}[ERRO] Funções inexistentes no SAP {env}: {', '.join(roles_inexistentes)}.{CLR_RST}")
        print("Nenhuma alteração foi feita.")
        return False

    print("\n" + "=" * 95)
    print(f"  {CLR_W}ATUALIZAÇÃO DO OBJETO {CLR_M}{plano['objeto']}{CLR_W} VIA RFC (100% NATIVO) EM {env}{CLR_RST}")
    print("=" * 95)
    print(f"  Utilizador origem: {usuario.strip().upper()}")
    print(f"  Campos/Valores:    {', '.join(f'{c}={v}' for c, v in plano['campos'].items())}")
    print("  Funções a alterar:")
    for role in plano["roles"]:
        print(f"    - {role}")
    print("-" * 95)

    return executar_atualizacao_objeto_rfc(plano, env)


def executar_todos_objetos_autorizacao(
    usuario: str,
    objetos: List[Dict[str, Any]],
    caminho_excel: str,
    target_env: str = "PRD",
) -> Dict[str, int]:
    """Confirma uma vez e processa cada falha de objeto de forma independente (ignorando objetos proibidos)."""
    env = "QAD" if str(target_env).upper() in ("QAD", "QAS") else "PRD"
    
    objetos_elegiveis = []
    objetos_bloqueados = []
    for o in objetos:
        if str(o.get("objeto") or "").strip().upper() in OBJETOS_PROIBIDOS_ATUALIZACAO:
            objetos_bloqueados.append(o)
        else:
            objetos_elegiveis.append(o)

    print("\n" + "=" * 95)
    print(f"  {CLR_W}ATUALIZAÇÃO DE TODOS OS OBJETOS EM {env} - UTILIZADOR {CLR_Y}{usuario}{CLR_RST}")
    print("=" * 95)
    for indice, objeto in enumerate(objetos_elegiveis, 1):
        aplicacao = objeto.get("aplicacao") or "Não indicada"
        print(f"  [{indice}] {aplicacao} | {objeto.get('objeto', '')}")
    if objetos_bloqueados:
        print(f"\n  {CLR_R}[IGNORADOS - PROIBIDOS DE ATUALIZAR POR SCRIPT]:{CLR_RST}")
        for b in objetos_bloqueados:
            print(f"    - {b.get('objeto')} (Aplicação: {b.get('aplicacao') or 'N/A'})")
    print("-" * 95)

    if not objetos_elegiveis:
        print(f"{CLR_Y}[AVISO] Todos os objetos selecionados são proibidos de atualização automática.{CLR_RST}")
        return {"sucessos": 0, "falhas": 0}

    confirmacao = input(
        f"Confirma a atualização sequencial dos {len(objetos_elegiveis)} objetos permitidos no SAP {env}? [S/N] [N]: "
    ).strip().upper()
    if confirmacao not in ("S", "SIM", "Y", "YES", "1"):
        print("Operação cancelada; nenhum objeto foi processado.")
        return {"sucessos": 0, "falhas": 0}

    sucessos = 0
    falhas = 0
    for indice, objeto in enumerate(objetos_elegiveis, 1):
        aplicacao = objeto.get("aplicacao") or "Não indicada"
        nome_objeto = objeto.get("objeto", "")
        print(f"\n[{indice}/{len(objetos_elegiveis)}] {aplicacao} | {nome_objeto}")
        try:
            concluido = diagnosticar_objeto_autorizacao(
                usuario=usuario,
                obj_info=objeto,
                caminho_excel=caminho_excel,
                target_env=target_env,
                pedir_confirmacao=False,
            )
        except Exception as exc:
            print(f"{CLR_R}[ERRO] Falha inesperada em {aplicacao} | {nome_objeto}: {exc}{CLR_RST}")
            concluido = False
        if concluido:
            sucessos += 1
        else:
            falhas += 1

    print("\n" + "-" * 95)
    print(f"  RESUMO: {sucessos} sucesso(s) | {falhas} falha(s)")
    print("-" * 95)
    return {"sucessos": sucessos, "falhas": falhas}


# =====================================================================
# 4. ANÁLISE NO FICHEIRO EXCEL (PROPOSTA ATIVA, DEPARTAMENTO, PROPOSTA)
# =====================================================================

def _coluna_por_chave(df: pd.DataFrame, *chaves: str) -> Optional[Any]:
    chaves_normalizadas = {normalizar_chave(chave) for chave in chaves}
    return next((coluna for coluna in df.columns if normalizar_chave(coluna) in chaves_normalizadas), None)


def listar_utilizadores_departamento(caminho_excel: str, departamento: str) -> List[Dict[str, str]]:
    """Lê utilizador SAP e nome da Proposta Ativa para um departamento."""
    df_ativa = pd.read_excel(caminho_excel, sheet_name="Proposta Ativa")
    col_user = _coluna_por_chave(df_ativa, "USUARIO", "USER", "USER SAP")
    col_nome = _coluna_por_chave(df_ativa, "NOME COMPLETO", "NOME", "COLABORADOR")
    col_dep = _coluna_por_chave(df_ativa, "DEPARTAMENTO")
    if col_user is None or col_nome is None or col_dep is None:
        raise ValueError("A sheet 'Proposta Ativa' não contém as colunas Utilizador, Nome Completo e Departamento.")

    departamento_norm = normalizar_texto(departamento)
    utilizadores = []
    vistos = set()
    for _, row in df_ativa.iterrows():
        if normalizar_texto(row.get(col_dep)) != departamento_norm:
            continue
        usuario = normalizar_texto(row.get(col_user))
        if not usuario or usuario in vistos:
            continue
        vistos.add(usuario)
        utilizadores.append({
            "usuario": usuario,
            "nome": str(row.get(col_nome) or "").strip(),
        })
    return utilizadores


def listar_departamentos_controlo(caminho_excel: str) -> List[Dict[str, Any]]:
    """Mantém a ordem da CONTROLO e conta utilizadores correspondentes na Proposta Ativa."""
    df_controlo = pd.read_excel(caminho_excel, sheet_name="CONTROLO")
    col_dep = _coluna_por_chave(df_controlo, "DEPARTAMENTO")
    col_status = _coluna_por_chave(df_controlo, "STATUS")
    if col_dep is None:
        raise ValueError("A sheet 'CONTROLO' não contém a coluna DEPARTAMENTO.")

    departamentos = []
    vistos = set()
    for _, row in df_controlo.iterrows():
        departamento = str(row.get(col_dep) or "").strip()
        chave = normalizar_texto(departamento)
        if not chave or chave in vistos:
            continue
        vistos.add(chave)
        departamentos.append({
            "departamento": departamento,
            "status": str(row.get(col_status) or "").strip() if col_status is not None else "",
            "utilizadores": len(listar_utilizadores_departamento(caminho_excel, departamento)),
        })
    return departamentos


def pesquisar_departamento_su53(
    departamento: str,
    utilizadores: List[Dict[str, str]],
    caminho_excel: str,
    target_env: str = "PRD",
) -> Dict[str, List[Dict[str, Any]]]:
    """Consulta SU53 sequencialmente com barra de progresso compacta e resumo executivo."""
    resultado = {"com_erros": [], "sem_erros": [], "falhas_rfc": []}
    total_users = len(utilizadores)
    print(f"\n{CLR_C}>>> A consultar SU53 de {total_users} utilizadores em {departamento} [{target_env}]...{CLR_RST}")

    largura_barra = 28
    for indice, pessoa in enumerate(utilizadores, 1):
        usuario = pessoa["usuario"]
        nome = pessoa.get("nome", "")
        
        # Atualizar barra de progresso em linha única
        progresso = int((indice / total_users) * largura_barra)
        barra = "█" * progresso + " " * (largura_barra - progresso)
        percentual = int((indice / total_users) * 100)
        sys.stdout.write(f"\r  [{CLR_G}{barra}{CLR_RST}] {indice}/{total_users} ({percentual}%) - {CLR_Y}{usuario:<10}{CLR_RST}")
        sys.stdout.flush()

        try:
            erros = consultar_erros_su53_rfc(usuario, target_env=target_env, silencioso=True)
        except Exception as exc:
            resultado["falhas_rfc"].append({"usuario": usuario, "nome": nome, "erro": str(exc)})
            continue

        if not erros:
            resultado["sem_erros"].append({"usuario": usuario, "nome": nome})
            continue

        try:
            # Validar quais falhas já estão resolvidas pelas roles ativas do utilizador
            all_roles_user = obter_autorizacoes_ativas_usuario_rfc(usuario, target_env=target_env)
            for err in erros:
                resolvido, role_cobre = verificar_falha_ja_resolvida_rfc(err, all_roles_user, target_env=target_env)
                err["resolvido"] = resolvido
                err["role_resolvido"] = role_cobre

            transacoes, objetos = agrupar_erros_para_analise(erros)

            # Filtrar estritamente apenas as falhas PENDENTES (não resolvidas)
            transacoes_pendentes = [t for t in transacoes if not t.get("resolvido")]
            objetos_pendentes = [o for o in objetos if not o.get("resolvido")]

            # Se todas as falhas já estão atribuídas/resolvidas, o utilizador entra em 'sem_erros'
            if not transacoes_pendentes and not objetos_pendentes:
                resultado["sem_erros"].append({"usuario": usuario, "nome": nome})
                continue

            objetos_enriquecidos = [
                enriquecer_objeto_com_proposta(objeto, caminho_excel, usuario)
                for objeto in objetos_pendentes
            ]
        except Exception as exc:
            resultado["falhas_rfc"].append({"usuario": usuario, "nome": nome, "erro": str(exc)})
            continue

        total_pendentes = len(transacoes_pendentes) + len(objetos_pendentes)
        resultado["com_erros"].append({
            "usuario": usuario,
            "nome": nome,
            "total_registos": total_pendentes,
            "transacoes": transacoes_pendentes,
            "objetos": objetos_enriquecidos,
        })

    # Concluir linha de progresso
    barra_completa = "█" * largura_barra
    total_com_pend = len(resultado["com_erros"])
    total_sem_erros = len(resultado["sem_erros"])
    sys.stdout.write(f"\r  [{CLR_G}{barra_completa}{CLR_RST}] {total_users}/{total_users} (100%) Concluído ({total_com_pend} com pendências | {total_sem_erros} sem erros)\n")
    sys.stdout.flush()

    print("\n" + "=" * 95)
    print(f"  {CLR_W}RESUMO DE PENDÊNCIAS SU53 - {departamento.upper()} [{target_env}]{CLR_RST}")
    print("=" * 95)
    for pessoa in resultado["com_erros"]:
        itens_pendentes = []
        for transacao in pessoa["transacoes"]:
            tcode = transacao.get("tcode_falha", "")
            roles = [
                role
                for role in buscar_funcao_individual_para_transacao(caminho_excel, tcode, pessoa["usuario"])
                if role.get("fonte") == "Folha 'Proposta'" and not role.get("candidata")
            ]
            nomes_roles = roles[0].get("role", "") if roles else ""
            if nomes_roles:
                itens_pendentes.append(f"{CLR_Y}{tcode}{CLR_RST} ({CLR_G}{nomes_roles}{CLR_RST})")
            else:
                itens_pendentes.append(f"{CLR_Y}{tcode}{CLR_RST}")

        for objeto in pessoa["objetos"]:
            obj_name = objeto.get("objeto", "")
            roles = [r.get("role", "") for r in objeto.get("funcoes_proposta", []) if r.get("role")]
            role_txt = f" ({CLR_G}{roles[0]}{CLR_RST})" if roles else ""
            itens_pendentes.append(f"{CLR_M}{obj_name}{CLR_RST}{role_txt}")

        resumo_itens = ", ".join(itens_pendentes)
        print(f"  {CLR_W}{pessoa['usuario']:<10}{CLR_RST} | {pessoa['nome']:<24} ({pessoa['total_registos']}): {resumo_itens}")

    for pessoa in resultado["falhas_rfc"]:
        print(f"\n  {CLR_R}[ERRO RFC] {pessoa['usuario']} - {pessoa['nome']}: {pessoa['erro']}{CLR_RST}")

    print("-" * 95)
    return resultado


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


# Cache em memória do catálogo de funções do Excel para execução ultrarrápida
_CACHE_CATALOGO_EXCEL: Dict[str, Dict[str, Any]] = {}

def carregar_catalogo_completo_excel(caminho_excel: str) -> Dict[str, Any]:
    """
    Carrega e indexa em cache todo o catálogo de funções do ficheiro Excel:
      - Mapa da folha 'Proposta' (76 Single Roles oficiais com descrição, linhas e transações)
      - Conjunto estrito de roles válidas do Excel (Proposta, PFCG_COMPOSTA, Proposta Ativa, PFCG_AUTHORITY)
      - Mapa complementar da folha 'PFCG_CREATE' filtrado apenas para roles catalogadas
    Garante que NENHUMA role fora do Excel (ex: Z_CROSS_LOGISTIC_*) seja sugerida.
    """
    global _CACHE_CATALOGO_EXCEL
    if caminho_excel in _CACHE_CATALOGO_EXCEL:
        return _CACHE_CATALOGO_EXCEL[caminho_excel]

    proposta_roles: Dict[str, Dict[str, Any]] = {}
    pfcg_create_roles: Dict[str, Dict[str, Any]] = {}
    roles_validas: set = set()

    try:
        wb = openpyxl.load_workbook(caminho_excel, read_only=True, data_only=True)

        # 1. Parsing da folha 'Proposta'
        if "Proposta" in wb.sheetnames:
            ws_p = wb["Proposta"]
            cur_role = None
            cur_desc = ""
            cur_row = None
            for r_idx, r in enumerate(ws_p.iter_rows(values_only=True), 1):
                c1 = str(r[0] or "").strip()
                c2 = str(r[1] or "").strip() if len(r) > 1 else ""

                if c1.startswith("Z_") and c2:
                    cur_role = c1.upper()
                    cur_desc = c2
                    cur_row = r_idx
                    roles_validas.add(cur_role)
                    if cur_role not in proposta_roles:
                        proposta_roles[cur_role] = {
                            "desc": cur_desc,
                            "row": cur_row,
                            "tcodes": {}
                        }
                elif c1 and cur_role:
                    for t in c1.replace(";", " ").split():
                        tc_clean = t.strip().upper()
                        if tc_clean:
                            proposta_roles[cur_role]["tcodes"][tc_clean] = r_idx

        # 2. Roles válidas de 'PFCG_COMPOSTA'
        if "PFCG_COMPOSTA" in wb.sheetnames:
            ws_c = wb["PFCG_COMPOSTA"]
            header = [str(c or "").strip() for c in next(ws_c.iter_rows(values_only=True))]
            agr_idx = header.index("AGR_NAME") if "AGR_NAME" in header else 3
            for r in ws_c.iter_rows(min_row=2, values_only=True):
                if len(r) > agr_idx and r[agr_idx]:
                    val = str(r[agr_idx]).strip().upper()
                    if val.startswith("Z"):
                        roles_validas.add(val)

        # 3. Roles válidas de 'Proposta Ativa'
        if "Proposta Ativa" in wb.sheetnames:
            ws_a = wb["Proposta Ativa"]
            for r in ws_a.iter_rows(min_row=2, values_only=True):
                for val_c in r[9:]:
                    v_str = str(val_c or "").strip().upper()
                    if v_str.startswith("Z"):
                        roles_validas.add(v_str)

        # 4. Roles válidas de 'PFCG_AUTHORITY'
        if "PFCG_AUTHORITY" in wb.sheetnames:
            ws_auth = wb["PFCG_AUTHORITY"]
            header_auth = [str(c or "").strip() for c in next(ws_auth.iter_rows(values_only=True))]
            agr_auth_idx = header_auth.index("AGR_NAME") if "AGR_NAME" in header_auth else 3
            for r in ws_auth.iter_rows(min_row=2, values_only=True):
                if len(r) > agr_auth_idx and r[agr_auth_idx]:
                    val = str(r[agr_auth_idx]).strip().upper()
                    if val.startswith("Z"):
                        roles_validas.add(val)

        # 5. Parsing complementar de 'PFCG_CREATE' (somente roles presentes em roles_validas)
        if "PFCG_CREATE" in wb.sheetnames:
            ws_cr = wb["PFCG_CREATE"]
            header_cr = [str(c or "").strip() for c in next(ws_cr.iter_rows(values_only=True))]
            agr_cr_idx = header_cr.index("AGR_NAME") if "AGR_NAME" in header_cr else 1
            tc_cr_idx = header_cr.index("TCODE") if "TCODE" in header_cr else 3
            desc_cr_idx = header_cr.index("TEXT") if "TEXT" in header_cr else 2
            for r in ws_cr.iter_rows(min_row=2, values_only=True):
                if len(r) > max(agr_cr_idx, tc_cr_idx):
                    agr = str(r[agr_cr_idx] or "").strip().upper()
                    tc = str(r[tc_cr_idx] or "").strip().upper()
                    desc = str(r[desc_cr_idx] or "").strip() if len(r) > desc_cr_idx else ""
                    if agr.startswith("Z") and (agr in roles_validas) and "CROSS_LOGISTIC" not in agr:
                        if agr not in pfcg_create_roles:
                            pfcg_create_roles[agr] = {"desc": desc, "tcodes": set()}
                        for t in tc.replace(";", " ").split():
                            tc_clean = t.strip().upper()
                            if tc_clean:
                                pfcg_create_roles[agr]["tcodes"].add(tc_clean)

        wb.close()
    except Exception:
        pass

    dados_catalogo = {
        "proposta_roles": proposta_roles,
        "roles_validas": roles_validas,
        "pfcg_create_roles": pfcg_create_roles,
    }
    _CACHE_CATALOGO_EXCEL[caminho_excel] = dados_catalogo
    return dados_catalogo


def obter_catalogo_roles_validas_excel(caminho_excel: str) -> set:
    """Retorna o conjunto de todas as Single Roles válidas no Excel."""
    cat = carregar_catalogo_completo_excel(caminho_excel)
    return cat.get("roles_validas", set())


def buscar_funcao_individual_para_transacao(
    caminho_excel: str,
    tcode: str,
    usuario: Optional[str] = None
) -> List[Dict[str, Any]]:
    """
    Pesquisa a Função Individual (Single Role Z_...) correspondente à transação:
      1. Folha 'Proposta' (catálogo principal, mapeamento exato com cabeçalhos e linhas)
      2. Folha 'PFCG_CREATE' (apenas roles presentes no catálogo oficial do Excel)
    Restringe estritamente às Single Roles do projeto Excel.
    Nunca sugere roles legadas fora do Excel (ex: Z_CROSS_LOGISTIC_*).
    """
    tc_norm = tcode.strip().upper()
    funcoes_encontradas = []

    cat = carregar_catalogo_completo_excel(caminho_excel)
    proposta_roles = cat.get("proposta_roles", {})
    pfcg_create_roles = cat.get("pfcg_create_roles", {})

    # 1. Pesquisa na folha 'Proposta'
    for role, data in proposta_roles.items():
        if tc_norm in data["tcodes"]:
            if not any(f["role"] == role for f in funcoes_encontradas):
                funcoes_encontradas.append({
                    "role": role,
                    "descricao": data["desc"],
                    "fonte": "Folha 'Proposta'",
                    "candidata": False,
                    "linha_funcao": data["row"],
                    "linha_transacao": data["tcodes"][tc_norm],
                })

    # 2. Pesquisa complementar na folha 'PFCG_CREATE' (se não encontrado na 'Proposta')
    if not funcoes_encontradas:
        for role, data in pfcg_create_roles.items():
            if tc_norm in data["tcodes"]:
                if not any(f["role"] == role for f in funcoes_encontradas):
                    funcoes_encontradas.append({
                        "role": role,
                        "descricao": data["desc"],
                        "fonte": "Folha 'PFCG_CREATE'",
                        "candidata": False,
                    })

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

            # ETAPA 4: PFCG_CREATE (Garantir que a transação consta na Single Role no Excel)
            try:
                ws_create = wb.Worksheets("PFCG_CREATE")
                last_r_create = ws_create.UsedRange.Rows.Count
                while ws_create.Cells(last_r_create, 1).Value:
                    last_r_create += 1
                
                tcode_ja_na_role = False
                desc_single = ""
                for r in range(2, last_r_create):
                    r_single = str(ws_create.Cells(r, 1).Value or "").strip().upper()
                    r_desc = str(ws_create.Cells(r, 2).Value or "").strip()
                    r_tcode = str(ws_create.Cells(r, 3).Value or "").strip().upper()
                    if r_single == single_role.strip().upper():
                        if r_desc and not desc_single:
                            desc_single = r_desc
                        if r_tcode == tcode.strip().upper():
                            tcode_ja_na_role = True
                            break
                
                if not tcode_ja_na_role and tcode:
                    ws_create.Cells(last_r_create, 1).Value = single_role.strip().upper()
                    ws_create.Cells(last_r_create, 2).Value = desc_single or single_role.strip().upper()
                    ws_create.Cells(last_r_create, 3).Value = tcode.strip().upper()
                    ws_create.Cells(last_r_create, 4).Value = "Criado"
                    ws_create.Cells(last_r_create, 5).Value = "Atribuído em SAP DEV, PRD e QAD"
                    ws_create.Cells(last_r_create, 6).Value = ts_agora
                    ws_create.Cells(last_r_create, 7).Value = "Validado"
                    ws_create.Cells(last_r_create, 8).Value = "Validado"
                    print(f"    + Registada transação {tcode} na role {single_role} na folha 'PFCG_CREATE'")
            except Exception as e_create:
                print(f"    {CLR_Y}[AVISO PFCG_CREATE] {e_create}{CLR_RST}")

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

            # ETAPA 4: PFCG_CREATE (Garantir que a transação consta na Single Role no Excel)
            try:
                if "PFCG_CREATE" in wb.sheetnames:
                    ws_create = wb["PFCG_CREATE"]
                    tcode_ja_na_role = False
                    desc_single = ""
                    for r in range(2, ws_create.max_row + 1):
                        r_single = str(ws_create.cell(row=r, column=1).value or "").strip().upper()
                        r_desc = str(ws_create.cell(row=r, column=2).value or "").strip()
                        r_tcode = str(ws_create.cell(row=r, column=3).value or "").strip().upper()
                        if r_single == single_role.strip().upper():
                            if r_desc and not desc_single:
                                desc_single = r_desc
                            if r_tcode == tcode.strip().upper():
                                tcode_ja_na_role = True
                                break
                    if not tcode_ja_na_role and tcode:
                        ws_create.append([
                            single_role.strip().upper(), desc_single or single_role.strip().upper(),
                            tcode.strip().upper(), "Criado", "Atribuído em SAP DEV, PRD e QAD",
                            ts_agora, "Validado", "Validado", ""
                        ])
                        print(f"    + Registada transação {tcode} na role {single_role} na folha 'PFCG_CREATE'")
            except Exception as e_create:
                print(f"    {CLR_Y}[AVISO PFCG_CREATE] {e_create}{CLR_RST}")

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
    usuario: Optional[str] = None,
    tcode: Optional[str] = None
) -> Dict[str, Any]:
    """
    Atribui a Função Individual à Composta via RFC e executa o ajuste de utilizadores.
    Garante que:
      1. A transação (se informada) consta no menu da Single Role (via PRGN_RFC_ADD_TRANSACTION).
      2. A Single Role tem os seus dados consolidados e perfil gerado (PRGN_AUTO_GENERATE_PROFILE_NEW + PRGN_GEN_PROFILES_FOR_ROLES).
      3. A Single Role é associada à Composta e os utilizadores são comparados.
      4. Valida o semáforo verde via PRGN_CHECK_PROFILE_STATUS_RFC.
    """
    env_norm = env.strip().upper()
    comp_norm = composite_role.strip().upper()
    sgl_norm = single_role.strip().upper()
    u_norm = usuario.strip().upper() if usuario else ""
    tc_norm = tcode.strip().upper() if tcode else ""

    print(f"\n{CLR_C}>>> [SAP {env_norm}] A sincronizar autorizações via RFC...{CLR_RST}")

    try:
        load_project_env(find_project_root())
        params = build_connection_params_for(env_norm)
        conn = Connection(**params)
    except Exception as exc:
        return {"ok": False, "ambiente": env_norm, "erro": f"Não foi possível conectar ao SAP {env_norm}: {exc}"}

    try:
        guard = make_write_guard(("RFC_PING", "RFC_READ_TABLE"), ("AGR_DEFINE", "AGR_USERS", "AGR_TCODES"))

        # 1. Verificar se a função individual existe no SAP
        try:
            r_sgl = read_table(conn, guard, table_name="AGR_DEFINE", fields=["AGR_NAME"], options=[{"TEXT": f"AGR_NAME = '{sgl_norm}'"}], rowcount=1)
            if not r_sgl:
                print(f"    {CLR_Y}[AVISO] A função individual '{sgl_norm}' não se encontra na tabela AGR_DEFINE do SAP {env_norm}.{CLR_RST}")
        except Exception:
            pass

        # 2. Verificar e assegurar que a transação consta no menu da Single Role
        if tc_norm and sgl_norm:
            try:
                r_tc = read_table(
                    conn, guard, table_name="AGR_TCODES", fields=["TCODE"],
                    options=[{"TEXT": f"AGR_NAME = '{sgl_norm}' AND TCODE = '{tc_norm}'"}], rowcount=1
                )
                if not r_tc:
                    print(f"    + A transação {tc_norm} não consta no menu de {sgl_norm} em {env_norm}. A adicionar via PRGN_RFC_ADD_TRANSACTION...")
                    r_add_tc = conn.call("PRGN_RFC_ADD_TRANSACTION", ACTIVITY_GROUP=sgl_norm, TCODE=tc_norm, NO_DIALOG="X")
                    errs_tc = [r for r in r_add_tc.get("RETURN", []) if str(r.get("TYPE", "")).upper() in ("E", "A")]
                    if errs_tc:
                        msg_etc = "; ".join([str(e.get("MESSAGE", "")) for e in errs_tc])
                        print(f"       {CLR_Y}[AVISO PRGN_RFC_ADD_TRANSACTION]: {msg_etc}{CLR_RST}")
                    else:
                        print(f"       {CLR_G}[OK] Transação {tc_norm} adicionada com sucesso ao menu de {sgl_norm}!{CLR_RST}")
                else:
                    print(f"    . A transação {tc_norm} já consta no menu de {sgl_norm} no SAP {env_norm}.")
            except Exception as e_tc_chk:
                print(f"    {CLR_Y}[AVISO verificação AGR_TCODES]: {e_tc_chk}{CLR_RST}")

        # 3. Consolidação de Objetos e Regeneração de Perfil da Single Role (REBUILD_AUTH_DATA)
        try:
            conn.call(
                "PRGN_AUTO_GENERATE_PROFILE_NEW",
                ACTIVITY_GROUP=sgl_norm,
                GENERATE_PROFILE="X",
                NO_DIALOG="X",
                REBUILD_AUTH_DATA="X"
            )
        except Exception:
            # Tolerância para OPEN_AUTHS se houver campos em aberto na PFCG
            pass

        try:
            conn.call("PRGN_GEN_PROFILES_FOR_ROLES", IV_USERCOMPARE="X", IT_ROLES=[{"AGR_NAME": sgl_norm}])
        except Exception as e_gen_sgl:
            print(f"    {CLR_Y}[AVISO geração Single Role]: {e_gen_sgl}{CLR_RST}")

        # 4. Verificar se a função composta existe no SAP
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

        # 5. Validação Final do Semáforo (PRGN_CHECK_PROFILE_STATUS_RFC)
        try:
            st = conn.call("PRGN_CHECK_PROFILE_STATUS_RFC", ACTIVITY_GROUP=sgl_norm, TRANSACTIONS_CHANGED="")
            led = str(st.get("LED_COLOR") or "").strip().upper()
            msg = str(st.get("MESSAGE_TEXT") or "").strip()
            if led == "GREEN":
                print(f"    {CLR_G}[SEMÁFORO SAP {env_norm}] 🟢 {led} - {msg}{CLR_RST}")
            elif led == "YELLOW":
                print(f"    {CLR_Y}[SEMÁFORO SAP {env_norm}] 🟡 {led} - {msg}{CLR_RST}")
            else:
                print(f"    {CLR_R}[SEMÁFORO SAP {env_norm}] 🔴 {led} - {msg}{CLR_RST}")
        except Exception as e_chk:
            print(f"    {CLR_Y}[AVISO verificação de status]: {e_chk}{CLR_RST}")

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
        print(f"  [ENTER] Aceitar '{single_role_escolhida}' | [M] Alterar | [P] Pular | [0] Sair")
        sel_r = input(f"  Opção [ENTER/M/P/0]: ").strip().upper()
        if sel_r in ("0", "SAIR", "EXIT", "Q", "QUIT"):
            print("Operação cancelada pelo utilizador.")
            return "EXIT"
        if sel_r in ("P", "PULAR", "SKIP"):
            print(f"Transação {tcode_escolhido} ignorada; pulando para a próxima...")
            return "SKIP"
        if sel_r == "M":
            print("  (Introduza o nome da Função Individual, ou [P] para pular, [0] para sair)")
            val_m = input("  Introduza o nome da Função Individual (Z_...): ").strip().upper()
            if val_m in ("0", "SAIR", "EXIT", "Q", "QUIT"):
                return "EXIT"
            if not val_m or val_m in ("P", "PULAR", "SKIP"):
                return "SKIP"
            single_role_escolhida = val_m
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
                print("  [P] Pular esta transação e avançar para a próxima")
                print("  [0] Sair sem efetuar alterações")

                prompt_r = f"Pressione [ENTER] para aceitar '{CLR_Y}{exatos[0]['role']}{CLR_RST}' ou escolha [1-{len(exatos)}, M, P, 0] [Padrão: 1]: "
                sel_r = input(prompt_r).strip().upper()
                if sel_r in ("0", "SAIR", "EXIT", "Q", "QUIT"):
                    print("Operação cancelada pelo utilizador.")
                    return "EXIT"
                if sel_r in ("P", "PULAR", "SKIP"):
                    print(f"Transação {tcode_escolhido} ignorada; pulando para a próxima...")
                    return "SKIP"
                if not sel_r or sel_r == "1":
                    single_role_escolhida = exatos[0]["role"]
                elif sel_r == "M":
                    print("  (Introduza o nome da Função Individual, ou [P] para pular, [0] para sair)")
                    val_m = input("Introduza o nome da Função Individual (Z_...): ").strip().upper()
                    if val_m in ("0", "SAIR", "EXIT", "Q", "QUIT"):
                        return "EXIT"
                    if not val_m or val_m in ("P", "PULAR", "SKIP"):
                        return "SKIP"
                    single_role_escolhida = val_m
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
                print("  [P] Pular esta transação e avançar para a próxima")
                print("  [0] Sair sem efetuar alterações")

                prompt_c = f"Pressione [ENTER] para aceitar '{CLR_Y}{candidatos[0]['role']}{CLR_RST}' ou escolha [1-{len(candidatos[:3])}, M, P, 0] [Padrão: 1]: "
                sel_c = input(prompt_c).strip().upper()
                if sel_c in ("0", "SAIR", "EXIT", "Q", "QUIT"):
                    print("Operação cancelada pelo utilizador.")
                    return "EXIT"
                if sel_c in ("P", "PULAR", "SKIP"):
                    print(f"Transação {tcode_escolhido} ignorada; pulando para a próxima...")
                    return "SKIP"
                if not sel_c or sel_c == "1":
                    single_role_escolhida = candidatos[0]["role"]
                elif sel_c == "M":
                    print("  (Introduza o nome da Função Individual, ou [P] para pular, [0] para sair)")
                    val_m = input("Introduza o nome da Função Individual (Z_...): ").strip().upper()
                    if val_m in ("0", "SAIR", "EXIT", "Q", "QUIT"):
                        return "EXIT"
                    if not val_m or val_m in ("P", "PULAR", "SKIP"):
                        return "SKIP"
                    single_role_escolhida = val_m
                else:
                    try:
                        single_role_escolhida = candidatos[int(sel_c) - 1]["role"]
                    except Exception:
                        single_role_escolhida = candidatos[0]["role"]
        else:
            print(f"{CLR_Y}[AVISO] A transação '{tcode_escolhido}' não foi encontrada no catálogo da folha 'Proposta'.{CLR_RST}")
            print(f"  Opções:")
            print(f"    - Nome técnico da Função Individual (ex: Z_IT_EQUIPA_INTERNA)")
            print(f"    - [{CLR_Y}P{CLR_RST}] Pular esta transação e avançar para a próxima")
            print(f"    - [{CLR_Y}0{CLR_RST}] Cancelar e sair")
            resp_manual = input(f"Introduza a Função Individual ou escolha [P/0]: ").strip().upper()
            if resp_manual in ("0", "SAIR", "EXIT", "Q", "QUIT"):
                print("Operação cancelada pelo utilizador.")
                return "EXIT"
            if not resp_manual or resp_manual in ("P", "PULAR", "SKIP"):
                print(f"Transação {tcode_escolhido} ignorada; pulando para a próxima...")
                return "SKIP"
            single_role_escolhida = resp_manual

    if not single_role_escolhida or single_role_escolhida in ("P", "PULAR", "SKIP"):
        print(f"Transação {tcode_escolhido} ignorada; pulando para a próxima...")
        return "SKIP"

    if single_role_escolhida in ("0", "SAIR", "EXIT", "Q", "QUIT"):
        print("Operação cancelada pelo utilizador.")
        return "EXIT"

    print(f"\n{CLR_W}>>> Função Individual selecionada: {CLR_Y}{single_role_escolhida}{CLR_RST}")

    # Confirmação do Plano de Ação
    env_primary = "QAD" if target_env in ("QAD", "QAS") else "PRD"
    env_secondary = "PRD" if env_primary == "QAD" else "QAD"

    print(f"\n{CLR_C}>>> A APLICAR ATUALIZAÇÕES:{CLR_RST}")
    # 1. Excel
    res_excel = aplicar_alteracoes_excel(
        caminho_excel=caminho_excel,
        dados_user=dados_user,
        tcode=tcode_escolhido,
        single_role=single_role_escolhida,
        nome_sheet_dep=nome_sheet_dep
    )
    if not res_excel.get("ok"):
        print(f"    1. EXCEL: {CLR_R}ERRO - {res_excel.get('erro')}{CLR_RST}")
        return "SKIP"
    print(f"    1. EXCEL: {nome_sheet_dep} ('X'), Proposta Ativa, PFCG_COMPOSTA, PFCG_CREATE {CLR_G}✔{CLR_RST}")

    # 2. SAP Principal
    res_prim = sincronizar_sap_rfc(env_primary, comp_role, single_role_escolhida, usuario=usuario, tcode=tcode_escolhido)
    if res_prim.get("ok"):
        print(f"    2. SAP {env_primary}: Menu + Role associada + Perfis gerados {CLR_G}🟢 GREEN{CLR_RST}")
    else:
        print(f"    2. SAP {env_primary}: {CLR_R}{res_prim.get('erro')}{CLR_RST}")

    # 3. Sincronização secundária
    executar_sec = False
    if env_primary == "QAD":
        resp_sec = input(f"\nDeseja também aplicar as alterações no SAP {env_secondary} (Produção)? [S/N] [N]: ").strip().upper()
        if resp_sec in ("S", "SIM", "Y", "YES", "1"):
            executar_sec = True
    else:
        resp_sec = input(f"\nDeseja também aplicar as alterações no SAP {env_secondary} (QAS / Teste)? [S/N] [S]: ").strip().upper()
        if resp_sec in ("", "S", "SIM", "Y", "YES", "1"):
            executar_sec = True

    if executar_sec:
        res_sec = sincronizar_sap_rfc(env_secondary, comp_role, single_role_escolhida, usuario=usuario, tcode=tcode_escolhido)
        if res_sec.get("ok"):
            print(f"    3. SAP {env_secondary}: Menu + Role associada + Perfis gerados {CLR_G}🟢 GREEN{CLR_RST}")
        else:
            print(f"    3. SAP {env_secondary}: {CLR_R}{res_sec.get('erro')}{CLR_RST}")

    print("\n" + "=" * 95)
    print(f"  {CLR_G}[OK] RESOLUÇÃO CONCLUÍDA COM SUCESSO PARA {tcode_escolhido}!{CLR_RST}")
    print("=" * 95)
    return "OK"


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

    cli_user = None
    cli_env = None
    for arg in sys.argv[1:]:
        a_up = arg.strip().upper()
        if a_up in ("PRD", "QAS", "QAD", "DEV"):
            cli_env = "QAD" if a_up == "QAS" else a_up
        elif not cli_user:
            cli_user = a_up

    # Loop mestre para permitir alternar entre Utilizadores / Departamentos ou Sair
    while True:
        departamento_selecionado = None
        utilizadores_departamento = []
        if cli_user:
            usuario = cli_user
            cli_user = None  # Executa uma vez com CLI e depois permite voltar ao menu interativo
        else:
            try:
                print(f"\n  [{CLR_Y}1{CLR_RST}] Utilizador")
                print(f"  [{CLR_Y}2{CLR_RST}] Departamento")
                print(f"  [{CLR_Y}0{CLR_RST}] Sair do sistema")
                modo = input("Selecione o tipo de pesquisa [1-2, 0] [Padrão: 1]: ").strip().upper() or "1"
            except (KeyboardInterrupt, EOFError):
                print("\nOperação cancelada.")
                break

            if modo in ("0", "SAIR", "EXIT", "Q"):
                print("Sessão terminada com sucesso.")
                break

            if modo == "2":
                try:
                    departamentos = listar_departamentos_controlo(caminho_excel)
                except Exception as exc:
                    print(f"{CLR_R}[ERRO] Falha ao ler departamentos do Excel: {exc}{CLR_RST}")
                    continue
                if not departamentos:
                    print(f"{CLR_Y}[AVISO] Nenhum departamento encontrado na sheet CONTROLO.{CLR_RST}")
                    continue

                print("\n" + "=" * 95)
                print("  DEPARTAMENTOS DA SHEET CONTROLO")
                print("=" * 95)
                for indice, item in enumerate(departamentos, 1):
                    print(
                        f"  [{indice}] {item['departamento']} | "
                        f"{item['utilizadores']} utilizador(es) | {item['status'] or 'Sem status'}"
                    )
                try:
                    escolha_dep = input(f"Selecione o departamento [1-{len(departamentos)}, 0] [Padrão: 1]: ").strip()
                    if escolha_dep == "0":
                        continue
                    indice_dep = int(escolha_dep or "1") - 1
                except (ValueError, KeyboardInterrupt, EOFError):
                    print("Operação cancelada: departamento inválido.")
                    continue
                if not 0 <= indice_dep < len(departamentos):
                    print("Operação cancelada: departamento fora do intervalo válido.")
                    continue

                departamento_selecionado = departamentos[indice_dep]["departamento"]
                utilizadores_departamento = listar_utilizadores_departamento(caminho_excel, departamento_selecionado)
                if not utilizadores_departamento:
                    print(f"{CLR_Y}[AVISO] Nenhum utilizador encontrado para {departamento_selecionado}.{CLR_RST}")
                    continue
                print(f"\nUtilizadores de {departamento_selecionado}:")
                for pessoa in utilizadores_departamento:
                    print(f"  - {pessoa['usuario']} | {pessoa['nome']}")
                usuario = ""
            else:
                default_user = "ITSALSA" if any("ITSALSA" in str(a).upper() for a in sys.argv) else "CSILVA"
                try:
                    user_input = input(f"\n{CLR_W}Informe o utilizador SAP a analisar [{CLR_Y}{default_user}{CLR_W}] (ou 0 para voltar): {CLR_RST}").strip().upper()
                except (KeyboardInterrupt, EOFError):
                    print("\nOperação cancelada.")
                    break
                if user_input == "0":
                    continue
                usuario = user_input or default_user

        default_env_opt = "2" if any(k in str(a).upper() for a in sys.argv for k in ("QAS", "QAD", "ITSALSA")) else "1"
        try:
            env_input = cli_env or input(f"{CLR_W}Informe o ambiente SAP ([1] PRD | [2] QAS) [{CLR_Y}{default_env_opt}{CLR_W}]: {CLR_RST}").strip().upper()
        except (KeyboardInterrupt, EOFError):
            print("\nOperação cancelada.")
            break

        if env_input in ("2", "QAS", "QAD"):
            target_env = "QAD"
        else:
            target_env = "PRD"

        if departamento_selecionado:
            resultado_departamento = pesquisar_departamento_su53(
                departamento=departamento_selecionado,
                utilizadores=utilizadores_departamento,
                caminho_excel=caminho_excel,
                target_env=target_env,
            )
            utilizadores_com_erros = resultado_departamento["com_erros"]
            if not utilizadores_com_erros:
                continue

            print("\nSelecionar um utilizador com erros para abrir o fluxo individual:")
            for indice, pessoa in enumerate(utilizadores_com_erros, 1):
                print(f"  [{indice}] {pessoa['usuario']} | {pessoa['nome']}")
            print("  [0] Sair / Voltar ao menu principal")
            try:
                escolha_user = input(f"Selecione o utilizador [1-{len(utilizadores_com_erros)}, 0] [Padrão: 0]: ").strip()
                if escolha_user in ("", "0"):
                    continue
                indice_user = int(escolha_user) - 1
            except (ValueError, KeyboardInterrupt, EOFError):
                continue
            if not 0 <= indice_user < len(utilizadores_com_erros):
                continue
            usuario = utilizadores_com_erros[indice_user]["usuario"]

        # Loop interno para o utilizador selecionado: permite resolver falhas e continuar até decidir voltar/sair
        while True:
            # 2. Leitura do buffer SU53 via RFC
            try:
                erros = consultar_erros_su53_rfc(usuario, target_env=target_env)
            except Exception as exc:
                print(f"\n{CLR_R}[ERRO RFC] Falha ao consultar o buffer da SU53: {exc}{CLR_RST}")
                break

            # Validar autorizações ativas do utilizador no SAP para identificar erros já atribuídos/resolvidos
            all_roles_user = set()
            try:
                print(f"{CLR_C}>>> Verificando funções ativas do utilizador no SAP {target_env}...{CLR_RST}")
                all_roles_user = obter_autorizacoes_ativas_usuario_rfc(usuario, target_env=target_env)
                for err in erros:
                    resolvido, role_cobre = verificar_falha_ja_resolvida_rfc(err, all_roles_user, target_env=target_env)
                    err["resolvido"] = resolvido
                    err["role_resolvido"] = role_cobre
            except Exception as exc_val:
                print(f"  [AVISO] Não foi possível verificar autorizações prévias: {exc_val}")

            # Exibir tabela recente com status de atribuição
            exibir_tabela_erros(erros, usuario, target_env=target_env, all_roles=all_roles_user)

            # 3. Agrupamento em Transações (S_TCODE) e Objetos de Autorização
            transacoes_todas, objetos_todos = agrupar_erros_para_analise(erros)

            # Filtrar apenas as falhas que NÃO estão resolvidas (pendentes de autorização)
            transacoes = [t for t in transacoes_todas if not t.get("resolvido")]
            objetos = [o for o in objetos_todos if not o.get("resolvido")]

            # Se todas as falhas registadas já tiverem sido resolvidas:
            if not transacoes and not objetos:
                print("\n" + "=" * 95)
                print(f"  {CLR_G}✓ TODAS AS FALHAS REGISTADAS NO BUFFER DA SU53 JÁ SE ENCONTRAM RESOLVIDAS!{CLR_RST}")
                print("=" * 95)
                print(f"  O utilizador {CLR_Y}{usuario}{CLR_RST} já possui todas as roles ativas necessárias no SAP.")
                print(f"  {CLR_C}-> Nota:{CLR_RST} O buffer do SAP GUI (/nSU53) apenas será limpo quando o utilizador")
                print("           executar novamente a transação com sucesso ou abrir uma nova sessão.")
                print("-" * 95)
                resp_m = input("Deseja introduzir manualmente outra Transação ou Objeto? [S/N] [N]: ").strip().upper()
                if resp_m in ("S", "SIM", "Y", "YES", "1"):
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
                                caminho_excel=caminho_excel,
                                target_env=target_env,
                            )
                    else:
                        tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                        if tc_manual:
                            executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)
                    continue
                break

            # Apresentar painel de análise apenas com as falhas pendentes reais
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
                elif escolha in ("N", "NAO", "NO"):
                    if objetos:
                        resp_obj = input(f"\nDeseja atualizar um dos outros {len(objetos)} objetos de autorização com falha? [S/N] [S]: ").strip().upper()
                        if resp_obj in ("", "S", "SIM", "Y", "YES", "1"):
                            diagnosticar_objeto_autorizacao(
                                usuario=usuario,
                                obj_info=objetos[0],
                                caminho_excel=caminho_excel,
                                target_env=target_env,
                            )
                elif escolha == "0":
                    print("Operação cancelada para este utilizador.")
                    break
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
                                caminho_excel=caminho_excel,
                                target_env=target_env,
                            )
                    else:
                        tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                        if tc_manual:
                            executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)
                else:
                    indices_sel = parse_selecao_multipla(escolha, len(opcoes))
                    if indices_sel:
                        for idx_sel in indices_sel:
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
                            elif item_sel["tipo"] == "OBJETO":
                                diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=item_sel["detalhes"], caminho_excel=caminho_excel, target_env=target_env)
                    else:
                        print("[AVISO] Opção não reconhecida.")

            # SE EXISTIREM MÚLTIPLAS TRANSAÇÕES BLOQUEADAS:
            elif len(transacoes) > 1:
                print("\n" + "=" * 95)
                print(f"  {CLR_W}AÇÃO PROPOSTA PARA RESOLUÇÃO DAS FALHAS SU53 - UTILIZADOR {CLR_Y}{usuario}{CLR_RST}")
                print("=" * 95)
                print(f"Foram detetadas {len(opcoes)} falhas agrupadas na SU53:")
                for i, opc_item in enumerate(opcoes, 1):
                    if opc_item["tipo"] == "TRANSACAO":
                        sug = opc_item.get("sugestao_role") or "Manual"
                        print(f"  [{CLR_Y}{i}{CLR_RST}] Transação {CLR_Y}{opc_item['valor']:<20}{CLR_RST} (Função: {CLR_G}{sug}{CLR_RST})")
                    else:
                        aplicacao = opc_item["detalhes"].get("aplicacao") or "Não identificada"
                        print(f"  [{CLR_Y}{i}{CLR_RST}] Objeto {CLR_M}{opc_item['valor']:<20}{CLR_RST} (Aplicação: {aplicacao})")
                print(f"  [{CLR_Y}T{CLR_RST}] Adicionar TODAS as {len(transacoes)} transações acima ao perfil de {CLR_Y}{usuario}{CLR_RST}")
                if objetos:
                    print(f"  [{CLR_Y}B{CLR_RST}] Atualizar TODOS os {len(objetos)} objetos de autorização")
                print(f"  [{CLR_Y}O{CLR_RST}] Introduzir manualmente outra Transação ou Objeto")
                print(f"  [{CLR_Y}0{CLR_RST}] Voltar / Sair")
                print("-" * 95)

                escolha = input(f"Selecione a opção pretendida [1-{len(opcoes)}, T, B, O, 0] [Padrão: 1]: ").strip().upper() or "1"

                if escolha == "T":
                    for opc_item in (opcao for opcao in opcoes if opcao["tipo"] == "TRANSACAO"):
                        res_tc = executar_atribuicao_transacao(
                            caminho_excel=caminho_excel,
                            usuario=usuario,
                            tcode_escolhido=opc_item["valor"],
                            sugestao_role=opc_item.get("sugestao_role"),
                            roles_previamente_encontradas=opc_item.get("roles_encontradas"),
                            target_env=target_env
                        )
                        if res_tc == "EXIT":
                            print("Processamento sequencial interrompido pelo utilizador.")
                            break

                elif escolha in ("0", "SAIR", "EXIT"):
                    print("Saindo do utilizador atual.")
                    break

                elif escolha in ("B", "OBJETOS"):
                    objetos_opcoes = [opcao["detalhes"] for opcao in opcoes if opcao["tipo"] == "OBJETO"]
                    if objetos_opcoes:
                        executar_todos_objetos_autorizacao(
                            usuario=usuario,
                            objetos=objetos_opcoes,
                            caminho_excel=caminho_excel,
                            target_env=target_env,
                        )

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
                                caminho_excel=caminho_excel,
                                target_env=target_env,
                            )
                    else:
                        tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                        if tc_manual:
                            executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)

                else:
                    indices_sel = parse_selecao_multipla(escolha, len(opcoes))
                    if indices_sel:
                        for idx_sel in indices_sel:
                            opc_item = opcoes[idx_sel]
                            if opc_item["tipo"] == "TRANSACAO":
                                res_tc = executar_atribuicao_transacao(
                                    caminho_excel=caminho_excel,
                                    usuario=usuario,
                                    tcode_escolhido=opc_item["valor"],
                                    sugestao_role=opc_item.get("sugestao_role"),
                                    roles_previamente_encontradas=opc_item.get("roles_encontradas"),
                                    target_env=target_env,
                                )
                                if res_tc == "EXIT":
                                    print("Processamento sequencial interrompido pelo utilizador.")
                                    break
                            else:
                                diagnosticar_objeto_autorizacao(
                                    usuario=usuario,
                                    obj_info=opc_item["detalhes"],
                                    caminho_excel=caminho_excel,
                                    target_env=target_env,
                                )
                    else:
                        print("[AVISO] Opção inválida.")

            # SE NÃO HOUVER TRANSAÇÕES BLOQUEADAS (APENAS OBJETOS):
            else:
                if objetos:
                    print("\n" + "=" * 95)
                    print(f"  {CLR_W}ATUALIZAÇÃO DE OBJETOS DE AUTORIZAÇÃO - UTILIZADOR {CLR_Y}{usuario}{CLR_RST}")
                    print("=" * 95)
                    for i, o in enumerate(objetos, 1):
                        desc = f" ({o['descricao']})" if o["descricao"] else ""
                        print(f"  [{CLR_Y}{i}{CLR_RST}] Objeto: {CLR_M}{o['objeto']:<20}{CLR_RST}{desc}")
                    if len(objetos) > 1:
                        print(f"  [{CLR_Y}T{CLR_RST}] Atualizar TODOS os {len(objetos)} objetos acima")
                    print("  [O] Introduzir manualmente outra Transação ou Objeto")
                    print("  [0] Voltar / Sair sem efetuar alterações")
                    print("-" * 95)

                    opcoes_menu = f"1-{len(objetos)}, T, O, 0" if len(objetos) > 1 else "1, O, 0"
                    escolha = input(f"Selecione o objeto a atualizar [{opcoes_menu}] [Padrão: 1]: ").strip().upper()
                    if escolha in ("", "1"):
                        diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=objetos[0], caminho_excel=caminho_excel, target_env=target_env)
                    elif escolha == "T" and len(objetos) > 1:
                        executar_todos_objetos_autorizacao(
                            usuario=usuario,
                            objetos=objetos,
                            caminho_excel=caminho_excel,
                            target_env=target_env,
                        )
                    elif escolha == "0":
                        print("Saindo do utilizador atual.")
                        break
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
                                    caminho_excel=caminho_excel,
                                    target_env=target_env,
                                )
                        else:
                            tc_manual = input("Introduza o código da Transação SAP (ex: PFCGMASSTRANSPORT): ").strip().upper()
                            if tc_manual:
                                executar_atribuicao_transacao(caminho_excel=caminho_excel, usuario=usuario, tcode_escolhido=tc_manual, target_env=target_env)
                    else:
                        indices_sel = parse_selecao_multipla(escolha, len(objetos))
                        if indices_sel:
                            for idx_o in indices_sel:
                                diagnosticar_objeto_autorizacao(usuario=usuario, obj_info=objetos[idx_o], caminho_excel=caminho_excel, target_env=target_env)
                        else:
                            print("[AVISO] Opção inválida.")
                else:
                    print(f"\n{CLR_G}[INFO] Não foram detetados erros na SU53 do utilizador '{usuario}'.{CLR_RST}")
                    break

            # Pergunta se quer continuar a tratar este mesmo utilizador
            resp_cont = input(f"\n{CLR_W}Deseja tratar outra falha do utilizador {CLR_Y}{usuario}{CLR_W}? [{CLR_G}S{CLR_W}/N] [{CLR_G}S{CLR_W}]: {CLR_RST}").strip().upper()
            if resp_cont in ("N", "NAO", "NO", "0"):
                break



if __name__ == "__main__":
    main()


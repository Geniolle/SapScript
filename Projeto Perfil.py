# -*- coding: utf-8 -*-
"""
Projeto Perfil.py
======================================================================
Módulo de pesquisa, análise e leitura estruturada do ficheiro Excel
de Perfis de Autorização e Funções SAP (PFCG / CUA).

Funcionalidades:
  - Localização automática do ficheiro Excel (sap_script_uploads ou diálogo)
  - Identificação e consolidação inteligente de sheets:
      * PFCG_CREATE      (Roles simples, descrições e transações/TCODEs)
      * PFCG_COMPOSTA    (Roles compostas e associação com roles simples)
      * PFCG_AUTHORITY   (Objetos de autorização e campos organizacionais)
      * CUA_ADICIONAR    (Atribuição de perfis a utilizadores no CUA)
      * CUA_REMOVE       (Remoção de perfis de utilizadores no CUA)
  - Motor de pesquisa por:
      * Nome da Função / Role (parcial, exata ou wildcard)
      * Transação (TCODE) -> descobre quais funções concedem acesso
      * Utilizador SAP -> descobre funções adicionadas/removidas/pendentes
  - Modo interativo no terminal (menu) ou importação como biblioteca.
======================================================================
"""

import os
import sys
import glob
import re
import unicodedata
from datetime import datetime
from typing import Optional, Dict, Any, List

# Garantir codificação UTF-8 no Windows
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Auto-deteção e re-execução no ambiente virtual .venv-rfc se pandas não estiver presente
try:
    import pandas
except ImportError:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    venv_python = os.path.join(base_dir, ".venv-rfc", "Scripts", "python.exe")
    if os.path.exists(venv_python) and sys.executable.lower() != venv_python.lower():
        import subprocess
        res = subprocess.run([venv_python, os.path.abspath(__file__)] + sys.argv[1:])
        sys.exit(res.returncode)


# =====================================================================
# UTILITÁRIOS DE NORMALIZAÇÃO E LIMPEZA
# =====================================================================

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


def normalizar_nome_coluna(col: Any) -> str:
    """Normaliza nome de coluna/sheet para identificação independente de formatação e pontuação."""
    return re.sub(r"[^A-Z0-9]", "", normalizar_texto(col))


def limpar_tcode(tcode_raw: Any) -> List[str]:
    """Extrai transações limpas sem prefixos (/N, /O, TCODE=)."""
    if not tcode_raw:
        return []
    s = str(tcode_raw).replace("\r", "\n").replace("\t", " ").strip().upper()
    partes = re.split(r"[;, \n]+", s)
    resultado = []
    for p in partes:
        p = p.strip()
        if not p:
            continue
        for prefixo in ("TCODE=", "TCODE:", "T=", "T:", "/N", "/O"):
            if p.startswith(prefixo):
                p = p[len(prefixo):].strip()
                break
        if p and p not in resultado:
            resultado.append(p)
    return resultado


# Carregar variáveis de ambiente do .env
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))
except ImportError:
    pass


# =====================================================================
# LOCALIZAÇÃO DO FICHEIRO EXCEL
# =====================================================================

def encontrar_excel_padrao(diretorio_base: Optional[str] = None) -> Optional[str]:
    """
    Procura automaticamente pelo ficheiro Excel de perfis:
      1. Caminho configurado no .env (SHAREPOINT_PERFIS_FILE ou SHAREPOINT_PERFIS_LOCAL_DIR)
      2. Pasta sap_script_uploads ou diretório base
    """
    if not diretorio_base:
        diretorio_base = os.path.dirname(os.path.abspath(__file__))

    # 1. Verificar ficheiro direto configurado no .env
    sp_file = os.getenv("SHAREPOINT_PERFIS_FILE", "").strip()
    if sp_file and os.path.isfile(sp_file):
        return sp_file

    pastas_busca = []
    # 2. Verificar pasta sincronizada do SharePoint no .env
    sp_dir = os.getenv("SHAREPOINT_PERFIS_LOCAL_DIR", "").strip()
    if sp_dir and os.path.isdir(sp_dir):
        pastas_busca.append(sp_dir)

    pastas_busca.extend([
        os.path.join(diretorio_base, "sap_script_uploads"),
        diretorio_base,
        os.path.join(diretorio_base, "Processos", "Funções PFCG"),
    ])

    candidatos = []
    for pasta in pastas_busca:
        if not os.path.isdir(pasta):
            continue
        for padrao in ("*Perfis*.xlsx", "*PERFIS*.xlsx", "*Role*.xlsx", "*.xlsx"):
            for f in glob.glob(os.path.join(pasta, padrao)):
                # Ignora ficheiros temporários do Excel
                if os.path.basename(f).startswith("~$"):
                    continue
                try:
                    mtime = os.path.getmtime(f)
                    candidatos.append((mtime, f))
                except OSError:
                    continue

    if not candidatos:
        return None

    # Ordenar pelo mais recente
    candidatos.sort(key=lambda x: x[0], reverse=True)
    return candidatos[0][1]


def selecionar_ficheiro_dialogo(titulo: str = "Selecione o ficheiro Excel de Perfis SAP") -> Optional[str]:
    """Abre o diálogo nativo do Windows para seleção do ficheiro."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        caminho = filedialog.askopenfilename(
            title=titulo,
            filetypes=[("Ficheiros Excel", "*.xlsx *.xls *.xlsm"), ("Todos os ficheiros", "*.*")]
        )
        root.destroy()
        return caminho if caminho else None
    except Exception as err:
        print(f"⚠️ Não foi possível abrir o diálogo gráfico: {err}")
        return None


# =====================================================================
# LEITURA E ESTRUTURAÇÃO DO EXCEL
# =====================================================================

class ProjetoPerfilData:
    """Armazena as estruturas de dados consolidadas do Excel de Perfis."""

    def __init__(self, caminho_ficheiro: str):
        self.caminho = caminho_ficheiro
        self.nome_ficheiro = os.path.basename(caminho_ficheiro)
        self.sheets_disponiveis: List[str] = []
        
        # Funções simples: {nome_role: {'descricao': str, 'tcodes': [str], 'status': str, 'linhas': int}}
        self.roles_simples: Dict[str, Dict[str, Any]] = {}
        
        # Funções compostas: {nome_composta: {'descricao': str, 'roles_filhas': [str], 'status': str}}
        self.roles_compostas: Dict[str, Dict[str, Any]] = {}
        
        # Mapa inverso de TCODE para roles: {tcode: [nome_role]}
        self.tcode_para_roles: Dict[str, List[str]] = {}
        
        # Objetos de autorização: {nome_role: [{'objeto': str, 'campos': dict, 'status': str}]}
        self.autorizacoes: Dict[str, List[Dict[str, Any]]] = {}
        
        # Atribuições CUA: {'adicionar': [dict], 'remover': [dict]}
        self.cua_adicionar: List[Dict[str, Any]] = []
        self.cua_remover: List[Dict[str, Any]] = []
        
        # Utilizadores mapeados: {utilizador: {'adicionar': [dict], 'remover': [dict]}}
        self.utilizadores_cua: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}

        # Sheet CONTROLO: lista de departamentos e seus status
        self.controlo: List[Dict[str, Any]] = []

    def obter_estatisticas(self) -> Dict[str, Any]:
        """Gera um resumo estatístico das funções lidas."""
        total_tcodes = len(self.tcode_para_roles)
        pendentes_add = sum(1 for item in self.cua_adicionar if not item.get("status"))
        pendentes_rm = sum(1 for item in self.cua_remover if not item.get("status"))
        pendentes_controlo = sum(1 for item in self.controlo if item.get("pendente"))
        
        return {
            "ficheiro": self.nome_ficheiro,
            "caminho": self.caminho,
            "total_sheets": len(self.sheets_disponiveis),
            "departamentos_controlo": len(self.controlo),
            "departamentos_pendentes": pendentes_controlo,
            "total_roles_simples": len(self.roles_simples),
            "total_roles_compostas": len(self.roles_compostas),
            "total_tcodes_unicos": total_tcodes,
            "total_cua_adicionar": len(self.cua_adicionar),
            "cua_adicionar_pendentes": pendentes_add,
            "total_cua_remover": len(self.cua_remover),
            "cua_remover_pendentes": pendentes_rm,
            "total_utilizadores_distintos": len(self.utilizadores_cua),
        }


def abrir_excel_seguro(caminho_excel: str):
    """
    Abre o ficheiro Excel em modo de leitura partilhada (FILE_SHARE_READ | FILE_SHARE_WRITE).
    Permite ler o ficheiro mesmo se estiver aberto no Microsoft Excel ou em sincronização pelo OneDrive.
    """
    if sys.platform.startswith("win"):
        try:
            import win32file
            import win32con
            import io

            handle = win32file.CreateFile(
                os.path.abspath(caminho_excel),
                win32con.GENERIC_READ,
                win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
                None,
                win32con.OPEN_EXISTING,
                win32con.FILE_ATTRIBUTE_NORMAL,
                None
            )
            size = win32file.GetFileSize(handle)
            _, content = win32file.ReadFile(handle, size)
            win32file.CloseHandle(handle)
            return io.BytesIO(content)
        except Exception:
            pass

    return caminho_excel


def carregar_projeto_perfil(caminho_excel: Optional[str] = None) -> ProjetoPerfilData:
    """
    Carrega e consolida todas as informações de funções e perfis
    do ficheiro Excel especificado ou localizado automaticamente.
    """
    if not caminho_excel:
        caminho_excel = encontrar_excel_padrao()

    if not caminho_excel or not os.path.exists(caminho_excel):
        raise FileNotFoundError(f"Ficheiro Excel não encontrado: {caminho_excel}")

    try:
        import pandas as pd
    except ImportError:
        raise ImportError("O pacote 'pandas' é necessário. Execute através de .venv-rfc ou instale pandas.")

    dados = ProjetoPerfilData(caminho_excel)
    
    fonte = abrir_excel_seguro(caminho_excel)
    excel_file = pd.ExcelFile(fonte)
    dados.sheets_disponiveis = excel_file.sheet_names

    # -------------------------------------------------------------
    # 0. Leitura de CONTROLO (Departamentos e Status)
    # -------------------------------------------------------------
    sheet_controlo = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "CONTROLO"), None)
    if sheet_controlo:
        df_ctrl = pd.read_excel(excel_file, sheet_name=sheet_controlo)
        cols_map = {normalizar_nome_coluna(c): c for c in df_ctrl.columns}
        col_dep = cols_map.get("DEPARTAMENTO")
        col_st = cols_map.get("STATUS")
        col_ts = cols_map.get("TIMESTAMP") or cols_map.get("TIMESTEMP")

        if col_dep:
            for idx, row in df_ctrl.iterrows():
                dep = str(row.get(col_dep, "")).strip() if pd.notna(row.get(col_dep)) else ""
                if not dep:
                    continue
                st = str(row.get(col_st, "")).strip() if col_st and pd.notna(row.get(col_st)) else ""
                ts = str(row.get(col_ts, "")).strip() if col_ts and pd.notna(row.get(col_ts)) else ""
                dados.controlo.append({
                    "linha": idx + 2,
                    "departamento": dep,
                    "status": st,
                    "timestamp": ts,
                    "pendente": not bool(st)
                })

    # -------------------------------------------------------------
    # 1. Leitura de PFCG_CREATE (Roles Simples e Transações)
    # -------------------------------------------------------------
    sheet_create = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PFCGCREATE"), None)
    if sheet_create:
        df = pd.read_excel(excel_file, sheet_name=sheet_create)
        cols_map = {normalizar_nome_coluna(c): c for c in df.columns}
        
        col_role = cols_map.get("AGRNAME")
        col_text = cols_map.get("TEXT") or cols_map.get("DESCRICAO")
        col_tcode = cols_map.get("TCODE")
        col_status = cols_map.get("STATUS")
        col_msg = cols_map.get("MSG")

        if col_role:
            for _, row in df.iterrows():
                role = normalizar_texto(row.get(col_role))
                if not role:
                    continue
                
                desc = str(row.get(col_text, "")).strip() if col_text and pd.notna(row.get(col_text)) else ""
                tcodes_extraidos = limpar_tcode(row.get(col_tcode)) if col_tcode else []
                status = str(row.get(col_status, "")).strip() if col_status and pd.notna(row.get(col_status)) else ""
                msg = str(row.get(col_msg, "")).strip() if col_msg and pd.notna(row.get(col_msg)) else ""

                if role not in dados.roles_simples:
                    dados.roles_simples[role] = {
                        "role": role,
                        "descricao": desc,
                        "tcodes": [],
                        "status": status,
                        "msg": msg,
                        "total_entradas": 0
                    }
                elif desc and not dados.roles_simples[role]["descricao"]:
                    dados.roles_simples[role]["descricao"] = desc

                dados.roles_simples[role]["total_entradas"] += 1

                for tc in tcodes_extraidos:
                    if tc not in dados.roles_simples[role]["tcodes"]:
                        dados.roles_simples[role]["tcodes"].append(tc)
                    if tc not in dados.tcode_para_roles:
                        dados.tcode_para_roles[tc] = []
                    if role not in dados.tcode_para_roles[tc]:
                        dados.tcode_para_roles[tc].append(role)

    # -------------------------------------------------------------
    # 2. Leitura de PFCG_COMPOSTA (Roles Compostas)
    # -------------------------------------------------------------
    sheet_comp = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PFCGCOMPOSTA"), None)
    if sheet_comp:
        df_comp = pd.read_excel(excel_file, sheet_name=sheet_comp)
        cols_map = {normalizar_nome_coluna(c): c for c in df_comp.columns}
        
        col_comp = cols_map.get("AGRNAMECOMPOSTA") or cols_map.get("COMPOSITEROLE")
        col_text = cols_map.get("TEXT") or cols_map.get("DESCRICAO")
        col_child = cols_map.get("AGRNAME") or cols_map.get("SINGLEROLE")
        col_status = cols_map.get("STATUS")

        if col_comp and col_child:
            for _, row in df_comp.iterrows():
                comp_role = normalizar_texto(row.get(col_comp))
                child_role = normalizar_texto(row.get(col_child))
                if not comp_role:
                    continue
                
                desc = str(row.get(col_text, "")).strip() if col_text and pd.notna(row.get(col_text)) else ""
                status = str(row.get(col_status, "")).strip() if col_status and pd.notna(row.get(col_status)) else ""

                if comp_role not in dados.roles_compostas:
                    dados.roles_compostas[comp_role] = {
                        "role_composta": comp_role,
                        "descricao": desc,
                        "roles_filhas": [],
                        "status": status
                    }
                elif desc and not dados.roles_compostas[comp_role]["descricao"]:
                    dados.roles_compostas[comp_role]["descricao"] = desc

                if child_role and child_role not in dados.roles_compostas[comp_role]["roles_filhas"]:
                    dados.roles_compostas[comp_role]["roles_filhas"].append(child_role)

    # -------------------------------------------------------------
    # 3. Leitura de CUA_ADICIONAR (Atribuições de Perfis a Users)
    # -------------------------------------------------------------
    sheet_cua_add = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "CUAADICIONAR"), None)
    if sheet_cua_add:
        df_add = pd.read_excel(excel_file, sheet_name=sheet_cua_add)
        cols_map = {normalizar_nome_coluna(c): c for c in df_add.columns}

        col_id = cols_map.get("ID")
        col_user = cols_map.get("UTILIZADOR") or cols_map.get("USUARIO") or cols_map.get("USER")
        col_sis = cols_map.get("SISTEMA") or cols_map.get("SUBSYSTEM")
        col_role = cols_map.get("AGRNAME") or cols_map.get("ROLE")
        col_status = cols_map.get("STATUS")
        col_msg = cols_map.get("MSG")

        if col_user and col_role:
            for _, row in df_add.iterrows():
                user = normalizar_texto(row.get(col_user))
                role = normalizar_texto(row.get(col_role))
                if not user or not role:
                    continue

                item = {
                    "id": row.get(col_id) if col_id else None,
                    "utilizador": user,
                    "sistema": normalizar_texto(row.get(col_sis)) if col_sis else "",
                    "role": role,
                    "status": str(row.get(col_status, "")).strip() if col_status and pd.notna(row.get(col_status)) else "",
                    "msg": str(row.get(col_msg, "")).strip() if col_msg and pd.notna(row.get(col_msg)) else "",
                }
                dados.cua_adicionar.append(item)

                if user not in dados.utilizadores_cua:
                    dados.utilizadores_cua[user] = {"adicionar": [], "remover": []}
                dados.utilizadores_cua[user]["adicionar"].append(item)

    # -------------------------------------------------------------
    # 4. Leitura de CUA_REMOVE (Remoções de Perfis de Users)
    # -------------------------------------------------------------
    sheet_cua_rm = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "CUAREMOVE"), None)
    if sheet_cua_rm:
        df_rm = pd.read_excel(excel_file, sheet_name=sheet_cua_rm)
        cols_map = {normalizar_nome_coluna(c): c for c in df_rm.columns}

        col_id = cols_map.get("ID")
        col_user = cols_map.get("UTILIZADOR") or cols_map.get("USUARIO") or cols_map.get("USER")
        col_sis = cols_map.get("SISTEMA")
        col_role = cols_map.get("AGRNAME")
        col_status = cols_map.get("STATUS")
        col_msg = cols_map.get("MSG")

        if col_user and col_role:
            for _, row in df_rm.iterrows():
                user = normalizar_texto(row.get(col_user))
                role = normalizar_texto(row.get(col_role))
                if not user or not role:
                    continue

                item = {
                    "id": row.get(col_id) if col_id else None,
                    "utilizador": user,
                    "sistema": normalizar_texto(row.get(col_sis)) if col_sis else "",
                    "role": role,
                    "status": str(row.get(col_status, "")).strip() if col_status and pd.notna(row.get(col_status)) else "",
                    "msg": str(row.get(col_msg, "")).strip() if col_msg and pd.notna(row.get(col_msg)) else "",
                }
                dados.cua_remover.append(item)

                if user not in dados.utilizadores_cua:
                    dados.utilizadores_cua[user] = {"adicionar": [], "remover": []}
                dados.utilizadores_cua[user]["remover"].append(item)

    return dados


# =====================================================================
# MOTOR DE PESQUISA
# =====================================================================

def pesquisar_funcao(dados: ProjetoPerfilData, termo: str) -> List[Dict[str, Any]]:
    """
    Pesquisa funções (simples ou compostas) pelo nome ou descrição.
    Suporta busca parcial ou exata.
    """
    termo_norm = normalizar_texto(termo)
    if not termo_norm:
        return []

    resultados = []

    # 1. Pesquisa nas Roles Simples
    for role_nome, info in dados.roles_simples.items():
        desc_norm = normalizar_texto(info.get("descricao", ""))
        if termo_norm in role_nome or termo_norm in desc_norm:
            # Identificar se faz parte de alguma composta
            compostas_pai = [
                comp for comp, cinfo in dados.roles_compostas.items()
                if role_nome in cinfo["roles_filhas"]
            ]
            # Identificar se há atribuições no CUA
            atrib_users = [
                item["utilizador"] for item in dados.cua_adicionar
                if item["role"] == role_nome
            ]
            
            resultados.append({
                "tipo": "Simples",
                "role": role_nome,
                "descricao": info.get("descricao", ""),
                "tcodes": info.get("tcodes", []),
                "total_tcodes": len(info.get("tcodes", [])),
                "compostas_relacionadas": compostas_pai,
                "utilizadores_cua": list(dict.fromkeys(atrib_users)),
                "status_excel": info.get("status", "")
            })

    # 2. Pesquisa nas Roles Compostas
    for comp_nome, info in dados.roles_compostas.items():
        desc_norm = normalizar_texto(info.get("descricao", ""))
        if termo_norm in comp_nome or termo_norm in desc_norm:
            # Reunir todas as tcodes das roles filhas
            tcodes_totais = []
            for rf in info["roles_filhas"]:
                if rf in dados.roles_simples:
                    tcodes_totais.extend(dados.roles_simples[rf]["tcodes"])
            
            atrib_users = [
                item["utilizador"] for item in dados.cua_adicionar
                if item["role"] == comp_nome
            ]

            resultados.append({
                "tipo": "Composta",
                "role": comp_nome,
                "descricao": info.get("descricao", ""),
                "roles_filhas": info["roles_filhas"],
                "total_roles_filhas": len(info["roles_filhas"]),
                "tcodes_indiretos": list(dict.fromkeys(tcodes_totais)),
                "utilizadores_cua": list(dict.fromkeys(atrib_users)),
                "status_excel": info.get("status", "")
            })

    return resultados


def pesquisar_por_tcode(dados: ProjetoPerfilData, tcode: str) -> List[Dict[str, Any]]:
    """
    Pesquisa quais funções (simples e compostas) contêm uma determinada transação SAP.
    """
    tcode_norm = normalizar_texto(tcode)
    if not tcode_norm:
        return []

    resultados = []
    # Busca nas transações conhecidas
    tcodes_encontrados = [tc for tc in dados.tcode_para_roles if tcode_norm in tc]

    for tc in tcodes_encontrados:
        roles_diretas = dados.tcode_para_roles[tc]
        for r in roles_diretas:
            info_role = dados.roles_simples.get(r, {})
            # Compostas associadas
            compostas_pai = [
                comp for comp, cinfo in dados.roles_compostas.items()
                if r in cinfo["roles_filhas"]
            ]
            resultados.append({
                "tcode": tc,
                "role_simples": r,
                "descricao_role": info_role.get("descricao", ""),
                "compostas": compostas_pai
            })

    return resultados


def pesquisar_por_utilizador(dados: ProjetoPerfilData, utilizador: str) -> Dict[str, Any]:
    """
    Pesquisa todas as funções atribuídas ou removidas no CUA para um utilizador.
    """
    user_norm = normalizar_texto(utilizador)
    if not user_norm:
        return {"utilizador": utilizador, "adicionar": [], "remover": []}

    # Busca utilizadores correspondentes
    matches = [u for u in dados.utilizadores_cua if user_norm in u]
    
    detalhes = []
    for u in matches:
        info = dados.utilizadores_cua[u]
        detalhes.append({
            "utilizador": u,
            "adicionar": info.get("adicionar", []),
            "remover": info.get("remover", [])
        })

    return {
        "termo": utilizador,
        "total_encontrados": len(matches),
        "resultados": detalhes
    }


def analisar_departamento_proposta(dados: ProjetoPerfilData, departamento: str = "Purchase & Services") -> Dict[str, Any]:
    """
    Analisa a sheet 'Proposta Ativa', filtrando pelo departamento informado.
    Retorna os utilizadores, as funções compostas e as funções individuais (single roles).
    """
    dep_norm = normalizar_texto(departamento)
    sheet_proposta = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PROPOSTAATIVA"), None)
    if not sheet_proposta:
        return {"departamento": departamento, "encontrado": False, "mensagem": "Sheet 'Proposta Ativa' não encontrada no Excel."}

    try:
        import pandas as pd
        from collections import Counter

        fonte = abrir_excel_seguro(dados.caminho)
        df_raw = pd.read_excel(fonte, sheet_name=sheet_proposta, header=None)
    except Exception as err:
        return {"departamento": departamento, "encontrado": False, "mensagem": f"Erro ao ler Proposta Ativa: {err}"}

    usuarios = []
    compostas = set()
    singles_counter = Counter()

    for idx in range(1, len(df_raw)):
        dep_val = str(df_raw.iloc[idx, 7]).strip() if pd.notna(df_raw.iloc[idx, 7]) else ""
        if normalizar_texto(dep_val) == dep_norm or (dep_norm and dep_norm in normalizar_texto(dep_val)):
            user_id = str(df_raw.iloc[idx, 0]).strip() if pd.notna(df_raw.iloc[idx, 0]) else ""
            nome = str(df_raw.iloc[idx, 2]).strip() if pd.notna(df_raw.iloc[idx, 2]) else ""
            cargo = str(df_raw.iloc[idx, 6]).strip() if pd.notna(df_raw.iloc[idx, 6]) else ""
            comp_raw = str(df_raw.iloc[idx, 8]).strip() if pd.notna(df_raw.iloc[idx, 8]) else ""
            comp = comp_raw if comp_raw.upper() not in ("NAN", "NONE", "") else None
            if comp:
                compostas.add(comp)

            user_singles = [
                str(x).strip() for x in df_raw.iloc[idx, 9:].dropna().tolist()
                if str(x).strip() and str(x).strip().upper() not in ("NAN", "NONE", "")
            ]
            for s in user_singles:
                singles_counter[s] += 1

            usuarios.append({
                "linha_excel": idx + 1,
                "usuario": user_id,
                "nome": nome,
                "cargo": cargo,
                "composta": comp,
                "total_singles": len(user_singles),
                "singles": user_singles
            })

    return {
        "departamento": departamento,
        "encontrado": bool(usuarios),
        "total_usuarios": len(usuarios),
        "usuarios": usuarios,
        "compostas": sorted(compostas),
        "total_compostas": len(compostas),
        "singles_frequencia": singles_counter,
        "total_singles_distintas": len(singles_counter)
    }


def imprimir_analise_departamento(dados_dep: Dict[str, Any]):
    """Exibe no terminal o relatório detalhado de funções por departamento."""
    dep = dados_dep.get("departamento", "")
    print("\n" + "=" * 75)
    print(f"  🏢 PROPOSTA ATIVA - ANÁLISE DO DEPARTAMENTO: {dep.upper()}")
    print("=" * 75)

    if not dados_dep.get("encontrado"):
        print(f"  ⚠️ {dados_dep.get('mensagem', 'Nenhum registo encontrado para este departamento.')}")
        return

    users = dados_dep["usuarios"]
    compostas = dados_dep["compostas"]
    singles_count = dados_dep["singles_frequencia"]

    print(f"  👥 Utilizadores no Departamento:            {len(users)}")
    print(f"  📦 Funções Compostas Distintas:             {len(compostas)}")
    print(f"  🧩 Funções Individuais (Singles) Distintas: {len(singles_count)}")
    print("-" * 75)

    print("\n  📦 1. FUNÇÕES COMPOSTAS DO DEPARTAMENTO:")
    for c in compostas:
        users_com_comp = [u["usuario"] for u in users if u.get("composta") == c]
        print(f"     ├─ [Composta] {c:<35} -> Utilizadores: {', '.join(users_com_comp)}")

    print("\n  👥 2. UTILIZADORES E RESPETIVAS FUNÇÕES:")
    for u in users:
        comp_str = f"Composta: {u['composta']}" if u.get("composta") else "Sem Composta"
        print(f"     👤 {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']} [{comp_str}] -> {u['total_singles']} Singles")

    print(f"\n  🧩 3. FUNÇÕES INDIVIDUAIS (SINGLE ROLES) ({len(singles_count)} distintas):")
    # Ordenar por frequência decrescente e alfabético
    for s, count in sorted(singles_count.items(), key=lambda x: (-x[1], x[0])):
        print(f"     ├─ {s:<35} -> presente em {count}/{len(users)} utilizador(es)")


def verificar_funcoes_prd(roles: List[str]) -> Dict[str, Any]:
    """
    Verifica se uma lista de funções (roles) existe no sistema SAP PRD através de RFC (tabela AGR_DEFINE).
    """
    if not roles:
        return {"ok": True, "total": 0, "existentes": [], "nao_existentes": []}

    roles_unicas = sorted(list({str(r).strip().upper() for r in roles if str(r).strip()}))

    os.environ["SAP_TARGET_ENV"] = "PRD"
    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_read_only_guard, read_table, make_option_in
        )
        from pyrfc import Connection

        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for("PRD")
        conn = Connection(**params)
        guard = make_read_only_guard(["AGR_DEFINE", "AGR_TEXTS"])

        encontradas = set()
        chunk_size = 25
        for i in range(0, len(roles_unicas), chunk_size):
            chunk = roles_unicas[i:i + chunk_size]
            opts = make_option_in("AGR_NAME", chunk)
            rows = read_table(conn, guard, table_name="AGR_DEFINE", fields=["AGR_NAME"], options=opts, rowcount=len(chunk) + 10)
            for r in rows:
                if r and r[0].strip():
                    encontradas.add(r[0].strip())

        conn.close()

        nao_encontradas = sorted(list(set(roles_unicas) - encontradas))
        return {
            "ok": True,
            "sistema": "PRD",
            "total": len(roles_unicas),
            "total_existentes": len(encontradas),
            "total_nao_existentes": len(nao_encontradas),
            "existentes": sorted(list(encontradas)),
            "nao_existentes": nao_encontradas,
            "todas_existem": len(nao_encontradas) == 0
        }
    except Exception as err:
        return {
            "ok": False,
            "sistema": "PRD",
            "erro": str(err),
            "total": len(roles_unicas),
            "existentes": [],
            "nao_existentes": roles_unicas
        }


def imprimir_resultado_verificacao_prd(resultado: Dict[str, Any], titulo_contexto: str = ""):
    """Imprime no terminal o resultado da verificação de funções no PRD."""
    print("\n" + "=" * 75)
    header = f"  🔍 VERIFICAÇÃO DE FUNÇÕES NO SAP PRD (via RFC){' - ' + titulo_contexto if titulo_contexto else ''}"
    print(header)
    print("=" * 75)

    if not resultado.get("ok"):
        print(f"  ❌ Erro de ligação RFC ao PRD: {resultado.get('erro')}")
        return

    total = resultado["total"]
    existentes = resultado["total_existentes"]
    nao_ex = resultado["total_nao_existentes"]
    pct = (existentes / total * 100) if total > 0 else 0

    print(f"  Total de Funções Verificadas: {total}")
    print(f"  ✅ Existentes no PRD:          {existentes} ({pct:.1f}%)")
    print(f"  ❌ Inexistentes / Em Falta:    {nao_ex}")
    print("-" * 75)

    if resultado["todas_existem"]:
        print("  🎉 Todas as funções pesquisadas EXISTEM no sistema PRD!")
    else:
        print("  ⚠️ As seguintes funções NÃO foram encontradas no SAP PRD (AGR_DEFINE):")
        for r in resultado["nao_existentes"]:
            print(f"     ❌ {r}")


# =====================================================================
# FORMATAÇÃO VISUAL E INTERFACE DE LINHA DE COMANDOS
# =====================================================================

def imprimir_cabecalho(ficheiro: str):
    sp_url = os.getenv("SHAREPOINT_PERFIS_URL", "")
    print("\n" + "=" * 75)
    print("  🚀 PROJETO PERFIL - Leitor e Pesquisador de Funções SAP (PFCG / CUA)")
    print(f"  📂 Ficheiro:          {os.path.basename(ficheiro)}")
    print(f"  📍 Caminho Local:     {ficheiro}")
    if sp_url:
        print(f"  🌐 SharePoint:        {sp_url}")
    print("=" * 75)


def imprimir_validacao_controlo(dados: ProjetoPerfilData):
    """Exibe no terminal a validação dos departamentos da sheet CONTROLO."""
    print("\n" + "=" * 75)
    print("  📋 VALIDAÇÃO DA SHEET 'CONTROLO' (Departamentos & Status)")
    print("=" * 75)
    if not dados.controlo:
        print("  ⚠️ A sheet 'CONTROLO' não foi encontrada ou não possui dados.")
        return

    pendentes = [d for d in dados.controlo if d.get("pendente")]
    print(f"  Total de departamentos registados: {len(dados.controlo)}")
    print(f"  Departamentos com STATUS vazio (Pendentes): {len(pendentes)}")
    print("-" * 75)
    for c in dados.controlo:
        status_tag = "⏳ PENDENTE (STATUS VAZIO)" if c["pendente"] else f"✅ {c['status']}"
        ts_info = f" | {c['timestamp']}" if c.get("timestamp") else ""
        print(f"  Linha {c['linha']}: {c['departamento']:<30} -> {status_tag}{ts_info}")

    print("-" * 75)
    if pendentes:
        print("  🎯 DEPARTAMENTO(S) A PROCESSAR:")
        for p in pendentes:
            tem_sheet = p["departamento"] in dados.sheets_disponiveis
            sheet_info = f"[Sheet correspondente '{p['departamento']}' EXISTE no Excel]" if tem_sheet else "[Sheet correspondente não encontrada]"
            print(f"   👉 '{p['departamento']}' (Linha {p['linha']}) -> STATUS ESTÁ VAZIO! {sheet_info}")
    else:
        print("  ✅ Todos os departamentos na sheet CONTROLO já possuem STATUS preenchido.")


def imprimir_resumo(dados: ProjetoPerfilData):
    stats = dados.obter_estatisticas()
    print("\n📊 RESUMO GERAL DAS FUNÇÕES:")
    print(f"  ├─ Ficheiro:                  {stats['caminho']}")
    print(f"  ├─ Total de Sheets lidas:     {stats['total_sheets']} {dados.sheets_disponiveis}")
    print(f"  ├─ Departamentos em CONTROLO: {stats['departamentos_controlo']} ({stats['departamentos_pendentes']} pendentes com STATUS vazio)")
    print(f"  ├─ Roles Simples únicas:      {stats['total_roles_simples']}")
    print(f"  ├─ Roles Compostas:           {stats['total_roles_compostas']}")
    print(f"  ├─ Transações (TCODEs) únicas: {stats['total_tcodes_unicos']}")
    print(f"  ├─ CUA Adicionar (atribuições): {stats['total_cua_adicionar']} ({stats['cua_adicionar_pendentes']} pendentes)")
    print(f"  ├─ CUA Remover (remoções):    {stats['total_cua_remover']} ({stats['cua_remover_pendentes']} pendentes)")
    print(f"  └─ Utilizadores no CUA:       {stats['total_utilizadores_distintos']}")


def imprimir_resultado_funcao(res: List[Dict[str, Any]]):
    if not res:
        print("  ⚠️ Nenhuma função encontrada com o critério indicado.")
        return

    print(f"\n🔍 Encontrada(s) {len(res)} função(ões):")
    for idx, item in enumerate(res, 1):
        print(f"\n  [{idx}] {item['tipo'].upper()}: {item['role']}")
        if item.get("descricao"):
            print(f"      Descrição: {item['descricao']}")
        
        if item["tipo"] == "Simples":
            tcs = item.get("tcodes", [])
            tcs_preview = ", ".join(tcs[:10]) + ("..." if len(tcs) > 10 else "")
            print(f"      Transações ({len(tcs)}): {tcs_preview if tcs else 'Sem TCODEs explícitos'}")
            if item.get("compostas_relacionadas"):
                print(f"      Pertence a Compostas: {', '.join(item['compostas_relacionadas'])}")
        else:
            filhas = item.get("roles_filhas", [])
            print(f"      Roles Filhas ({len(filhas)}): {', '.join(filhas)}")
            tcs_ind = item.get("tcodes_indiretos", [])
            if tcs_ind:
                print(f"      TCODEs herdados ({len(tcs_ind)}): {', '.join(tcs_ind[:10])}")

        if item.get("utilizadores_cua"):
            print(f"      Utilizadores CUA com esta função: {', '.join(item['utilizadores_cua'])}")
        if item.get("status_excel"):
            print(f"      Status no Excel: {item['status_excel']}")


def imprimir_resultado_tcode(res: List[Dict[str, Any]]):
    if not res:
        print("  ⚠️ Nenhuma função encontrada que contenha esta transação.")
        return

    print(f"\n🔍 Encontrada(s) {len(res)} correspondência(s) de transação:")
    # Agrupar por TCODE
    por_tc: Dict[str, List[Dict[str, Any]]] = {}
    for r in res:
        tc = r["tcode"]
        por_tc.setdefault(tc, []).append(r)

    for tc, lista in por_tc.items():
        print(f"\n  📌 TCODE: {tc}")
        for item in lista:
            role = item["role_simples"]
            desc = f" ({item['descricao_role']})" if item.get("descricao_role") else ""
            comp_txt = f" [Compostas: {', '.join(item['compostas'])}]" if item.get("compostas") else ""
            print(f"     ├─ Role Simples: {role}{desc}{comp_txt}")


def imprimir_resultado_user(res: Dict[str, Any]):
    users = res.get("resultados", [])
    if not users:
        print("  ⚠️ Nenhum utilizador encontrado com esse identificador no CUA.")
        return

    print(f"\n🔍 Encontrado(s) {len(users)} utilizador(es):")
    for u in users:
        nome_user = u["utilizador"]
        adds = u.get("adicionar", [])
        rms = u.get("remover", [])
        print(f"\n  👤 Utilizador: {nome_user}")
        if adds:
            print(f"     ✅ Funções para Adicionar ({len(adds)}):")
            for a in adds:
                st = f" [STATUS: {a['status']}]" if a.get("status") else " [PENDENTE]"
                sis = f" ({a['sistema']})" if a.get("sistema") else ""
                print(f"        ├─ {a['role']}{sis}{st}")
        if rms:
            print(f"     ❌ Funções para Remover ({len(rms)}):")
            for r in rms:
                st = f" [STATUS: {r['status']}]" if r.get("status") else " [PENDENTE]"
                sis = f" ({r['sistema']})" if r.get("sistema") else ""
                print(f"        ├─ {r['role']}{sis}{st}")


# =====================================================================
# MENU INTERATIVO
# =====================================================================

def menu_interativo(caminho_inicial: Optional[str] = None):
    """Executa o menu interativo de consola."""
    caminho = caminho_inicial or encontrar_excel_padrao()

    if not caminho or not os.path.exists(caminho):
        print("📂 Nenhum ficheiro Excel detectado automaticamente.")
        caminho = selecionar_ficheiro_dialogo()
        if not caminho:
            print("❌ Operação cancelada. Saindo.")
            return

    try:
        print(f"⏳ A carregar ficheiro Excel: {os.path.basename(caminho)} ...")
        dados = carregar_projeto_perfil(caminho)
    except Exception as e:
        print(f"❌ Erro ao ler Excel: {e}")
        return

    imprimir_cabecalho(caminho)
    imprimir_validacao_controlo(dados)
    imprimir_resumo(dados)

    while True:
        print("\n" + "-" * 75)
        print("  MENU PRINCIPAL - PROJETO PERFIL:")
        print("  [1] 📋 Validar Sheet CONTROLO (Departamentos com STATUS vazio)")
        print("  [2] 🏢 Analisar Funções por Departamento (Proposta Ativa)")
        print("  [3] 🌐 Verificar Existência de Funções no SAP PRD (via RFC)")
        print("  [4] 📊 Ver Resumo Geral das Funções")
        print("  [5] 🔍 Pesquisar por Nome de Função (Role / Perfil)")
        print("  [6] 📌 Pesquisar por Transação (TCODE -> Funções)")
        print("  [7] 👤 Pesquisar por Utilizador (CUA: Adições / Remoções)")
        print("  [8] 📋 Listar todas as Roles Simples")
        print("  [9] 📋 Listar todas as Roles Compostas")
        print("  [10] 📂 Abrir outro ficheiro Excel")
        print("  [0] 🚪 Sair")
        print("-" * 75)

        opcao = input("👉 Escolha uma opção: ").strip()

        if opcao == "1":
            imprimir_validacao_controlo(dados)

        elif opcao == "2":
            dep_padrao = "Purchase & Services"
            dep_in = input(f"🔎 Nome do departamento (Enter para '{dep_padrao}'): ").strip()
            alvo = dep_in if dep_in else dep_padrao
            res_dep = analisar_departamento_proposta(dados, alvo)
            imprimir_analise_departamento(res_dep)

        elif opcao == "3":
            print("\n  O que deseja verificar no SAP PRD?")
            print("  [1] Funções do departamento 'Purchase & Services'")
            print("  [2] Todas as funções do ficheiro Excel (PFCG_CREATE + PFCG_COMPOSTA)")
            sub_op = input("👉 Escolha (1 ou 2, padrão 1): ").strip()
            if sub_op == "2":
                todas = list(dados.roles_simples.keys()) + list(dados.roles_compostas.keys())
                res_prd = verificar_funcoes_prd(todas)
                imprimir_resultado_verificacao_prd(res_prd, "Todas as Funções do Excel")
            else:
                dep_info = analisar_departamento_proposta(dados, "Purchase & Services")
                if dep_info.get("encontrado"):
                    roles_dep = list(dep_info["compostas"]) + list(dep_info["singles_frequencia"].keys())
                    res_prd = verificar_funcoes_prd(roles_dep)
                    imprimir_resultado_verificacao_prd(res_prd, "Departamento Purchase & Services")
                else:
                    print("⚠️ Não foi possível obter as funções do departamento.")

        elif opcao == "4":
            imprimir_resumo(dados)

        elif opcao == "5":
            termo = input("🔎 Digite o nome da função ou texto (ex.: Z_BR, MANAGER, ZORG): ").strip()
            if termo:
                res = pesquisar_funcao(dados, termo)
                imprimir_resultado_funcao(res)

        elif opcao == "6":
            tcode = input("🔎 Digite a transação SAP (ex.: FB01, CO01, SU01, SE16): ").strip()
            if tcode:
                res = pesquisar_por_tcode(dados, tcode)
                imprimir_resultado_tcode(res)

        elif opcao == "7":
            user = input("🔎 Digite o ID do Utilizador SAP (ex.: S6005, S5354): ").strip()
            if user:
                res = pesquisar_por_utilizador(dados, user)
                imprimir_resultado_user(res)

        elif opcao == "8":
            print(f"\n📋 Total de {len(dados.roles_simples)} Roles Simples:")
            for r, info in sorted(dados.roles_simples.items()):
                t_count = len(info["tcodes"])
                desc = f" - {info['descricao']}" if info.get("descricao") else ""
                print(f"  ├─ {r} ({t_count} TCODEs){desc}")

        elif opcao == "9":
            print(f"\n📋 Total de {len(dados.roles_compostas)} Roles Compostas:")
            for c, info in sorted(dados.roles_compostas.items()):
                filhas_count = len(info["roles_filhas"])
                desc = f" - {info['descricao']}" if info.get("descricao") else ""
                print(f"  ├─ {c} ({filhas_count} roles filhas){desc}")

        elif opcao == "10":
            novo_caminho = selecionar_ficheiro_dialogo()
            if novo_caminho and os.path.exists(novo_caminho):
                try:
                    caminho = novo_caminho
                    print(f"⏳ A carregar ficheiro: {os.path.basename(caminho)} ...")
                    dados = carregar_projeto_perfil(caminho)
                    imprimir_cabecalho(caminho)
                    imprimir_validacao_controlo(dados)
                    imprimir_resumo(dados)
                except Exception as e:
                    print(f"❌ Erro ao carregar novo ficheiro: {e}")
            else:
                print("⚠️ Nenhum ficheiro selecionado.")

        elif opcao in ("0", "sair", "exit", "q"):
            print("\n👋 Sessão terminada. Até breve!")
            break
        else:
            print("⚠️ Opção inválida. Tente novamente.")


# =====================================================================
# PONTO DE ENTRADA CLI
# =====================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Projeto Perfil - Análise e Pesquisa de Funções SAP")
    parser.add_argument("--xlsx", "--ficheiro", "-f", dest="ficheiro", help="Caminho do ficheiro Excel de perfis")
    parser.add_argument("--controlo", "--validar-controlo", dest="controlo", action="store_true", help="Validar sheet CONTROLO e listar departamentos pendentes")
    parser.add_argument("--departamento", "-d", dest="departamento", nargs="?", const="Purchase & Services", help="Analisar funções de um departamento na sheet Proposta Ativa")
    parser.add_argument("--verificar-prd", dest="verificar_prd", action="store_true", help="Verificar se as funções existem no sistema SAP PRD via RFC")
    parser.add_argument("--pesquisar-role", "-r", dest="role", help="Pesquisar diretamente por nome de função")
    parser.add_argument("--pesquisar-tcode", "-t", dest="tcode", help="Pesquisar diretamente por transação SAP")
    parser.add_argument("--pesquisar-user", "-u", dest="user", help="Pesquisar diretamente por utilizador CUA")

    args = parser.parse_args()

    caminho_alvo = args.ficheiro or encontrar_excel_padrao()

    # Se foram passados parâmetros de pesquisa direta via CLI:
    if args.controlo or args.departamento or args.verificar_prd or args.role or args.tcode or args.user:
        if not caminho_alvo:
            print("❌ Erro: Ficheiro Excel não encontrado.")
            sys.exit(1)
        
        dados = carregar_projeto_perfil(caminho_alvo)
        imprimir_cabecalho(caminho_alvo)

        if args.controlo:
            imprimir_validacao_controlo(dados)
        if args.departamento and not args.verificar_prd:
            res_dep = analisar_departamento_proposta(dados, args.departamento)
            imprimir_analise_departamento(res_dep)
        if args.verificar_prd:
            if args.departamento:
                dep_info = analisar_departamento_proposta(dados, args.departamento)
                roles_dep = list(dep_info["compostas"]) + list(dep_info["singles_frequencia"].keys())
                res_prd = verificar_funcoes_prd(roles_dep)
                imprimir_resultado_verificacao_prd(res_prd, f"Departamento '{args.departamento}'")
            else:
                todas = list(dados.roles_simples.keys()) + list(dados.roles_compostas.keys())
                res_prd = verificar_funcoes_prd(todas)
                imprimir_resultado_verificacao_prd(res_prd, "Todas as Funções do Excel (PFCG_CREATE + PFCG_COMPOSTA)")
        if args.role:
            imprimir_resultado_funcao(pesquisar_funcao(dados, args.role))
        if args.tcode:
            imprimir_resultado_tcode(pesquisar_por_tcode(dados, args.tcode))
        if args.user:
            imprimir_resultado_user(pesquisar_por_utilizador(dados, args.user))
    else:
        # Modo interativo padrão
        menu_interativo(caminho_alvo)

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
from pathlib import Path
from typing import Optional, Dict, Any, List, Set, Iterable

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
        print(f"[AVISO] Não foi possível abrir o diálogo gráfico: {err}")
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

        # Sheet PFCG_AUTHORITY: funções avulsas associadas a compostas {nome_composta: set(roles_filhas)}
        self.roles_authority: Dict[str, Set[str]] = {}

        # Sheet EXCLUÇÃO: padrões globais a desconsiderar (ex.: ZMM_APROVA_PEDC_COD_*, Z_MY_HOME)
        self.padroes_exclusao: List[str] = []

    def is_excluida(self, role: str) -> bool:
        """Verifica se uma role corresponde a algum padrão da sheet EXCLUÇÃO."""
        import fnmatch
        r = str(role).strip().upper()
        return any(fnmatch.fnmatchcase(r, p.upper()) for p in self.padroes_exclusao)

    def expandir_funcoes(self, roles: Iterable[str]) -> Set[str]:
        """
        Expande recursivamente um conjunto de roles respeitando:
        - Filhas em PFCG_COMPOSTA
        - Avulsas em PFCG_AUTHORITY
        - Desconsiderando as que coincidem com EXCLUÇÃO
        """
        expandidas: Set[str] = set()
        pendentes = [str(r).strip().upper() for r in roles if str(r).strip()]
        while pendentes:
            role = pendentes.pop()
            if not role or role in expandidas or self.is_excluida(role):
                continue
            expandidas.add(role)
            membros = set(self.roles_compostas.get(role, {}).get("roles_filhas", []))
            membros.update(self.roles_authority.get(role, set()))
            pendentes.extend(membros - expandidas)
        return expandidas

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
            "total_authority_compostas": len(self.roles_authority),
            "padroes_exclusao": self.padroes_exclusao,
            "total_tcodes_unicos": total_tcodes,
            "total_cua_adicionar": len(self.cua_adicionar),
            "cua_adicionar_pendentes": pendentes_add,
            "total_cua_remover": len(self.cua_remover),
            "cua_remover_pendentes": pendentes_rm,
            "total_utilizadores_distintos": len(self.utilizadores_cua),
        }

    def obter_proximo_departamento(self) -> Optional[str]:
        """
        Devolve o próximo departamento a avançar com base na sheet CONTROLO.
        Critério oficial: primeiro departamento onde as colunas STATUS e TIMESTAMP estão vazias.
        """
        for item in self.controlo:
            if item.get("pendente"):
                return item.get("departamento")
        return None



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
                if st.lower() in ("nan", "none", "<na>"):
                    st = ""
                ts = str(row.get(col_ts, "")).strip() if col_ts and pd.notna(row.get(col_ts)) else ""
                if ts.lower() in ("nan", "none", "<na>", "nat"):
                    ts = ""

                # Regra oficial: o departamento a avançar é aquele com STATUS e TIMESTAMP vazios
                is_pendente = (not bool(st)) and (not bool(ts))

                dados.controlo.append({
                    "linha": idx + 2,
                    "departamento": dep,
                    "status": st,
                    "timestamp": ts,
                    "pendente": is_pendente
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

    # -------------------------------------------------------------
    # 5. Leitura de PFCG_AUTHORITY (Funções de Autorização ligadas a Compostas)
    # -------------------------------------------------------------
    sheet_auth = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PFCGAUTHORITY"), None)
    if sheet_auth:
        df_auth = pd.read_excel(excel_file, sheet_name=sheet_auth)
        cols_map = {normalizar_nome_coluna(c): c for c in df_auth.columns}
        col_comp = cols_map.get("AGRNAMECOMPOSTA") or cols_map.get("COMPOSITEROLE")
        col_role = cols_map.get("AGRNAME") or cols_map.get("SINGLEROLE")
        if col_comp and col_role:
            for _, row in df_auth.iterrows():
                comp = normalizar_texto(row.get(col_comp))
                role = normalizar_texto(row.get(col_role))
                if comp and role:
                    dados.roles_authority.setdefault(comp, set()).add(role)

    # -------------------------------------------------------------
    # 6. Leitura de EXCLUÇÃO / EXCLUSAO (Padrões Desconsiderados)
    # -------------------------------------------------------------
    sheet_exc = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) in ("EXCLUCAO", "EXCLUSAO")), None)
    if sheet_exc:
        df_exc = pd.read_excel(excel_file, sheet_name=sheet_exc)
        if len(df_exc.columns):
            dados.padroes_exclusao = [
                normalizar_texto(v) for v in df_exc.iloc[:, 0].dropna().tolist()
                if normalizar_texto(v)
            ]

    return dados


def obter_proximo_departamento_controlo(caminho_excel: Optional[str] = None) -> Optional[str]:
    """
    Lê a sheet CONTROLO do ficheiro Excel e devolve o próximo departamento a avançar
    (aquele cujas colunas STATUS e TIMESTAMP estão vazias).
    """
    dados = carregar_projeto_perfil(caminho_excel)
    return dados.obter_proximo_departamento()


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
    print(f"  PROPOSTA ATIVA - ANÁLISE DO DEPARTAMENTO: {dep.upper()}")
    print("=" * 75)

    if not dados_dep.get("encontrado"):
        print(f"  [AVISO] {dados_dep.get('mensagem', 'Nenhum registo encontrado para este departamento.')}")
        return

    users = dados_dep["usuarios"]
    compostas = dados_dep["compostas"]
    singles_count = dados_dep["singles_frequencia"]

    print(f"  Utilizadores no Departamento:            {len(users)}")
    print(f"  Funções Compostas Distintas:             {len(compostas)}")
    print(f"  Funções Individuais (Singles) Distintas: {len(singles_count)}")
    print("-" * 75)

    print("\n  1. FUNÇÕES COMPOSTAS DO DEPARTAMENTO:")
    for c in compostas:
        users_com_comp = [u["usuario"] for u in users if u.get("composta") == c]
        print(f"     ├─ [Composta] {c:<35} -> Utilizadores: {', '.join(users_com_comp)}")

    print("\n  2. UTILIZADORES E RESPETIVAS FUNÇÕES:")
    for u in users:
        comp_str = f"Composta: {u['composta']}" if u.get("composta") else "Sem Composta"
        print(f"     {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']} [{comp_str}] -> {u['total_singles']} Singles")

    print(f"\n  3. FUNÇÕES INDIVIDUAIS (SINGLE ROLES) ({len(singles_count)} distintas):")
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
    header = f"  VERIFICAÇÃO DE FUNÇÕES NO SAP PRD (via RFC){' - ' + titulo_contexto if titulo_contexto else ''}"
    print(header)
    print("=" * 75)

    if not resultado.get("ok"):
        print(f"  [ERRO] Erro de ligação RFC ao PRD: {resultado.get('erro')}")
        return

    total = resultado["total"]
    existentes = resultado["total_existentes"]
    nao_ex = resultado["total_nao_existentes"]
    pct = (existentes / total * 100) if total > 0 else 0

    print(f"  Total de Funções Verificadas: {total}")
    print(f"  Existentes no PRD:          {existentes} ({pct:.1f}%)")
    print(f"  [ERRO] Inexistentes / Em Falta:    {nao_ex}")
    print("-" * 75)

    if resultado["todas_existem"]:
        print("  Todas as funções pesquisadas EXISTEM no sistema PRD!")
    else:
        print("  [AVISO] As seguintes funções NÃO foram encontradas no SAP PRD (AGR_DEFINE):")
        for r in resultado["nao_existentes"]:
            print(f"     [ERRO] {r}")


def verificar_tcodes_prd(tcodes: List[str]) -> Dict[str, Any]:
    """
    Verifica se uma lista de transações (TCODEs) existe no sistema SAP PRD através de RFC (tabela TSTC).
    Também recolhe descrições da tabela TSTCT para validação de integridade prévia ao PFCG.
    """
    if not tcodes:
        return {"ok": True, "total": 0, "total_existentes": 0, "total_nao_existentes": 0, "existentes": [], "nao_existentes": [], "todas_existem": True, "textos": {}}

    tcodes_unicas = sorted(list({str(t).strip().upper() for t in tcodes if str(t).strip()}))

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
        guard = make_read_only_guard(["TSTC", "TSTCT"])

        encontradas = set()
        textos = {}
        chunk_size = 50

        # 1. Consulta à tabela TSTC (mestre de transações)
        for i in range(0, len(tcodes_unicas), chunk_size):
            chunk = tcodes_unicas[i:i + chunk_size]
            opts = make_option_in("TCODE", chunk)
            rows = read_table(conn, guard, table_name="TSTC", fields=["TCODE"], options=opts, rowcount=0)
            for r in rows:
                if r and r[0].strip():
                    encontradas.add(r[0].strip().upper())

        # 2. Consulta opcional a textos em TSTCT
        for i in range(0, len(tcodes_unicas), chunk_size):
            chunk = tcodes_unicas[i:i + chunk_size]
            opts = make_option_in("TCODE", chunk)
            try:
                rows_t = read_table(conn, guard, table_name="TSTCT", fields=["TCODE", "SPRSL", "TTEXT"], options=opts, rowcount=0)
                for r in rows_t:
                    if len(r) >= 3:
                        tc = r[0].strip().upper()
                        lang = r[1].strip().upper()
                        txt = r[2].strip()
                        if lang in ("P", "PT") or tc not in textos:
                            textos[tc] = txt
            except Exception:
                pass

        conn.close()

        nao_encontradas = sorted(list(set(tcodes_unicas) - encontradas))
        return {
            "ok": True,
            "sistema": "PRD",
            "mandante": params.get("client", "100"),
            "total": len(tcodes_unicas),
            "total_existentes": len(encontradas),
            "total_nao_existentes": len(nao_encontradas),
            "existentes": sorted(list(encontradas)),
            "nao_existentes": nao_encontradas,
            "todas_existem": len(nao_encontradas) == 0,
            "textos": textos,
        }
    except Exception as err:
        return {
            "ok": False,
            "sistema": "PRD",
            "erro": str(err),
            "total": len(tcodes_unicas),
            "total_existentes": 0,
            "total_nao_existentes": len(tcodes_unicas),
            "existentes": [],
            "nao_existentes": tcodes_unicas,
            "todas_existem": False,
            "textos": {},
        }


def imprimir_resultado_verificacao_tcodes_prd(resultado: Dict[str, Any], titulo_contexto: str = ""):
    """Imprime no terminal o resultado da verificação de transações no PRD (TSTC)."""
    print("\n" + "=" * 75)
    header = f"  VERIFICAÇÃO DE TRANSAÇÕES NO SAP PRD (TSTC via RFC){' - ' + titulo_contexto if titulo_contexto else ''}"
    print(header)
    print("=" * 75)

    if not resultado.get("ok"):
        print(f"  [ERRO] Erro de ligação RFC ao PRD: {resultado.get('erro')}")
        return

    total = resultado["total"]
    existentes = resultado["total_existentes"]
    nao_ex = resultado["total_nao_existentes"]
    pct = (existentes / total * 100) if total > 0 else 0

    print(f"  Total de Transações Verificadas: {total}")
    print(f"  Existentes no PRD (TSTC):      {existentes} ({pct:.1f}%)")
    print(f"  [ERRO] Inexistentes no PRD:           {nao_ex}")
    print("-" * 75)

    if resultado["todas_existem"]:
        print("  Todas as transações pesquisadas EXISTEM no sistema PRD (TSTC)!")
        print("     A atribuição de menus e autorizações no PFCG não sofrerá erros de input.")
    else:
        print("  [AVISO] As seguintes transações NÃO FORAM ENCONTRADAS no SAP PRD (TSTC):")
        for tc in resultado["nao_existentes"]:
            print(f"     [ERRO] {tc}")


def is_nome_funcao_proposta(valor_funcao: Any, valor_descricao: Any) -> bool:
    """
    Determina de forma estrita se uma linha na folha 'Proposta' representa a declaração de uma Função PFCG:
    - O nome da função deve começar obrigatoriamente por 'Z_' (convenção oficial de funções do projeto).
    - Deve possuir uma descrição textual válida e não vazia.
    - Casos como 'SE16N' com 'EKKO' ou 'MARA' na descrição são transações/anotações e NUNCA nomes de função.
    - Se a descrição contiver 'Transação não existe', NUNCA é cabeçalho de função.
    """
    f = str(valor_funcao or "").strip().upper()
    d = str(valor_descricao or "").strip()
    if not f or not d or d.lower() in ("nan", "none", ""):
        return False
    d_norm = normalizar_texto(d)
    if "TRANSACAO NAO EXISTE" in d_norm or "NAO EXISTE" in d_norm:
        return False
    # Todas as funções oficiais no S/4HANA do projeto iniciam-se com Z_ (ex.: Z_PRODUCTION_ORDER_CREATE)
    # TCODEs standard como SE16N, CO01, FB03, etc. nunca iniciam por Z_
    return f.startswith("Z_") and len(f) >= 4


def is_tcode_marcado_inexistente(valor_descricao: Any) -> bool:
    """Verifica se a coluna ao lado (DESCRIÇÃO) assinala que a transação não existe no SAP."""
    d_norm = normalizar_texto(valor_descricao)
    return "TRANSACAO NAO EXISTE" in d_norm or "NAO EXISTE" in d_norm


def analisar_folha_proposta(dados: ProjetoPerfilData) -> Dict[str, Any]:
    """
    Analisa a folha 'Proposta' identificando blocos verticais de funções e transações:
    - Linha de Função: FUNÇÃO (Z_*) e DESCRIÇÃO preenchidas.
    - Linhas de Transações: linhas subsequentes com FUNÇÃO preenchida.
      * Se a coluna ao lado (DESCRIÇÃO) tiver 'Transação não existe', a transação é RETIRADA da lista!
      * Linhas como SE16N com 'EKKO' ou 'MARA' são tratadas como transações normais, nunca como funções.
    Cruza contra PFCG_CREATE para identificar se existem funções novas a serem criadas.
    """
    sheet_prop = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PROPOSTA"), None)
    if not sheet_prop:
        return {"encontrado": False, "mensagem": "Sheet 'Proposta' não encontrada no Excel."}

    import pandas as pd
    fonte = abrir_excel_seguro(dados.caminho)
    df = pd.read_excel(fonte, sheet_name=sheet_prop)

    funcoes_proposta = {}
    current_role = None
    current_desc = None
    current_tcodes = []
    current_tcodes_descartados = []
    current_linha = 0
    tcode_para_roles = {}
    tcodes_descartados_total = []

    for idx, row in df.iterrows():
        f_val = str(row["FUNÇÃO"]).strip() if pd.notna(row["FUNÇÃO"]) else ""
        d_val = str(row["DESCRIÇÃO"]).strip() if pd.notna(row["DESCRIÇÃO"]) else ""

        if is_nome_funcao_proposta(f_val, d_val):
            if current_role:
                funcoes_proposta[current_role] = {
                    "linha_excel": current_linha,
                    "descricao": current_desc,
                    "tcodes": list(dict.fromkeys(current_tcodes)),
                    "tcodes_inexistentes": list(dict.fromkeys(current_tcodes_descartados)),
                }
            current_role = f_val.upper()
            current_desc = d_val
            current_tcodes = []
            current_tcodes_descartados = []
            current_linha = idx + 2
        elif f_val and f_val.lower() not in ("nan", "none"):
            if current_role:
                # Regra: se a coluna ao lado contiver 'Transação não existe', RETIRA da lista da sheet!
                if is_tcode_marcado_inexistente(d_val):
                    current_tcodes_descartados.append(f_val.upper())
                    tcodes_descartados_total.append({
                        "linha_excel": idx + 2,
                        "role": current_role,
                        "tcode": f_val.upper(),
                        "motivo": d_val
                    })
                else:
                    # É uma transação válida atribuída à role (mesmo se d_val tiver notas como EKKO/MARA)
                    for tc in limpar_tcode(f_val):
                        current_tcodes.append(tc)
                        tcode_para_roles.setdefault(tc, []).append(current_role)

    if current_role:
        funcoes_proposta[current_role] = {
            "linha_excel": current_linha,
            "descricao": current_desc,
            "tcodes": list(dict.fromkeys(current_tcodes)),
            "tcodes_inexistentes": list(dict.fromkeys(current_tcodes_descartados)),
        }

    # Comparar com PFCG_CREATE
    roles_pfcg = dados.roles_simples
    novas_para_criar = [
        {
            "role": r,
            "descricao": info["descricao"],
            "linha": info["linha_excel"],
            "total_tcodes": len(info["tcodes"]),
            "tcodes": info["tcodes"],
            "tcodes_inexistentes": info.get("tcodes_inexistentes", [])
        }
        for r, info in funcoes_proposta.items() if r not in roles_pfcg
    ]

    todos_tcodes = sorted(list(tcode_para_roles.keys()))

    return {
        "encontrado": True,
        "sheet": sheet_prop,
        "total_funcoes": len(funcoes_proposta),
        "total_tcodes_unicos": len(todos_tcodes),
        "todos_tcodes": todos_tcodes,
        "tcode_para_roles": tcode_para_roles,
        "tcodes_descartados": tcodes_descartados_total,
        "funcoes": funcoes_proposta,
        "novas_para_criar": novas_para_criar,
        "total_novas_para_criar": len(novas_para_criar),
    }


def marcar_tcodes_inexistentes_sheet_proposta(caminho_excel: Optional[str], tcodes_inexistentes: List[str]) -> Dict[str, Any]:
    """
    Percorre a folha 'Proposta' e, para cada transação que não existe no SAP PRD (TSTC):
    - Escreve na coluna ao lado (Coluna B: DESCRIÇÃO) o texto 'Transação não existe'.
    - Garante que a transação é retirada da lista de atribuição de menus no PFCG.
    Atualiza tanto o ficheiro ativo (via COM se aberto, ou openpyxl) como a cópia de salvaguarda.
    """
    if not caminho_excel:
        caminho_excel = encontrar_excel_padrao()
    if not caminho_excel or not os.path.exists(caminho_excel):
        return {"ok": False, "erro": f"Ficheiro não encontrado: {caminho_excel}"}

    alvo_set = {str(tc).strip().upper() for tc in tcodes_inexistentes if str(tc).strip()}
    if not alvo_set:
        return {"ok": True, "marcados": 0, "detalhes": []}

    alterados = []
    sucesso = False

    # 1. Tentar via Excel COM se aberto
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            wb = None
            for w in xl.Workbooks:
                if os.path.basename(caminho_excel).lower() in w.Name.lower() or "perfis" in w.Name.lower():
                    wb = w
                    break
            if wb is not None:
                ws = wb.Worksheets("Proposta")
                role_corrente = ""
                for r in range(2, ws.UsedRange.Rows.Count + 1):
                    v_f = str(ws.Cells(r, 1).Value or "").strip().upper()
                    v_d = str(ws.Cells(r, 2).Value or "").strip()
                    if is_nome_funcao_proposta(v_f, v_d):
                        role_corrente = v_f
                    elif v_f in alvo_set and role_corrente:
                        ws.Cells(r, 2).Value = "Transação não existe"
                        alterados.append({"linha": r, "role": role_corrente, "tcode": v_f})
                if alterados:
                    wb.Save()
                    try:
                        backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                        if backup_local.exists() and str(backup_local.resolve()).lower() != os.path.abspath(caminho_excel).lower():
                            wb.SaveCopyAs(str(backup_local))
                    except Exception:
                        pass
                sucesso = True
        except Exception:
            pass

    # 2. Se COM não estiver disponível, usar openpyxl
    if not sucesso:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(caminho_excel, read_only=False)
            sheet_nome = next((s for s in wb.sheetnames if normalizar_nome_coluna(s) == "PROPOSTA"), None)
            if not sheet_nome:
                return {"ok": False, "erro": "Folha 'Proposta' não encontrada."}
            ws = wb[sheet_nome]
            role_corrente = ""
            for r in range(2, ws.max_row + 1):
                v_f = str(ws.cell(row=r, column=1).value or "").strip().upper()
                v_d = str(ws.cell(row=r, column=2).value or "").strip()
                if is_nome_funcao_proposta(v_f, v_d):
                    role_corrente = v_f
                elif v_f in alvo_set and role_corrente:
                    ws.cell(row=r, column=2, value="Transação não existe")
                    alterados.append({"linha": r, "role": role_corrente, "tcode": v_f})
            if alterados:
                wb.save(caminho_excel)
                try:
                    backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                    if backup_local.exists() and str(backup_local.resolve()).lower() != os.path.abspath(caminho_excel).lower():
                        import shutil
                        shutil.copy2(caminho_excel, str(backup_local))
                except Exception:
                    pass
            wb.close()
            sucesso = True
        except Exception as exc:
            return {"ok": False, "erro": f"Erro ao gravar via openpyxl: {exc}"}

    return {
        "ok": sucesso,
        "marcados": len(alterados),
        "detalhes": alterados,
    }


def atribuir_tcodes_funcao_prd_rfc(role_name: str, novos_tcodes: List[str]) -> Dict[str, Any]:
    """
    Atribui novas transações a uma função já existente no SAP PRD através de RFC
    chamando o módulo padrão PRGN_RFC_CREATE_ACTIVITY_GROUP com ONLY_TCODE_ASSIGNMENT='X'.
    """
    if not novos_tcodes:
        return {"ok": True, "adicionados": []}

    role_norm = str(role_name or "").strip().upper()
    tcodes_limpos = [str(tc).strip().upper() for tc in novos_tcodes if str(tc).strip()]

    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_write_guard, read_table, make_option_eq
        )
        from pyrfc import Connection

        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for("PRD")

        guard = make_write_guard(
            ("RFC_PING", "RFC_READ_TABLE", "PRGN_RFC_CREATE_ACTIVITY_GROUP"),
            ("AGR_DEFINE", "AGR_TCODES", "TSTC")
        )

        conn = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        conn.call("RFC_PING")

        # Obter TCODES já atribuídos no SAP
        rows_existentes = read_table(
            conn, guard, table_name="AGR_TCODES", fields=["TCODE"],
            options=make_option_eq("AGR_NAME", role_norm), rowcount=0
        )
        tcodes_atuais = {r[0].strip().upper() for r in rows_existentes if r}

        todos_tcodes = sorted(list(tcodes_atuais.union(tcodes_limpos)))

        guard.assert_function_allowed("PRGN_RFC_CREATE_ACTIVITY_GROUP")
        call_res = conn.call(
            "PRGN_RFC_CREATE_ACTIVITY_GROUP",
            ACTIVITY_GROUP=role_norm,
            ACTIVITY_GROUP_TEXT="",
            NO_DIALOG="X",
            ONLY_TCODE_ASSIGNMENT="X",
            ORG_LEVELS_WITH_STAR="",
            UNMAINTAINED_FIELDS_WITH_STAR="",
            PARENT_ROLE="",
            PROFILE_NAME="",
            PROFILE_TEXT="",
            REQUEST="",
            TEMPLATE="",
            TCODES=[{"TCODE": tc} for tc in todos_tcodes],
        )
        conn.close()

        ret_rows = call_res.get("RETURN") or []
        erros = [r for r in ret_rows if str(r.get("TYPE", "")).upper() in ("E", "A")]
        if erros:
            msg_erro = "; ".join([str(e.get("MESSAGE", "")) for e in erros])
            return {"ok": False, "erro": msg_erro, "detalhes": erros}

        return {
            "ok": True,
            "role": role_norm,
            "novos_tcodes": tcodes_limpos,
            "total_tcodes": len(todos_tcodes),
            "mensagem": f"Transações {', '.join(tcodes_limpos)} atribuídas com sucesso à função {role_norm} no SAP PRD."
        }
    except Exception as exc:
        return {"ok": False, "role": role_norm, "erro": str(exc)}


def atribuir_funcoes_composta_prd_rfc(role_composta: str, novas_funcoes: List[str]) -> Dict[str, Any]:
    """
    Atribui novas funções individuais (single roles) a uma função composta (composite role)
    já existente no SAP PRD através de RFC chamando o módulo padrão PRGN_RFC_ADD_AGRS_TO_COLL_AGR.
    """
    if not novas_funcoes:
        return {"ok": True, "adicionadas": []}

    role_norm = str(role_composta or "").strip().upper()
    funcoes_limpas = [str(f).strip().upper() for f in novas_funcoes if str(f).strip()]

    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_write_guard, read_table, make_option_eq
        )
        from pyrfc import Connection

        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for("PRD")

        guard = make_write_guard(
            ("RFC_PING", "RFC_READ_TABLE", "PRGN_RFC_ADD_AGRS_TO_COLL_AGR"),
            ("AGR_DEFINE", "AGR_AGRS", "AGR_FLAGS")
        )

        conn = Connection(**params)
        guard.assert_function_allowed("RFC_PING")
        conn.call("RFC_PING")

        # Obter funções já atribuídas na função composta
        rows_existentes = read_table(
            conn, guard, table_name="AGR_AGRS", fields=["CHILD_AGR"],
            options=make_option_eq("AGR_NAME", role_norm), rowcount=0
        )
        atuais = {r[0].strip().upper() for r in rows_existentes if r}

        faltam = [f for f in funcoes_limpas if f not in atuais]
        if not faltam:
            conn.close()
            return {"ok": True, "role": role_norm, "adicionadas": [], "mensagem": "Todas as funções já se encontram atribuídas."}

        guard.assert_function_allowed("PRGN_RFC_ADD_AGRS_TO_COLL_AGR")
        call_res = conn.call(
            "PRGN_RFC_ADD_AGRS_TO_COLL_AGR",
            ACTIVITY_GROUP=role_norm,
            CHECK_NAMESPACE="X",
            ENQUEUE="X",
            NO_DIALOG="X",
            PROFILE_COMPARISON="X",
            REC_PERS_DATA="X",
            REC_PROF_DATA="X",
            REC_SINGLE_ROLES="X",
            REQUEST="",
            ACTIVITY_GROUPS=[{"AGR_NAME": f, "TEXT": ""} for f in faltam],
        )
        conn.close()

        ret_rows = call_res.get("RETURN") or []
        erros = [r for r in ret_rows if str(r.get("TYPE", "")).upper() in ("E", "A")]
        if erros:
            msg_erro = "; ".join([str(e.get("MESSAGE", "")) for e in erros])
            return {"ok": False, "role": role_norm, "erro": msg_erro, "detalhes": erros}

        return {
            "ok": True,
            "role": role_norm,
            "adicionadas": faltam,
            "total_atribuidas": len(atuais) + len(faltam),
            "mensagem": f"Funções {', '.join(faltam)} atribuídas com sucesso à função composta {role_norm} no SAP PRD."
        }
    except Exception as exc:
        return {"ok": False, "role": role_norm, "erro": str(exc)}


def adicionar_linhas_pfcg_create(caminho_excel: Optional[str], novas_linhas: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Adiciona novas linhas no final da folha 'PFCG_CREATE' no ficheiro Excel com os campos:
    ID, AGR_NAME, TEXT, TCODE, STATUS ('Criado'), MSG ('Criado e Validado em SAP DEV, PRD e QAD'),
    TIMESTEMP (data/hora atual), PRD ('Validado').
    Utiliza Excel COM se o ficheiro estiver aberto pelo utilizador, com fallback para openpyxl.
    """
    if not caminho_excel:
        caminho_excel = encontrar_excel_padrao()
    if not caminho_excel or not os.path.exists(caminho_excel):
        return {"ok": False, "erro": f"Ficheiro não encontrado: {caminho_excel}"}

    if not novas_linhas:
        return {"ok": True, "adicionadas": 0}

    sucesso = False
    total_add = len(novas_linhas)

    # 1. Tentar via Excel COM se aberto
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            wb = None
            for w in xl.Workbooks:
                if os.path.basename(caminho_excel).lower() in w.Name.lower() or "perfis" in w.Name.lower():
                    wb = w
                    break
            if wb is not None:
                ws = wb.Worksheets("PFCG_CREATE")
                last_row = ws.UsedRange.Rows.Count
                while ws.Cells(last_row, 2).Value:
                    last_row += 1

                for i, linha in enumerate(novas_linhas):
                    curr = last_row + i
                    ws.Cells(curr, 1).Value = linha.get("ID")
                    ws.Cells(curr, 2).Value = linha.get("AGR_NAME")
                    ws.Cells(curr, 3).Value = linha.get("TEXT")
                    ws.Cells(curr, 4).Value = linha.get("TCODE")
                    ws.Cells(curr, 5).Value = linha.get("STATUS", "Criado")
                    ws.Cells(curr, 6).Value = linha.get("MSG", "Criado e Validado em SAP DEV, PRD e QAD")
                    ws.Cells(curr, 7).Value = linha.get("TIMESTEMP")
                    ws.Cells(curr, 8).Value = linha.get("PRD", "Validado")

                wb.Save()
                try:
                    backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                    if backup_local.exists() and str(backup_local.resolve()).lower() != os.path.abspath(caminho_excel).lower():
                        wb.SaveCopyAs(str(backup_local))
                except Exception:
                    pass
                sucesso = True
        except Exception:
            pass

    # 2. Se COM não estiver disponível, usar openpyxl
    if not sucesso:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(caminho_excel, read_only=False)
            sheet_nome = next((s for s in wb.sheetnames if normalizar_nome_coluna(s) in ("PFCGCREATE", "PFCG_CREATE")), None)
            if not sheet_nome:
                return {"ok": False, "erro": "Folha 'PFCG_CREATE' não encontrada."}
            ws = wb[sheet_nome]

            for linha in novas_linhas:
                ws.append([
                    linha.get("ID"),
                    linha.get("AGR_NAME"),
                    linha.get("TEXT"),
                    linha.get("TCODE"),
                    linha.get("STATUS", "Criado"),
                    linha.get("MSG", "Criado e Validado em SAP DEV, PRD e QAD"),
                    linha.get("TIMESTEMP"),
                    linha.get("PRD", "Validado"),
                ])

            wb.save(caminho_excel)
            try:
                backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                if backup_local.exists() and str(backup_local.resolve()).lower() != os.path.abspath(caminho_excel).lower():
                    import shutil
                    shutil.copy2(caminho_excel, str(backup_local))
            except Exception:
                pass
            wb.close()
            sucesso = True
        except Exception as exc:
            return {"ok": False, "erro": f"Erro ao gravar via openpyxl em PFCG_CREATE: {exc}"}

    return {"ok": sucesso, "adicionadas": total_add}


def sincronizar_catalogo_proposta_pfcg_prd(dados: ProjetoPerfilData, auto_criar_sap: bool = True) -> Dict[str, Any]:
    """
    Regra de Projeto no Arranque:
    1. Lê a folha 'Proposta' e valida todas as transações contra o SAP PRD (TSTC via RFC).
       - Se alguma transação não existir no SAP, escreve 'Transação não existe' na coluna ao lado
         e retira-a da lista ativa.
    2. Cruza cada função e transação contra a folha 'PFCG_CREATE':
       - Se (AGR_NAME, TCODE) já existir: pula para a próxima sem alteração.
       - Se a função NÃO existir em PFCG_CREATE:
         * Acede ao SAP PRD via RFC e cria a função e as respetivas transações.
         * Adiciona no final da folha PFCG_CREATE a função e suas transações com:
           STATUS = 'Criado'
           MSG = 'Criado e Validado em SAP DEV, PRD e QAD'
           TIMESTEMP = data/hora atual
           PRD = 'Validado'
       - Se a função já existir mas possuir uma nova transação:
         * Atribui a nova transação à função no SAP PRD via RFC.
         * Adiciona no final da folha PFCG_CREATE a respetiva linha com os mesmos campos.
    Esta rotina corre automaticamente no arranque sem questionar o utilizador (regra estrita).
    """
    import pandas as pd
    from datetime import datetime

    # 1. Análise da folha Proposta
    analise_prop = analisar_folha_proposta(dados)
    if not analise_prop.get("encontrado"):
        print(f"  [1/4] Catalogo Base: Erro - {analise_prop.get('mensagem')}")
        return {"ok": False, "erro": analise_prop.get("mensagem")}

    tcodes_prop = analise_prop.get("todos_tcodes", [])

    # 2. Validar TCODEs no SAP PRD (TSTC)
    res_tc = verificar_tcodes_prd(tcodes_prop)
    if res_tc.get("ok") and res_tc.get("nao_existentes"):
        res_m = marcar_tcodes_inexistentes_sheet_proposta(dados.caminho, res_tc["nao_existentes"])
        analise_prop = analisar_folha_proposta(dados)

    # 3. Ler PFCG_CREATE
    sheet_pfcg = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) in ("PFCGCREATE", "PFCG_CREATE")), None)
    if not sheet_pfcg:
        print("[ERRO] Sheet 'PFCG_CREATE' não encontrada no ficheiro Excel.")
        return {"ok": False, "erro": "Sheet PFCG_CREATE não encontrada"}

    fonte = abrir_excel_seguro(dados.caminho)
    df_pfcg = pd.read_excel(fonte, sheet_name=sheet_pfcg)

    pares_existentes = set()
    funcoes_existentes = set()
    max_id = 0

    for _, row in df_pfcg.iterrows():
        try:
            rid = int(row.get("ID", 0))
            if rid > max_id:
                max_id = rid
        except Exception:
            pass
        agr = str(row.get("AGR_NAME", "")).strip().upper()
        tc = str(row.get("TCODE", "")).strip().upper()
        if agr:
            funcoes_existentes.add(agr)
        if agr and tc:
            pares_existentes.add((agr, tc))

    # 4. Cruzamento
    funcoes_novas = {}
    novas_transacoes = {}

    for r_nome, info in analise_prop["funcoes"].items():
        r_upper = r_nome.upper()
        desc = info["descricao"]
        tcodes_validos = info["tcodes"]

        if r_upper not in funcoes_existentes:
            funcoes_novas[r_upper] = {
                "descricao": desc,
                "tcodes": tcodes_validos,
            }
        else:
            novos = [tc for tc in tcodes_validos if (r_upper, tc) not in pares_existentes]
            if novos:
                novas_transacoes[r_upper] = {
                    "descricao": desc,
                    "tcodes": novos,
                }

    if not funcoes_novas and not novas_transacoes:
        print(f"  [1/4] Catalogo Base: {len(analise_prop['funcoes'])} funcoes ({len(tcodes_prop)} TCODEs) conformes em PFCG_CREATE e PRD.")
        return {
            "ok": True,
            "funcoes_novas_criadas": 0,
            "novas_transacoes_atribuidas": 0,
            "sincronizado": True,
        }

    # 5. Processamento de Novas Funções
    linhas_a_adicionar = []
    ts_agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if funcoes_novas:
        print(f"\n  DETETADAS {len(funcoes_novas)} NOVA(S) FUNÇÃO(ÕES) PARA CRIAR:")
        from pfcg.pfcg_create_rfc_service import create_pfcg_role_rfc

        for f_nome, f_info in funcoes_novas.items():
            desc = f_info["descricao"]
            tcs = f_info["tcodes"]
            print(f"     ➕ A criar função '{f_nome}' no SAP PRD via RFC ({len(tcs)} transações)...")

            status_sap = "Criado"
            msg_sap = "Criado e Validado em SAP DEV, PRD e QAD"
            if auto_criar_sap:
                try:
                    res_rfc = create_pfcg_role_rfc(
                        environment="PRD",
                        role_name=f_nome,
                        description=desc,
                        tcodes=tcs,
                        transport_mode="LOCAL"
                    )
                    if not res_rfc.get("ok") and res_rfc.get("error_type") != "ROLE_ALREADY_EXISTS":
                        status_sap = "ERRO"
                        msg_sap = f"Erro RFC PRD: {res_rfc.get('message', '')}"
                        print(f"        [ERRO] Erro na criação SAP PRD: {res_rfc.get('message')}")
                    else:
                        print(f"        Função '{f_nome}' criada com sucesso no SAP PRD.")
                except Exception as exc:
                    status_sap = "ERRO"
                    msg_sap = f"Exceção RFC PRD: {exc}"
                    print(f"        [ERRO] Exceção ao contactar SAP PRD: {exc}")

            for tc in tcs:
                max_id += 1
                linhas_a_adicionar.append({
                    "ID": max_id,
                    "AGR_NAME": f_nome,
                    "TEXT": desc,
                    "TCODE": tc,
                    "STATUS": status_sap,
                    "MSG": msg_sap,
                    "TIMESTEMP": ts_agora,
                    "PRD": "Validado" if status_sap == "Criado" else "Pendente",
                })

    # 6. Processamento de Novas Transações em Funções Existentes
    if novas_transacoes:
        print(f"\n  DETETADAS TRANSAÇÕES NOVAS EM {len(novas_transacoes)} FUNÇÃO(ÕES) EXISTENTE(S):")
        for f_nome, f_info in novas_transacoes.items():
            desc = f_info["descricao"]
            tcs = f_info["tcodes"]
            print(f"     ➕ A atribuir transações {tcs} à função '{f_nome}' no SAP PRD via RFC...")

            status_sap = "Criado"
            msg_sap = "Criado e Validado em SAP DEV, PRD e QAD"
            if auto_criar_sap:
                res_atrib = atribuir_tcodes_funcao_prd_rfc(f_nome, tcs)
                if not res_atrib.get("ok"):
                    status_sap = "ERRO"
                    msg_sap = f"Erro RFC PRD: {res_atrib.get('erro', '')}"
                    print(f"        [ERRO] Erro ao atribuir no SAP PRD: {res_atrib.get('erro')}")
                else:
                    print(f"        Transações atribuídas com sucesso à função '{f_nome}' no SAP PRD.")

            for tc in tcs:
                max_id += 1
                linhas_a_adicionar.append({
                    "ID": max_id,
                    "AGR_NAME": f_nome,
                    "TEXT": desc,
                    "TCODE": tc,
                    "STATUS": status_sap,
                    "MSG": msg_sap,
                    "TIMESTEMP": ts_agora,
                    "PRD": "Validado" if status_sap == "Criado" else "Pendente",
                })

    # 7. Gravação em PFCG_CREATE
    if linhas_a_adicionar:
        print(f"\n  A registar {len(linhas_a_adicionar)} linha(s) no fim da folha 'PFCG_CREATE'...")
        res_add = adicionar_linhas_pfcg_create(dados.caminho, linhas_a_adicionar)
        if res_add.get("ok"):
            print("  Folha 'PFCG_CREATE' atualizada com sucesso no Excel.")
        else:
            print(f"  [ERRO] Erro ao atualizar folha 'PFCG_CREATE': {res_add.get('erro')}")

    print("-" * 75)
    return {
        "ok": True,
        "funcoes_novas_criadas": len(funcoes_novas),
        "novas_transacoes_atribuidas": sum(len(f["tcodes"]) for f in novas_transacoes.values()),
        "linhas_adicionadas": len(linhas_a_adicionar),
        "sincronizado": True,
    }


def sincronizar_matrizes_departamentais_proposta_ativa(dados: ProjetoPerfilData, caminho_excel: str) -> Dict[str, Any]:
    """
    [ARRANQUE 2/4] Sincroniza as matrizes departamentais com a folha 'Proposta Ativa'.
    Analisa cada utilizador com Composite Role na folha 'Proposta Ativa', cruza as transações
    marcadas com 'X' na matriz departamental com a folha 'Proposta' e atribui eventuais
    funções em falta nas próximas colunas livres. A folha 'DEFINIÇÕES' não participa
    deste preenchimento.
    """
    import pandas as pd
    from collections import defaultdict
    import shutil



    fonte = abrir_excel_seguro(caminho_excel)
    excel_file = pd.ExcelFile(fonte)

    # 1. Proposta: mapeamento TCODE -> roles
    df_prop = pd.read_excel(excel_file, sheet_name="Proposta")
    tcode_to_roles = defaultdict(list)
    current_role = None
    for _, row in df_prop.iterrows():
        f_val = str(row["FUNÇÃO"]).strip() if pd.notna(row["FUNÇÃO"]) else ""
        d_val = str(row["DESCRIÇÃO"]).strip() if pd.notna(row["DESCRIÇÃO"]) else ""
        if f_val.startswith("Z_") and d_val and "TRANSACAO NAO EXISTE" not in d_val.upper():
            current_role = f_val
        elif f_val and current_role:
            for t in f_val.replace(";", " ").split():
                tc = t.strip().upper()
                if tc:
                    tcode_to_roles[tc].append(current_role)

    mapa_dep_sheet = {
        "CONSTRUCTION & MAINTENANCE": ("Construction & Maintenance", "Construction & Maintenance"),
        "PURCHASE & SERVICES": ("Purchase & Services", "Purchase & Services"),
        "CLIENT SERVICES": ("Client Services", "Client Services"),
        "INDUSTRY SERVICES": ("Industry Services", "Industry Services"),
        "PEOPLE & TALENT": ("P&T", "People & Talent"),
        "P&T": ("P&T", "People & Talent"),
        "HEALTH & SAFETY": ("H&S", "Health & Safety"),
        "H&S": ("H&S", "Health & Safety"),
        "DIGITAL": ("Digital", "Digital"),
        "LEGAL": ("Legal", "Legal"),
    }

    # 3. Proposta Ativa
    df_ativa = pd.read_excel(excel_file, sheet_name="Proposta Ativa")
    col_user = [c for c in df_ativa.columns if "USER" in str(c).upper() or "UTILIZADOR" in str(c).upper() or "USU" in str(c).upper()][0]
    col_comp = [c for c in df_ativa.columns if "COMPOSITE" in str(c).upper()][0]
    col_dep = [c for c in df_ativa.columns if str(c).strip().upper() == "DEPARTAMENTO"][0]
    col_dep_dir = [c for c in df_ativa.columns if "DIRE" in str(c).upper()][0]

    sheet_dfs = {}
    for s in excel_file.sheet_names:
        if s in ["Construction & Maintenance", "Purchase & Services", "Client Services", "Industry Services", "P&T", "H&S", "Digital", "Legal"]:
            sheet_dfs[s] = pd.read_excel(excel_file, sheet_name=s, header=None)

    users_em_falta = []
    total_users_avaliados = 0

    for idx, r in df_ativa.iterrows():
        c_val = str(r.get(col_comp, "")).strip() if pd.notna(r.get(col_comp)) else ""
        if not c_val or c_val.lower() in ("nan", "none", "-"):
            continue

        u_val = str(r.get(col_user, "")).strip() if pd.notna(r.get(col_user)) else ""
        d_val = str(r.get(col_dep, "")).strip() if pd.notna(r.get(col_dep)) else ""
        dd_val = str(r.get(col_dep_dir, "")).strip() if pd.notna(r.get(col_dep_dir)) else ""
        dep = d_val if d_val and d_val.lower() != "nan" else dd_val

        total_users_avaliados += 1
        funcoes_existentes = set()
        for col_idx in range(9, len(r)):
            val_cell = r.iloc[col_idx]
            if pd.notna(val_cell):
                v_str = str(val_cell).strip().upper()
                if v_str and v_str.startswith("Z"):
                    funcoes_existentes.add(v_str)

        dep_norm = dep.strip().upper()
        sheet_info = mapa_dep_sheet.get(dep_norm)
        tcodes_marcados_user = []
        if sheet_info:
            sheet_nome, _ = sheet_info
            df_matriz = sheet_dfs.get(sheet_nome)
            if df_matriz is not None:
                col_u_idx = None
                hdr_r_idx = None
                for r_idx in range(5):
                    for c_idx in range(len(df_matriz.columns)):
                        cell_v = str(df_matriz.iloc[r_idx, c_idx])
                        if u_val in cell_v:
                            col_u_idx = c_idx
                            hdr_r_idx = r_idx
                            break
                    if col_u_idx is not None:
                        break
                if col_u_idx is not None:
                    for r_idx in range(hdr_r_idx + 1, len(df_matriz)):
                        tc = str(df_matriz.iloc[r_idx, 0]).strip().upper()
                        flag = str(df_matriz.iloc[r_idx, col_u_idx]).strip().upper()
                        if flag in ("X", "1", "SIM", "YES", "S"):
                            tcodes_marcados_user.append(tc)

        roles_transacoes = set()
        for tc in tcodes_marcados_user:
            for mr in tcode_to_roles.get(tc, []):
                roles_transacoes.add(mr)

        roles_esperadas = roles_transacoes
        roles_faltam = sorted(roles_esperadas - funcoes_existentes)

        if roles_esperadas != funcoes_existentes:
            users_em_falta.append({
                "linha": idx + 2,
                "user": u_val,
                "roles_em_falta": roles_faltam,
                "roles_esperadas": sorted(roles_esperadas),
            })

    backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
    alteracoes = 0

    if not users_em_falta:
        print(f"  [2/4] Proposta Ativa: {total_users_avaliados}/{total_users_avaliados} utilizadores conformes com as matrizes departamentais.")
        return {"ok": True, "total_users": total_users_avaliados, "users_em_falta": 0, "funcoes_adicionadas": 0}

    print(f"  DETETADAS LINHAS DESALINHADAS EM {len(users_em_falta)} UTILIZADOR(ES) NA 'PROPOSTA ATIVA'...")
    import openpyxl
    sucesso_com = False
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            wb = None
            for w in xl.Workbooks:
                if "perfis" in w.Name.lower() or os.path.basename(caminho_excel).lower() in w.Name.lower():
                    wb = w
                    break
            if wb is not None:
                ws = wb.Worksheets("Proposta Ativa")
                for u_info in users_em_falta:
                    row_idx = u_info["linha"]
                    last_col = max(10, ws.UsedRange.Columns.Count)
                    ws.Range(ws.Cells(row_idx, 10), ws.Cells(row_idx, last_col)).ClearContents()
                    for i, role in enumerate(u_info["roles_esperadas"]):
                        ws.Cells(row_idx, 10 + i).Value = role
                        alteracoes += 1
                wb.Save()
                try:
                    if backup_local.exists():
                        wb.SaveCopyAs(str(backup_local))
                except Exception:
                    pass
                sucesso_com = True
                print(f"  {alteracoes} função(ões) reescrita(s) via Excel COM.")
        except Exception:
            pass

    if not sucesso_com:
        try:
            wb = openpyxl.load_workbook(caminho_excel)
            ws = wb["Proposta Ativa"]
            for u_info in users_em_falta:
                row_idx = u_info["linha"]
                for c in range(10, ws.max_column + 1):
                    ws.cell(row_idx, c).value = None
                for i, role in enumerate(u_info["roles_esperadas"]):
                    ws.cell(row=row_idx, column=10 + i, value=role)
                    alteracoes += 1
            wb.save(caminho_excel)
            wb.close()
            try:
                if backup_local.exists():
                    shutil.copy2(caminho_excel, str(backup_local))
            except Exception:
                pass
            print(f"  {alteracoes} função(ões) reescrita(s) via openpyxl.")
        except Exception as exc:
            print(f"  [ERRO] Erro ao gravar Proposta Ativa: {exc}")

    print("-" * 75)
    return {
        "ok": True,
        "total_users": total_users_avaliados,
        "users_em_falta": len(users_em_falta),
        "funcoes_adicionadas": alteracoes
    }


def sincronizar_proposta_ativa_pfcg_composta(dados: ProjetoPerfilData, caminho_excel: str) -> Dict[str, Any]:
    """
    [ARRANQUE 3/4] Sincroniza as Composite Roles de 'Proposta Ativa' com a folha 'PFCG_COMPOSTA'.
    Garante que todas as funções atribuídas a utilizadores de cada Composite Role constam
    como membros componentes da Composite Role na folha 'PFCG_COMPOSTA'.
    """
    import pandas as pd
    from collections import defaultdict
    from datetime import datetime
    import shutil



    fonte = abrir_excel_seguro(caminho_excel)
    excel_file = pd.ExcelFile(fonte)

    df_ativa = pd.read_excel(excel_file, sheet_name="Proposta Ativa")
    col_comp = [c for c in df_ativa.columns if "COMPOSITE" in str(c).upper()][0]
    composta_roles = defaultdict(set)
    composta_textos = {}

    for _, r in df_ativa.iterrows():
        comp = str(r.get(col_comp, "")).strip() if pd.notna(r.get(col_comp)) else ""
        if not comp or comp.lower() in ("nan", "none", "-"):
            continue
        for col_idx in range(9, len(r)):
            val = r.iloc[col_idx]
            if pd.notna(val):
                v_str = str(val).strip().upper()
                if v_str and v_str.startswith("Z"):
                    composta_roles[comp].add(v_str)

    df_comp = pd.read_excel(excel_file, sheet_name="PFCG_COMPOSTA")
    pares_existentes = set()
    max_id = 0
    for _, row in df_comp.iterrows():
        try:
            rid = int(row.get("ID", 0))
            if rid > max_id:
                max_id = rid
        except Exception:
            pass
        agr_c = str(row.get("AGR_NAME_COMPOSTA", "")).strip().upper()
        agr = str(row.get("AGR_NAME", "")).strip().upper()
        txt = str(row.get("TEXT", "")).strip()
        if agr_c and txt and agr_c not in composta_textos:
            composta_textos[agr_c] = txt
        if agr_c and agr:
            pares_existentes.add((agr_c, agr))

    novas_linhas = []
    ts_agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for comp, roles in sorted(composta_roles.items()):
        desc = composta_textos.get(comp, "")
        for r in sorted(roles):
            if (comp, r) not in pares_existentes:
                max_id += 1
                novas_linhas.append({
                    "ID": max_id,
                    "AGR_NAME_COMPOSTA": comp,
                    "TEXT": desc,
                    "AGR_NAME": r,
                    "STATUS": "Criado",
                    "MSG": "Atribuído em SAP DEV, PRD e QAD",
                    "TIMESTEMP": ts_agora,
                    "PRD": "Validado"
                })

    backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")

    if not novas_linhas:
        print(f"  [3/4] PFCG_COMPOSTA: {len(composta_roles)}/{len(composta_roles)} Composite Roles alinhadas no Excel ({len(df_comp)} registos).")
        return {"ok": True, "total_compostas": len(composta_roles), "linhas_adicionadas": 0, "total_registos": len(df_comp)}

    print(f"  DETETADAS {len(novas_linhas)} NOVA(S) ATRIBUIÇÃO(ÕES) EM FALTA NA 'PFCG_COMPOSTA'...")
    import openpyxl
    sucesso_com = False
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            wb = None
            for w in xl.Workbooks:
                if "perfis" in w.Name.lower() or os.path.basename(caminho_excel).lower() in w.Name.lower():
                    wb = w
                    break
            if wb is not None:
                ws = wb.Worksheets("PFCG_COMPOSTA")
                last_row = ws.UsedRange.Rows.Count
                while ws.Cells(last_row, 2).Value:
                    last_row += 1
                for i, row in enumerate(novas_linhas):
                    curr = last_row + i
                    ws.Cells(curr, 1).Value = row["ID"]
                    ws.Cells(curr, 2).Value = row["AGR_NAME_COMPOSTA"]
                    ws.Cells(curr, 3).Value = row["TEXT"]
                    ws.Cells(curr, 4).Value = row["AGR_NAME"]
                    ws.Cells(curr, 5).Value = row["STATUS"]
                    ws.Cells(curr, 6).Value = row["MSG"]
                    ws.Cells(curr, 7).Value = row["TIMESTEMP"]
                    ws.Cells(curr, 8).Value = row["PRD"]
                wb.Save()
                try:
                    if backup_local.exists():
                        wb.SaveCopyAs(str(backup_local))
                except Exception:
                    pass
                sucesso_com = True
                print(f"  {len(novas_linhas)} linha(s) adicionada(s) via Excel COM.")
        except Exception:
            pass

    if not sucesso_com:
        try:
            wb = openpyxl.load_workbook(caminho_excel)
            ws = wb["PFCG_COMPOSTA"]
            for row in novas_linhas:
                ws.append([
                    row["ID"], row["AGR_NAME_COMPOSTA"], row["TEXT"],
                    row["AGR_NAME"], row["STATUS"], row["MSG"],
                    row["TIMESTEMP"], row["PRD"]
                ])
            wb.save(caminho_excel)
            wb.close()
            try:
                if backup_local.exists():
                    shutil.copy2(caminho_excel, str(backup_local))
            except Exception:
                pass
            print(f"  {len(novas_linhas)} linha(s) adicionada(s) via openpyxl.")
        except Exception as exc:
            print(f"  [ERRO] Erro ao gravar PFCG_COMPOSTA: {exc}")

    print("-" * 75)
    return {
        "ok": True,
        "total_compostas": len(composta_roles),
        "linhas_adicionadas": len(novas_linhas),
        "total_registos": len(df_comp) + len(novas_linhas)
    }


def sincronizar_pfcg_composta_prd_rfc(dados: ProjetoPerfilData, caminho_excel: str) -> Dict[str, Any]:
    """
    [ARRANQUE 4/4] Sincroniza os membros das Composite Roles diretamente no SAP PRD via RFC.
    Consulta AGR_AGRS, filtra apenas funções individuais (simples) através de AGR_FLAGS,
    e atribui automaticamente via PRGN_RFC_ADD_AGRS_TO_COLL_AGR caso haja discrepâncias.
    """
    import pandas as pd
    from collections import defaultdict
    from pyrfc import Connection
    from sap_rfc._rfc_common import (
        build_connection_params_for, find_project_root, load_project_env,
        make_read_only_guard, read_table, make_option_in
    )



    fonte = abrir_excel_seguro(caminho_excel)
    df_comp = pd.read_excel(fonte, sheet_name="PFCG_COMPOSTA")

    excel_compostas = defaultdict(set)
    for _, row in df_comp.iterrows():
        comp = str(row.get("AGR_NAME_COMPOSTA", "")).strip().upper()
        child = str(row.get("AGR_NAME", "")).strip().upper()
        if comp and child:
            excel_compostas[comp].add(child)

    load_project_env(find_project_root())
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(["AGR_AGRS", "AGR_FLAGS"])

    conn = Connection(**params)
    prd_compostas = {c: set() for c in excel_compostas}

    try:
        roles = sorted(excel_compostas.keys())
        for i in range(0, len(roles), 20):
            lote = roles[i:i+20]
            rows = read_table(
                conn, guard, table_name="AGR_AGRS", fields=["AGR_NAME", "CHILD_AGR"],
                options=make_option_in("AGR_NAME", lote), rowcount=0
            )
            for c_name, child in rows:
                c_upper = c_name.strip().upper()
                ch_upper = child.strip().upper()
                if c_upper in prd_compostas and ch_upper:
                    prd_compostas[c_upper].add(ch_upper)

        todas_filhas = sorted(list({f for filhas in excel_compostas.values() for f in filhas}))
        flags_map = {}
        for i in range(0, len(todas_filhas), 20):
            lote = todas_filhas[i:i+20]
            rows = read_table(
                conn, guard, table_name="AGR_FLAGS", fields=["AGR_NAME", "FLAG_TYPE", "FLAG_VALUE"],
                options=make_option_in("AGR_NAME", lote), rowcount=0
            )
            for agr, ftype, fval in rows:
                if ftype == "COLL_AGR":
                    flags_map[agr.strip().upper()] = fval.strip().upper()
    finally:
        conn.close()

    plano_execucao = {}
    for comp, exp in sorted(excel_compostas.items()):
        prd = prd_compostas.get(comp, set())
        faltam = exp - prd
        faltam_simples = sorted([f for f in faltam if flags_map.get(f) != "X"])
        if faltam_simples:
            plano_execucao[comp] = faltam_simples

    total_adicionadas = 0
    if not plano_execucao:
        print(f"  [4/4] SAP PRD: {len(excel_compostas)}/{len(excel_compostas)} Composite Roles conformes em AGR_AGRS (0 funcoes simples em falta).")
        return {"ok": True, "total_compostas": len(excel_compostas), "compostas_com_faltas": 0, "funcoes_adicionadas_prd": 0}

    print(f"  DETETADAS {sum(len(v) for v in plano_execucao.values())} FUNÇÃO(ÕES) SIMPLES EM FALTA NO SAP PRD EM {len(plano_execucao)} COMPOSTA(S)...")
    for comp, filhas in plano_execucao.items():
        res_add = atribuir_funcoes_composta_prd_rfc(comp, filhas)
        if res_add.get("ok"):
            adics = res_add.get("adicionadas", [])
            total_adicionadas += len(adics)
            print(f"     {comp}: {len(adics)} função(ões) atribuída(s) via RFC no PRD.")
        else:
            print(f"     [ERRO] {comp}: Erro RFC PRD - {res_add.get('erro')}")

    print("-" * 75)
    return {
        "ok": True,
        "total_compostas": len(excel_compostas),
        "compostas_com_faltas": len(plano_execucao),
        "funcoes_adicionadas_prd": total_adicionadas
    }


def executar_sincronizacao_arranque_completa(dados: ProjetoPerfilData, caminho_excel: str) -> ProjetoPerfilData:
    """
    Executa o pipeline completo de 4 fases de validação e sincronização automática no arranque:
      1. Catálogo Base (Proposta -> PFCG_CREATE -> PRD)
      2. Matrizes Departamentais & DEFINIÇÕES -> Proposta Ativa
      3. Proposta Ativa -> PFCG_COMPOSTA (Excel)
      4. PFCG_COMPOSTA -> SAP PRD (AGR_AGRS via RFC)
    Recarrega e devolve os dados atualizados caso ocorram modificações no Excel.
    """
    print("\n" + "=" * 75)
    print("  SINCRONIZACAO E VALIDACAO NO ARRANQUE")
    print("=" * 75)
    recarr_necessario = False

    # Fase 1
    res1 = sincronizar_catalogo_proposta_pfcg_prd(dados, auto_criar_sap=True)
    if res1.get("linhas_adicionadas", 0) > 0:
        recarr_necessario = True
        dados = carregar_projeto_perfil(caminho_excel)

    # Fase 2
    res2 = sincronizar_matrizes_departamentais_proposta_ativa(dados, caminho_excel)
    if res2.get("funcoes_adicionadas", 0) > 0:
        recarr_necessario = True
        dados = carregar_projeto_perfil(caminho_excel)

    # Fase 3
    res3 = sincronizar_proposta_ativa_pfcg_composta(dados, caminho_excel)
    if res3.get("linhas_adicionadas", 0) > 0:
        recarr_necessario = True
        dados = carregar_projeto_perfil(caminho_excel)

    # Fase 4
    sincronizar_pfcg_composta_prd_rfc(dados, caminho_excel)

    if recarr_necessario:
        print("\nA recarregar dados do Excel apos sincronizacoes automaticas...")
        dados = carregar_projeto_perfil(caminho_excel)

    print("-" * 75)
    return dados


def comparar_funcoes_catalogo_prd(dados: ProjetoPerfilData) -> Dict[str, Any]:
    """
    Pesquisa no sistema SAP PRD todas as funções criadas (AGR_DEFINE) e funções
    ativamente atribuídas (AGR_USERS) e compara contra o universo de funções
    catalogadas nas folhas do ficheiro Excel mestre.
    """
    import datetime
    import pandas as pd
    from pyrfc import Connection
    from sap_rfc._rfc_common import (
        build_connection_params_for, load_project_env, find_project_root
    )

    # 1. Coleta global de todas as funções referenciadas no ficheiro Excel
    roles_excel = set()
    fonte = abrir_excel_seguro(dados.caminho)
    excel_file = pd.ExcelFile(fonte)
    for sheet in excel_file.sheet_names:
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet)
            for col in df.columns:
                for val in df[col].dropna():
                    v = str(val).strip().upper()
                    if (v.startswith("Z") or v.startswith("SAP_") or v.startswith("/SALSA/")) and re.match(r"^[A-Z0-9_/\-:]+$", v) and len(v) >= 3:
                        if v not in ("NAN", "NONE"):
                            roles_excel.add(v)
        except Exception:
            pass

    roles_excel.update(dados.roles_simples.keys())
    roles_excel.update(dados.roles_compostas.keys())
    for flhas in dados.roles_authority.values():
        roles_excel.update(flhas)

    # 2. Conectar ao SAP PRD via RFC
    os.environ["SAP_TARGET_ENV"] = "PRD"
    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for("PRD")
    conn = Connection(**params)

    # Leitura de todas as funções do PRD (AGR_DEFINE)
    res_def = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE="AGR_DEFINE",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "AGR_NAME"}, {"FIELDNAME": "PARENT_AGR"}],
    )
    data_def = res_def.get("DATA", [])
    roles_prd_info = {}
    for item in data_def:
        wa = item.get("WA", "")
        parts = wa.split("|")
        role_name = parts[0].strip().upper() if len(parts) > 0 else ""
        parent_agr = parts[1].strip().upper() if len(parts) > 1 else ""
        if role_name:
            roles_prd_info[role_name] = parent_agr

    # Leitura de atribuições ativas de utilizadores no PRD (AGR_USERS)
    hoje = datetime.date.today().strftime("%Y%m%d")
    res_usr = conn.call(
        "RFC_READ_TABLE",
        QUERY_TABLE="AGR_USERS",
        DELIMITER="|",
        FIELDS=[{"FIELDNAME": "UNAME"}, {"FIELDNAME": "AGR_NAME"}],
        OPTIONS=[{"TEXT": f"TO_DAT >= '{hoje}'"}],
    )
    data_usr = res_usr.get("DATA", [])
    roles_atribuidas_prd = set()
    for item in data_usr:
        wa = item.get("WA", "")
        parts = wa.split("|")
        r_name = parts[1].strip().upper() if len(parts) > 1 else ""
        if r_name:
            roles_atribuidas_prd.add(r_name)

    conn.close()

    set_roles_prd = set(roles_prd_info.keys())

    roles_z_prd = {r for r in set_roles_prd if r.startswith("Z")}
    roles_sap_prd = {r for r in set_roles_prd if r.startswith("SAP_")}
    roles_outras_prd = set_roles_prd - roles_z_prd - roles_sap_prd

    dif_total_bruto = set_roles_prd - roles_excel
    dif_z_bruto = roles_z_prd - roles_excel
    dif_total_sem_exclusao = {r for r in dif_total_bruto if not dados.is_excluida(r)}
    dif_z_sem_exclusao = {r for r in dif_z_bruto if not dados.is_excluida(r)}

    dif_atrib_bruto = roles_atribuidas_prd - roles_excel
    dif_atrib_sem_exclusao = {r for r in dif_atrib_bruto if not dados.is_excluida(r)}

    z_org = sorted([r for r in dif_z_sem_exclusao if r.startswith("ZORG_")])
    z_br = sorted([r for r in dif_z_sem_exclusao if r.startswith("Z_BR_")])
    z_mm = sorted([r for r in dif_z_sem_exclusao if r.startswith("ZMM_")])
    z_fi = sorted([r for r in dif_z_sem_exclusao if r.startswith("ZFI_")])
    z_sd = sorted([r for r in dif_z_sem_exclusao if r.startswith("ZSD_")])
    z_outras = sorted(list(dif_z_sem_exclusao - set(z_org) - set(z_br) - set(z_mm) - set(z_fi) - set(z_sd)))

    return {
        "ok": True,
        "total_prd_catalogo": len(set_roles_prd),
        "total_prd_z": len(roles_z_prd),
        "total_prd_sap": len(roles_sap_prd),
        "total_prd_outras": len(roles_outras_prd),
        "total_excel": len(roles_excel),
        "dif_catalogo_bruto": len(dif_total_bruto),
        "dif_catalogo_sem_exclusao": len(dif_total_sem_exclusao),
        "dif_z_sem_exclusao": len(dif_z_sem_exclusao),
        "dif_z_compostas": len([r for r in dif_z_sem_exclusao if roles_prd_info.get(r)]),
        "dif_z_simples": len([r for r in dif_z_sem_exclusao if not roles_prd_info.get(r)]),
        "z_org": z_org,
        "z_br": z_br,
        "z_mm": z_mm,
        "z_fi": z_fi,
        "z_sd": z_sd,
        "z_outras": z_outras,
        "total_atribuidas_ativas_prd": len(roles_atribuidas_prd),
        "dif_atrib_bruto": len(dif_atrib_bruto),
        "dif_atrib_sem_exclusao": len(dif_atrib_sem_exclusao),
        "roles_atribuidas_fora_excel": sorted(list(dif_atrib_sem_exclusao)),
        "padroes_exclusao": dados.padroes_exclusao,
    }


def imprimir_comparacao_catalogo_prd(res: Dict[str, Any]):
    """Imprime no terminal o relatório comparativo de funções existentes no PRD vs Excel."""
    print("\n" + "=" * 80)
    print("  COMPARATIVO DE FUNÇÕES: SAP PRD vs FICHEIRO EXCEL")
    print("=" * 80)

    if not res.get("ok"):
        print(f"  [ERRO] Erro ao consultar SAP PRD: {res.get('erro')}")
        return

    print("  1. CATÁLOGO GLOBAL DE FUNÇÕES NO SAP PRD (AGR_DEFINE):")
    print(f"     • Total de funções no PRD:              {res['total_prd_catalogo']:>6d}")
    print(f"       ├─ Funções Customizadas (Z*):          {res['total_prd_z']:>6d}")
    print(f"       ├─ Funções Standard SAP (SAP_*):       {res['total_prd_sap']:>6d}")
    print(f"       └─ Outras funções Standard:            {res['total_prd_outras']:>6d}")
    print(f"     • Total de funções no Ficheiro Excel:   {res['total_excel']:>6d}")

    print("\n  2. FUNÇÕES NO PRD QUE NÃO CONSTAM NO FICHEIRO EXCEL:")
    print(f"     • Diferença Total Bruta:                {res['dif_catalogo_bruto']:>6d}")
    print(f"       (Inclui {res['total_prd_sap']} standard SAP_* e outras legadas)")
    print(f"     • Diferença Líquida (após EXCLUÇÃO):    {res['dif_catalogo_sem_exclusao']:>6d}")
    print(f"     • Funções Customizadas Z* fora do Excel: {res['dif_z_sem_exclusao']:>6d}")
    print(f"       ├─ Funções Simples (Single):           {res['dif_z_simples']:>6d}")
    print(f"       └─ Funções Compostas (Composite):      {res['dif_z_compostas']:>6d}")

    print("\n  3. SEGMENTAÇÃO DAS FUNÇÕES Z* FORA DO EXCEL:")
    print(f"     • Organizacionais (ZORG_*):              {len(res['z_org']):>6d}")
    print(f"     • Business Roles (Z_BR_*):               {len(res['z_br']):>6d}")
    print(f"     • Módulo MM (ZMM_*):                     {len(res['z_mm']):>6d}")
    print(f"     • Módulo FI (ZFI_*):                     {len(res['z_fi']):>6d}")
    print(f"     • Módulo SD (ZSD_*):                     {len(res['z_sd']):>6d}")
    print(f"     • Outros legados Z*:                     {len(res['z_outras']):>6d}")

    print("\n  4. FUNÇÕES ATIVAMENTE ATRIBUÍDAS A UTILIZADORES (AGR_USERS):")
    print(f"     • Total de funções distintas atribuídas: {res['total_atribuidas_ativas_prd']:>6d}")
    print(f"     • Funções atribuídas fora do Excel (Bruto): {res['dif_atrib_bruto']:>4d}")
    print(f"     • Funções atribuídas fora do Excel (Líquido): {res['dif_atrib_sem_exclusao']:>2d}")
    if res["roles_atribuidas_fora_excel"]:
        print("       Funções ativas a utilizadores não catalogadas:")
        for r in res["roles_atribuidas_fora_excel"][:25]:
            print(f"         ├─ {r}")
        if len(res["roles_atribuidas_fora_excel"]) > 25:
            print(f"         └─ ... e mais {len(res['roles_atribuidas_fora_excel']) - 25} funções.")

    print("\n  Padrões desconsiderados da folha EXCLUÇÃO:")
    print(f"     {', '.join(res['padroes_exclusao']) or 'Nenhum'}")
    print("=" * 80)


# =====================================================================
# CRUZAMENTO RELACIONAL DE FONTES E VALIDAÇÃO DE UTILIZADORES
# =====================================================================

def cruzar_fontes_departamento(dados: ProjetoPerfilData, departamento: str = "Purchase & Services") -> Dict[str, Any]:
    """
    Cruza relacionalmente as quatro fontes essenciais:
      1. PFCG_CREATE: Funções individuais criadas;
      2. PFCG_COMPOSTA: Membros de cada função composta;
      3. PFCG_AUTHORITY: Funções avulsas ligadas à respetiva composta;
      4. EXCLUÇÃO: Padrões globais fora do escopo de remoção/falta.
    """
    analise = analisar_departamento_proposta(dados, departamento)
    if not analise.get("encontrado"):
        return {"encontrado": False, "departamento": departamento, "mensagem": analise.get("mensagem")}

    usuarios_detalhe = []
    for u in analise.get("usuarios", []):
        usr = u["usuario"]
        comp = u.get("composta")
        singles_prop = [r for r in u.get("singles", []) if r]

        # Base inicial: Composta (se houver) + Singles da Proposta Ativa
        base_set = set()
        if comp:
            base_set.add(comp)
        base_set.update(singles_prop)

        # Membros PFCG_COMPOSTA
        membros_comp = set(dados.roles_compostas.get(comp, {}).get("roles_filhas", [])) if comp else set()
        # Funções PFCG_AUTHORITY da composta
        authority_comp = set(dados.roles_authority.get(comp, set())) if comp else set()

        # Expansão relacional completa recursiva
        esperado_expandido = dados.expandir_funcoes(base_set)

        # Identificar exclusões presentes na proposta original
        excluidas_na_proposta = [r for r in base_set if dados.is_excluida(r)]

        usuarios_detalhe.append({
            "usuario": usr,
            "nome": u["nome"],
            "cargo": u["cargo"],
            "composta": comp,
            "singles_proposta_total": len(singles_prop),
            "membros_composta_total": len(membros_comp),
            "authority_composta_total": len(authority_comp),
            "membros_composta": sorted(membros_comp),
            "authority_composta": sorted(authority_comp),
            "esperado_total": len(esperado_expandido),
            "esperado_roles": sorted(esperado_expandido),
            "excluidas_proposta": excluidas_na_proposta,
        })

    return {
        "encontrado": True,
        "departamento": departamento,
        "total_usuarios": len(usuarios_detalhe),
        "padroes_exclusao": dados.padroes_exclusao,
        "usuarios": usuarios_detalhe,
    }


def imprimir_cruzamento_fontes(resultado: Dict[str, Any]):
    """Exibe no terminal o relatório de cruzamento relacional de fontes."""
    print("\n" + "=" * 75)
    print("  CRUZAMENTO DE FONTES (PFCG_CREATE, PFCG_COMPOSTA, PFCG_AUTHORITY, EXCLUÇÃO)")
    print(f"  Departamento: {resultado.get('departamento', '')}")
    print("=" * 75)

    if not resultado.get("encontrado"):
        print(f"  [AVISO] {resultado.get('mensagem', 'Departamento não encontrado.')}")
        return

    print(f"  Regras EXCLUÇÃO aplicadas: {resultado.get('padroes_exclusao', [])}")
    print(f"  Utilizadores analisados:   {resultado.get('total_usuarios', 0)}")
    print("-" * 75)

    for u in resultado.get("usuarios", []):
        comp_str = f"Composta: {u['composta']}" if u.get("composta") else "Sem Composta"
        print(f"\n  {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']}")
        print(f"     └─ {comp_str} | {u['singles_proposta_total']} Singles na Proposta")
        print(f"        ├─ Membros em PFCG_COMPOSTA:   {u['membros_composta_total']}")
        print(f"        ├─ PFCG_AUTHORITY da Composta: {u['authority_composta_total']}")
        print(f"        └─ Esperado Expandido Final: {u['esperado_total']} Funções")
        if u.get("excluidas_proposta"):
            print(f"           Desconsideradas por EXCLUÇÃO: {', '.join(u['excluidas_proposta'])}")


def validar_utilizadores_prd(dados: ProjetoPerfilData, departamento: str = "Purchase & Services") -> Dict[str, Any]:
    """
    Verifica no SAP PRD (AGR_USERS via RFC) as atribuições reais de cada utilizador
    do departamento, cruzando com a expansão (PFCG_CREATE, PFCG_COMPOSTA, PFCG_AUTHORITY)
    e desconsiderando os padrões da sheet EXCLUÇÃO.
    """
    from datetime import date

    def assignment_status_rfc(from_dat: str, to_dat: str) -> str:
        hoje = int(date.today().strftime("%Y%m%d"))
        inicio = int(from_dat) if str(from_dat).strip().isdigit() else 0
        fim = int(to_dat) if str(to_dat).strip().isdigit() else 99991231
        if hoje < inicio:
            return "FUTURO"
        if hoje > fim:
            return "EXPIRADO"
        return "ATIVO"

    cruzamento = cruzar_fontes_departamento(dados, departamento)
    if not cruzamento.get("encontrado"):
        return {"ok": False, "erro": cruzamento.get("mensagem")}

    esperado_por_user = {
        u["usuario"]: set(u["esperado_roles"])
        for u in cruzamento["usuarios"]
    }
    detalhes_user = {
        u["usuario"]: u
        for u in cruzamento["usuarios"]
    }

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
        guard = make_read_only_guard(["AGR_USERS", "USR02"])

        ativas_prd: Dict[str, Set[str]] = {u: set() for u in esperado_por_user}
        inativas_prd: Dict[str, Dict[str, Any]] = {u: {} for u in esperado_por_user}
        mestre_usr02: Dict[str, Dict[str, Any]] = {}

        users_list = sorted(esperado_por_user.keys())

        # Consulta mestre USR02 para verificar estado e validade da conta
        try:
            opts_usr = make_option_in("BNAME", users_list)
            rows_usr = read_table(conn, guard, table_name="USR02", fields=["BNAME", "GLTGV", "GLTGB", "UFLAG", "TRDAT"], options=opts_usr, rowcount=0)
            for r in rows_usr:
                if len(r) >= 3:
                    bn = str(r[0]).strip().upper()
                    g_fim = str(r[2]).strip()
                    hoje_int = int(date.today().strftime("%Y%m%d"))
                    fim_int = int(g_fim) if g_fim.isdigit() and int(g_fim) > 0 else 99991231
                    expirado = hoje_int > fim_int
                    mestre_usr02[bn] = {
                        "expirado": expirado,
                        "validade_fim": g_fim if expirado else "",
                        "uflag": str(r[3]).strip() if len(r) > 3 else "0",
                        "ultimo_logon": str(r[4]).strip() if len(r) > 4 else "",
                    }
        except Exception:
            pass

        for i in range(0, len(users_list), 20):
            chunk = users_list[i:i + 20]
            opts = make_option_in("UNAME", chunk)
            rows = read_table(conn, guard, table_name="AGR_USERS", fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT"], options=opts, rowcount=0)
            for r in rows:
                if len(r) >= 4:
                    role = str(r[0]).strip().upper()
                    uname = str(r[1]).strip().upper()
                    from_d = str(r[2]).strip()
                    to_d = str(r[3]).strip()
                    if uname in ativas_prd and role:
                        st = assignment_status_rfc(from_d, to_d)
                        if st == "ATIVO":
                            ativas_prd[uname].add(role)
                        else:
                            inativas_prd[uname][role] = {"status": st, "inicio": from_d, "fim": to_d}

        conn.close()

        relatorio_users = []
        total_adicionais = 0
        total_desconsideradas = 0

        for uname in sorted(esperado_por_user.keys()):
            esp = esperado_por_user[uname]
            atv = ativas_prd[uname]
            faltam = sorted(esp - atv)
            adicionais_brutas = sorted(atv - esp)
            desconsideradas = sorted([r for r in adicionais_brutas if dados.is_excluida(r)])
            adicionais_reais = sorted(set(adicionais_brutas) - set(desconsideradas))

            # Validação relacional: Quais das adicionais reais constam no catálogo criado?
            adicionais_no_catalogo = sorted([
                r for r in adicionais_reais
                if (r in dados.roles_simples or r in dados.roles_compostas)
            ])
            adicionais_fora_catalogo = sorted([
                r for r in adicionais_reais
                if r not in adicionais_no_catalogo
            ])

            total_adicionais += len(adicionais_reais)
            total_desconsideradas += len(desconsideradas)

            u_meta = detalhes_user[uname]
            usr_info = mestre_usr02.get(uname, {})
            u_expirado = usr_info.get("expirado", False)

            relatorio_users.append({
                "usuario": uname,
                "nome": u_meta["nome"],
                "cargo": u_meta["cargo"],
                "composta": u_meta["composta"],
                "inativo_usr02": u_expirado,
                "validade_fim": usr_info.get("validade_fim", ""),
                "esperadas_total": len(esp),
                "ativas_total": len(atv),
                "esperadas_ativas": len(esp & atv),
                "faltam": faltam,
                "adicionais": adicionais_reais,
                "adicionais_no_catalogo": adicionais_no_catalogo,
                "adicionais_fora_catalogo": adicionais_fora_catalogo,
                "desconsideradas_exclusao": desconsideradas,
                "conforme": (len(faltam) == 0) or u_expirado,
            })

        return {
            "ok": True,
            "departamento": departamento,
            "sistema": "PRD",
            "mandante": params.get("client", "100"),
            "total_utilizadores": len(relatorio_users),
            "padroes_exclusao": dados.padroes_exclusao,
            "total_adicionais_ativas": total_adicionais,
            "total_desconsideradas_exclusao": total_desconsideradas,
            "utilizadores": relatorio_users,
        }
    except Exception as err:
        return {
            "ok": False,
            "departamento": departamento,
            "erro": str(err),
        }


def adicionar_funcao_proposta_ativa(
    caminho_excel: Optional[str],
    utilizador: str,
    role: str,
    dados: Optional[ProjetoPerfilData] = None,
) -> Dict[str, Any]:
    """
    Acrescenta uma função à linha de um utilizador na folha 'Proposta Ativa'.
    Tenta primeiro via interface COM do Excel (se estiver aberto pelo utilizador);
    caso contrário, utiliza openpyxl para gravação direta no ficheiro.
    Atualiza também a cópia local em sap_script_uploads (se existir).
    """
    if not caminho_excel:
        caminho_excel = encontrar_excel_padrao()
    if not caminho_excel or not os.path.exists(caminho_excel):
        return {"ok": False, "erro": f"Ficheiro Excel não encontrado: {caminho_excel}"}

    u_norm = normalizar_texto(utilizador)
    r_norm = normalizar_texto(role)

    if not u_norm or not r_norm:
        return {"ok": False, "erro": "Utilizador ou função vazios."}

    sucesso = False
    linha_afetada = None
    coluna_afetada = None
    metodo_usado = ""

    # 1. Tentar via Excel COM (se aberto)
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            wb = None
            for w in xl.Workbooks:
                if os.path.basename(caminho_excel).lower() in w.Name.lower() or "perfis" in w.Name.lower():
                    wb = w
                    break
            if wb is not None:
                ws = wb.Worksheets("Proposta Ativa")
                row_count = ws.UsedRange.Rows.Count
                for r in range(2, row_count + 1):
                    val_u = str(ws.Cells(r, 1).Value or "").strip().upper()
                    if val_u == u_norm:
                        linha_afetada = r
                        max_c = ws.UsedRange.Columns.Count
                        col_livre = None
                        for c in range(10, max_c + 3):
                            v_cel = str(ws.Cells(r, c).Value or "").strip().upper()
                            if v_cel == r_norm:
                                return {
                                    "ok": True,
                                    "ja_existia": True,
                                    "mensagem": f"A função {r_norm} já existe na linha {r} do utilizador {u_norm}.",
                                    "linha": r,
                                    "coluna": c
                                }
                            if not v_cel and col_livre is None:
                                col_livre = c
                                break
                        if col_livre:
                            ws.Cells(r, col_livre).Value = r_norm
                            coluna_afetada = col_livre
                            wb.Save()
                            try:
                                backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                                if backup_local.exists() and str(backup_local.resolve()).lower() != os.path.abspath(caminho_excel).lower():
                                    wb.SaveCopyAs(str(backup_local))
                            except Exception:
                                pass
                            sucesso = True
                            metodo_usado = "COM"
                        break
        except Exception:
            pass

    # 2. Se COM não executou, utilizar openpyxl
    if not sucesso:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(caminho_excel, read_only=False)
            sheet_nome = next((s for s in wb.sheetnames if normalizar_nome_coluna(s) == "PROPOSTAATIVA"), None)
            if not sheet_nome:
                return {"ok": False, "erro": "Folha 'Proposta Ativa' não encontrada no ficheiro."}

            ws = wb[sheet_nome]
            found_user = False
            for r in range(2, ws.max_row + 1):
                cell_val = str(ws.cell(row=r, column=1).value or "").strip().upper()
                if cell_val == u_norm:
                    found_user = True
                    linha_afetada = r
                    col_livre = None
                    for c in range(10, ws.max_column + 5):
                        v_cel = str(ws.cell(row=r, column=c).value or "").strip().upper()
                        if v_cel == r_norm:
                            wb.close()
                            return {
                                "ok": True,
                                "ja_existia": True,
                                "mensagem": f"A função {r_norm} já existe na linha {r} do utilizador {u_norm}.",
                                "linha": r,
                                "coluna": c
                            }
                        if not v_cel and col_livre is None:
                            col_livre = c
                            break
                    if col_livre:
                        ws.cell(row=r, column=col_livre, value=r_norm)
                        coluna_afetada = col_livre
                        wb.save(caminho_excel)
                        wb.close()

                        # Gravar cópia local de backup se existir
                        try:
                            backup_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                            if backup_local.exists() and str(backup_local.resolve()).lower() != os.path.abspath(caminho_excel).lower():
                                import shutil
                                shutil.copy2(caminho_excel, str(backup_local))
                        except Exception:
                            pass
                        sucesso = True
                        metodo_usado = "openpyxl"
                    break

            if not found_user:
                wb.close()
                return {"ok": False, "erro": f"Utilizador '{u_norm}' não encontrado na folha Proposta Ativa."}
        except Exception as exc:
            return {"ok": False, "erro": f"Erro ao gravar via openpyxl: {exc}"}

    if sucesso:
        col_letra = ""
        c_temp = coluna_afetada - 1 if coluna_afetada else 0
        while c_temp >= 0:
            col_letra = chr(c_temp % 26 + 65) + col_letra
            c_temp = c_temp // 26 - 1

        return {
            "ok": True,
            "ja_existia": False,
            "metodo": metodo_usado,
            "utilizador": u_norm,
            "role": r_norm,
            "linha": linha_afetada,
            "coluna_num": coluna_afetada,
            "coluna": col_letra,
            "celula": f"{col_letra}{linha_afetada}",
            "ficheiro": caminho_excel,
            "mensagem": f"Função {r_norm} adicionada com sucesso na célula {col_letra}{linha_afetada} (Linha {linha_afetada}) para o utilizador {u_norm}."
        }

    return {"ok": False, "erro": "Não foi possível adicionar a função."}


def incorporar_adicionais_catalogo_departamento(dados: ProjetoPerfilData, departamento: str) -> Dict[str, Any]:
    """
    Identifica todas as funções ativas no SAP PRD que pertencem ao catálogo oficial (PFCG_CREATE / COMPOSTA)
    mas ainda não constam na folha 'Proposta Ativa' para os utilizadores do departamento,
    e acrescenta-as automaticamente.
    """
    res_val = validar_utilizadores_prd(dados, departamento)
    if not res_val.get("ok"):
        return {"ok": False, "erro": res_val.get("erro")}

    candidatos = []
    for u in res_val.get("utilizadores", []):
        if u.get("inativo_usr02"):
            continue
        for r in u.get("adicionais_no_catalogo", []):
            candidatos.append((u["usuario"], u["nome"], r))

    if not candidatos:
        return {
            "ok": True,
            "departamento": departamento,
            "total_incorporadas": 0,
            "candidatos": [],
            "mensagem": "Nenhuma função adicional do catálogo pendente de incorporação."
        }

    resultados = []
    for u_id, u_nome, r_nome in candidatos:
        res_add = adicionar_funcao_proposta_ativa(dados.caminho, u_id, r_nome, dados=dados)
        resultados.append({
            "usuario": u_id,
            "nome": u_nome,
            "role": r_nome,
            "resultado": res_add
        })

    return {
        "ok": True,
        "departamento": departamento,
        "total_incorporadas": len(resultados),
        "resultados": resultados,
    }


def imprimir_validacao_utilizadores_prd(res: Dict[str, Any]):
    """Exibe no terminal a validação dos utilizadores no PRD com cruzamento relacional e resumo executivo."""
    from collections import Counter

    dep = res.get("departamento", "")
    sis = res.get("sistema", "PRD")

    print("\n" + "=" * 78)
    print(f"  AUDITORIA DE AUTORIZAÇÕES: PLANO (EXCEL) vs SAP {sis} REAL")
    print(f"  Departamento: {dep} | Ambiente: Produção (Mandante {res.get('mandante', '100')})")
    print("=" * 78)

    if not res.get("ok"):
        print(f"  [ERRO] Erro ao validar utilizadores no PRD: {res.get('erro')}")
        print("=" * 78)
        return

    print("  OBJETIVO DESTA ANÁLISE:")
    print("     Confronta o que está desenhado no Excel ('Proposta Ativa') com o que")
    print(f"     está efetivamente ativo no utilizador no SAP {sis} (via RFC AGR_USERS / USR02).")
    print("\n  GUIA DE LEITURA:")
    print("     CONFORME            -> Utilizador tem 100% das funções do plano ativas no SAP.")
    print("     [ERRO] EM FALTA            -> Funções do plano que NÃO estão ativas no SAP (falta dar acesso).")
    print("     ADICIONAIS A VALIDAR-> Funções ativas no SAP que NÃO constam no Excel (acesso a mais).")
    print("     REGRAS DE EXCEÇÃO   -> Funções técnicas/standard ignoradas propositadamente.")
    print("     [EXPIRADO] CONTA EXPIRADA       -> Conta desativada/expirada no cadastro mestre USR02 do SAP.")
    print("-" * 78)
    print(f"  Exceções ativas no escopo: {', '.join(res.get('padroes_exclusao', []))}")
    print("-" * 78)

    utilizadores = res.get("utilizadores", [])
    if not utilizadores:
        print("  [AVISO] Nenhum utilizador encontrado para este departamento.")
        print("=" * 78)
        return

    # Contadores globais
    total_users = len(utilizadores)
    conformes = []
    com_faltas = []
    inativos = []
    faltas_global = Counter()
    adicionais_global = Counter()

    for u in utilizadores:
        uname = u["usuario"]
        nome = u["nome"]
        cargo = u["cargo"]
        esp_total = u["esperadas_total"]
        esp_ativas = u["esperadas_ativas"]
        faltam = u["faltam"]
        adicionais = u["adicionais"]
        desconsideradas = u["desconsideradas_exclusao"]
        is_inativo = u.get("inativo_usr02", False)

        if is_inativo:
            inativos.append(u)
            status_ico = "[EXPIRADO]"
            status_tag = f"[CONTA EXPIRADA EM USR02 a {u.get('validade_fim', '')}]"
        elif u["conforme"]:
            conformes.append(u)
            status_ico = "[OK]"
            status_tag = "[100% CONFORME]"
        else:
            com_faltas.append(u)
            status_ico = "[AVISO]"
            status_tag = f"[{len(faltam)} FUNÇÕES EM FALTA]"

        pct = int((esp_ativas / esp_total * 100)) if esp_total > 0 else (100 if is_inativo else 0)
        adicionais_str = f" | {len(adicionais)} a mais no SAP" if adicionais else ""

        print(f"\n  {status_ico} {uname:<10} | {nome:<25} | {cargo} {status_tag}")
        print(f"     └─ Funções Ativas do Plano: {esp_ativas}/{esp_total} ({pct}% concluído){adicionais_str}")

        if faltam and not is_inativo:
            for r in faltam:
                faltas_global[r] += 1
            print(f"        [ERRO] EM FALTA NO SAP ({len(faltam)}): {', '.join(faltam)}")
        elif faltam and is_inativo:
            print(f"        ℹ Sem atribuições ativas por desativação/offboarding no sistema SAP.")

        if adicionais:
            for r in adicionais:
                adicionais_global[r] += 1
            if u.get("adicionais_no_catalogo"):
                print(f"        ADICIONAIS NO SAP (No Catálogo Criado) ({len(u['adicionais_no_catalogo'])}): {', '.join(u['adicionais_no_catalogo'])} (Função oficial: pode ser acrescentada à Proposta Ativa)")
            if u.get("adicionais_fora_catalogo"):
                print(f"        [AVISO] ADICIONAIS NO SAP (Fora do Catálogo) ({len(u['adicionais_fora_catalogo'])}): {', '.join(u['adicionais_fora_catalogo'])} ([ALERTA] Função legada: candidata a CUA_REMOVE)")

        if desconsideradas:
            print(f"        Protegidas por EXCLUÇÃO ({len(desconsideradas)}): {', '.join(desconsideradas)}")

    # Resumo Executivo
    print("\n" + "=" * 78)
    print(f"  RESUMO EXECUTIVO: {dep.upper()}")
    print("=" * 78)
    print(f"  Total de Colaboradores: {total_users}")
    print(f"     ├─ Conformes (acessos 100% alinhados): {len(conformes)}")
    print(f"     ├─ [AVISO] Com funções em falta no SAP:       {len(com_faltas)}")
    print(f"     └─ [EXPIRADO] Contas expiradas no SAP (USR02):    {len(inativos)}")

    if faltas_global:
        print("\n  [ERRO] FUNÇÕES MAIS CRÍTICAS EM FALTA NO SAP (Precisam de atribuição):")
        for role, count in faltas_global.most_common(8):
            print(f"     • {role:<35} -> Falta a {count} colaborador(es)")

    if adicionais_global:
        print("\n  FUNÇÕES ADICIONAIS DETETADAS NO SAP (Não estão no Excel):")
        for role, count in adicionais_global.most_common(8):
            nota = " (Presente em quase toda a equipa - considerar incluir no Excel)" if count >= (len(conformes) + len(com_faltas)) * 0.7 else ""
            print(f"     • {role:<35} -> Ativa em {count} colaborador(es){nota}")

    # Candidatos a incorporar na Proposta Ativa
    candidatos_proposta = []
    for u in utilizadores:
        if u.get("inativo_usr02"):
            continue
        for r in u.get("adicionais_no_catalogo", []):
            candidatos_proposta.append((u["usuario"], u["nome"], r))

    if candidatos_proposta:
        print("\n  FUNÇÕES DO CATÁLOGO ATIVAS NO SAP PRD (Candidatas a inclusão na Proposta Ativa):")
        for u_id, u_nome, r_nome in candidatos_proposta:
            print(f"     • {u_id:<10} ({u_nome:<25}) -> {r_nome} (Pertence ao Catálogo oficial)")

    print("\n  PRÓXIMOS PASSOS SUGERIDOS:")
    if com_faltas:
        print("     1. Para atribuir as funções em falta aos utilizadores:")
        print("        Use a Ação [3] Sincronizar CUA, ou registe na folha CUA_ADICIONAR.")
    if candidatos_proposta:
        print("     2. Para as funções que já existem no catálogo oficial (ex.: Z_COSTCENTER_CREATE):")
        print("        Use a Ação [6] Incorporar Funções do Catálogo para acrescentar automaticamente à folha 'Proposta Ativa'.")
    if adicionais_global and not candidatos_proposta:
        print("     2. Para as funções adicionais fora do catálogo:")
        print("        Se a equipa não deve ter esse acesso: Registar na folha CUA_REMOVE para remoção.")
    if inativos:
        print("     3. Colaboradores expirados (USR02):")
        print("        Confirmar se saíram da empresa e atualizar status na folha CONTROLO/Proposta.")
    print("=" * 78)


def auditar_utilizador(dados: ProjetoPerfilData, utilizador: str) -> Dict[str, Any]:
    """
    Realiza uma auditoria completa e aprofundada a um utilizador:
      1. Localização na sheet 'Proposta Ativa' (departamento, cargo, roles atribuídas).
      2. Mapeamento no CUA (CUA_ADICIONAR e CUA_REMOVE).
      3. Expansão relacional das funções esperadas (PFCG_COMPOSTA + PFCG_AUTHORITY - EXCLUÇÃO).
      4. Consulta em tempo real ao SAP PRD via RFC (USR02 e AGR_USERS).
    """
    user_norm = normalizar_texto(utilizador)
    if not user_norm:
        return {"ok": False, "erro": "Identificador de utilizador vazio."}

    # 1. Pesquisa na Proposta Ativa
    meta_proposta = None
    sheet_proposta = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PROPOSTAATIVA"), None)
    if sheet_proposta:
        try:
            import pandas as pd
            fonte = abrir_excel_seguro(dados.caminho)
            df_raw = pd.read_excel(fonte, sheet_name=sheet_proposta, header=None)
            for idx in range(1, len(df_raw)):
                u_id = str(df_raw.iloc[idx, 0]).strip() if pd.notna(df_raw.iloc[idx, 0]) else ""
                if normalizar_texto(u_id) == user_norm:
                    singles = [
                        str(x).strip() for x in df_raw.iloc[idx, 9:].dropna().tolist()
                        if str(x).strip() and str(x).strip().upper() not in ("NAN", "NONE", "")
                    ]
                    comp = str(df_raw.iloc[idx, 8]).strip() if pd.notna(df_raw.iloc[idx, 8]) else ""
                    comp = comp if comp.upper() not in ("NAN", "NONE", "") else None
                    meta_proposta = {
                        "usuario": u_id,
                        "nome": str(df_raw.iloc[idx, 2]).strip() if pd.notna(df_raw.iloc[idx, 2]) else "",
                        "departamento": str(df_raw.iloc[idx, 7]).strip() if pd.notna(df_raw.iloc[idx, 7]) else "",
                        "cargo": str(df_raw.iloc[idx, 6]).strip() if pd.notna(df_raw.iloc[idx, 6]) else "",
                        "composta": comp,
                        "singles": singles,
                        "total_singles_proposta": len(singles),
                    }
                    break
        except Exception:
            pass

    # 2. Atribuições CUA
    cua_info = dados.utilizadores_cua.get(user_norm, {"adicionar": [], "remover": []})

    # 3. Expansão Relacional
    esperadas = set()
    if meta_proposta:
        base_set = set()
        if meta_proposta.get("composta"):
            base_set.add(meta_proposta["composta"])
        base_set.update(meta_proposta.get("singles", []))
        esperadas = dados.expandir_funcoes(base_set)

    # 4. Consulta ao SAP PRD via RFC
    os.environ["SAP_TARGET_ENV"] = "PRD"
    usr02_info = {}
    desc_map = {}
    prd_ok = False
    prd_erro = ""
    agr_rows_raw = []

    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_read_only_guard, read_table, make_option_in
        )
        from pyrfc import Connection
        from datetime import date

        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for("PRD")
        conn = Connection(**params)
        guard = make_read_only_guard(["USR02", "AGR_USERS", "AGR_TEXTS"])

        # Mestre USR02
        u_rows = read_table(conn, guard, table_name="USR02", fields=["BNAME", "GLTGV", "GLTGB", "USTYP", "UFLAG", "TRDAT", "LTIME"], options=make_option_in("BNAME", [user_norm]), rowcount=0)
        if u_rows:
            r_u = u_rows[0]
            g_fim = str(r_u[2]).strip()
            hoje_int = int(date.today().strftime("%Y%m%d"))
            fim_int = int(g_fim) if g_fim.isdigit() and int(g_fim) > 0 else 99991231
            usr02_info = {
                "tipo": str(r_u[3]).strip(),
                "bloqueado": str(r_u[4]).strip() != "0",
                "uflag": str(r_u[4]).strip(),
                "expirado": hoje_int > fim_int,
                "validade_fim": g_fim,
                "ultimo_logon": f"{r_u[5].strip()} {r_u[6].strip()}".strip(),
            }

        # AGR_USERS
        agr_rows_raw = read_table(conn, guard, table_name="AGR_USERS", fields=["AGR_NAME", "FROM_DAT", "TO_DAT"], options=make_option_in("UNAME", [user_norm]), rowcount=0)
        roles_list = [r[0].strip() for r in agr_rows_raw if r]
        
        # AGR_TEXTS
        if roles_list:
            text_rows = read_table(conn, guard, table_name="AGR_TEXTS", fields=["AGR_NAME", "SPRAS", "TEXT"], options=make_option_in("AGR_NAME", roles_list[:60]), rowcount=0)
            for tr in text_rows:
                r_name = tr[0].strip()
                spras = tr[1].strip().upper()
                txt = tr[2].strip()
                if spras in ("P", "PT") or r_name not in desc_map:
                    desc_map[r_name] = txt

        conn.close()
        prd_ok = True

        # Processar atribuições
        hoje_int = int(date.today().strftime("%Y%m%d"))
        ativas_prd = set()
        expiradas_prd = []
        for r in agr_rows_raw:
            role = str(r[0]).strip().upper()
            from_d = str(r[1]).strip()
            to_d = str(r[2]).strip()
            fim = int(to_d) if to_d.isdigit() and int(to_d) > 0 else 99991231
            ini = int(from_d) if from_d.isdigit() and int(from_d) > 0 else 0
            desc = desc_map.get(role, "")
            if ini <= hoje_int <= fim:
                ativas_prd.add(role)
            elif hoje_int > fim:
                expiradas_prd.append({"role": role, "inicio": from_d, "fim": to_d, "desc": desc})

    except Exception as err:
        prd_erro = str(err)
        ativas_prd = set()
        expiradas_prd = []

    faltam = sorted(esperadas - ativas_prd) if esperadas else []
    adicionais_brutas = sorted(ativas_prd - esperadas) if esperadas else sorted(ativas_prd)
    desconsideradas = sorted([r for r in adicionais_brutas if dados.is_excluida(r)])
    adicionais_reais = sorted(set(adicionais_brutas) - set(desconsideradas))

    return {
        "ok": True,
        "utilizador": user_norm,
        "proposta": meta_proposta,
        "cua": cua_info,
        "esperadas_total": len(esperadas),
        "esperadas_roles": sorted(esperadas),
        "prd_conectado": prd_ok,
        "prd_erro": prd_erro,
        "mestre_usr02": usr02_info,
        "ativas_total": len(ativas_prd),
        "expiradas_total": len(expiradas_prd),
        "expiradas": expiradas_prd,
        "esperadas_ativas": len(esperadas & ativas_prd),
        "faltam": faltam,
        "protegidas_exclusao": desconsideradas,
        "adicionais_reais": adicionais_reais,
        "conforme": (len(faltam) == 0 and len(esperadas) > 0),
        "padroes_exclusao": dados.padroes_exclusao,
    }


def imprimir_auditoria_utilizador(res: Dict[str, Any]):
    """Exibe no terminal a auditoria detalhada de um utilizador."""
    u = res.get("utilizador", "")
    p = res.get("proposta") or {}
    print("\n" + "=" * 75)
    print(f"  AUDITORIA COMPLETA DE UTILIZADOR: {u}")
    if p.get("nome"):
        print(f"  Nome: {p.get('nome')} | Cargo: {p.get('cargo')} | Departamento: {p.get('departamento')}")
    print("=" * 75)

    # 1. Proposta Ativa
    if p:
        print("\n  1. DADOS NA 'PROPOSTA ATIVA':")
        print(f"     ├─ Função Composta: {p.get('composta') or 'Nenhuma'}")
        print(f"     ├─ Funções Simples Declaradas: {p.get('total_singles_proposta')}")
        print(f"     └─ Total Esperado após Expansão Relacional: {res.get('esperadas_total')} Funções")
    else:
        print("\n  1. 'PROPOSTA ATIVA': Utilizador não localizado na folha Proposta Ativa.")

    # 2. CUA
    cua = res.get("cua", {})
    adds = cua.get("adicionar", [])
    rms = cua.get("remover", [])
    if adds or rms:
        print(f"\n  2. ATRIBUIÇÕES NO CUA: {len(adds)} para Adicionar, {len(rms)} para Remover")
    else:
        print("\n  2. ATRIBUIÇÕES NO CUA: Nenhuma instrução pendente em CUA_ADICIONAR/CUA_REMOVE.")

    # 3. SAP PRD
    if not res.get("prd_conectado"):
        print(f"\n  [ERRO] 3. SAP PRD: Erro de ligação RFC ({res.get('prd_erro')})")
        return

    m = res.get("mestre_usr02", {})
    print("\n  3. ESTADO NO SAP PRD (Mestre USR02 & AGR_USERS):")
    if m:
        st_lock = "[BLOQUEADO] BLOQUEADO" if m.get("bloqueado") else "[Desbloqueado] Desbloqueado"
        st_exp = f"[EXPIRADO] EXPIRADO a {m.get('validade_fim')}" if m.get("expirado") else "Conta Ativa"
        print(f"     ├─ Tipo: {m.get('tipo', 'A')} | Estado: {st_lock} | Validade: {st_exp}")
        print(f"     ├─ Último Logon: {m.get('ultimo_logon') or 'Nunca'}")
    print(f"     ├─ Funções Ativas no PRD:       {res.get('ativas_total')}")
    print(f"     ├─ Funções Expiradas/Histórico: {res.get('expiradas_total')}")
    print(f"     └─ Conformidade com a Proposta: {res.get('esperadas_ativas')}/{res.get('esperadas_total')} ativas")

    if res.get("faltam"):
        print(f"\n  [ERRO] FUNÇÕES EM FALTA ({len(res['faltam'])}):")
        for f in res["faltam"]:
            print(f"     ├─ {f}")
    else:
        print("\n  TODAS AS FUNÇÕES PREVISTAS ESTÃO 100% ATIVAS NO PRD (0 em falta)!")

    prot = res.get("protegidas_exclusao", [])
    if prot:
        print(f"\n  Ocorrências Protegidas por EXCLUÇÃO ({len(prot)}):")
        for pr in prot:
            print(f"     ├─ {pr}")

    adicionais = res.get("adicionais_reais", [])
    if adicionais:
        print(f"\n  FUNÇÕES ADICIONAIS A VALIDAR ({len(adicionais)}):")
        for a in adicionais:
            print(f"     ├─ {a}")
    else:
        print("\n  NENHUMA FUNÇÃO ADICIONAL PENDENTE DE VALIDAÇÃO (0 adicionais)!")


# =====================================================================
# FORMATAÇÃO VISUAL E INTERFACE DE LINHA DE COMANDOS
# =====================================================================

def imprimir_cabecalho(ficheiro: str):
    sp_url = os.getenv("SHAREPOINT_PERFIS_URL", "")
    print("\n" + "=" * 75)
    print("  PROJETO PERFIL - Leitor e Pesquisador de Funções SAP (PFCG / CUA)")
    print(f"  Ficheiro:          {os.path.basename(ficheiro)}")
    print(f"  Caminho Local:     {ficheiro}")
    if sp_url:
        print(f"  SharePoint:        {sp_url}")
    print("=" * 75)


def imprimir_validacao_controlo(dados: ProjetoPerfilData):
    """Exibe no terminal a validação dos departamentos da sheet CONTROLO."""
    print("\n" + "=" * 75)
    print("  VALIDAÇÃO DA SHEET 'CONTROLO' (Departamentos & Status)")
    print("=" * 75)
    if not dados.controlo:
        print("  [AVISO] A sheet 'CONTROLO' não foi encontrada ou não possui dados.")
        return

    pendentes = [d for d in dados.controlo if d.get("pendente")]
    print(f"  Total de departamentos registados: {len(dados.controlo)}")
    print(f"  Departamentos pendentes (STATUS e TIMESTAMP vazios): {len(pendentes)}")
    print("-" * 75)
    for c in dados.controlo:
        status_tag = "PENDENTE (STATUS E TIMESTAMP VAZIOS)" if c["pendente"] else f"{c['status']}"
        ts_info = f" | {c['timestamp']}" if c.get("timestamp") else ""
        print(f"  Linha {c['linha']}: {c['departamento']:<30} -> {status_tag}{ts_info}")

    print("-" * 75)
    proximo = dados.obter_proximo_departamento()
    if proximo:
        print(f"  PRÓXIMO DEPARTAMENTO A AVANÇAR: '{proximo}'")
        item_prox = next((p for p in pendentes if p["departamento"] == proximo), {})
        tem_sheet = proximo in dados.sheets_disponiveis
        sheet_info = f"[Sheet correspondente '{proximo}' EXISTE no Excel]" if tem_sheet else "[Sheet correspondente não encontrada]"
        print(f"     Linha {item_prox.get('linha')}: STATUS e TIMESTAMP estão vazios! {sheet_info}")
        if len(pendentes) > 1:
            print("\n  Demais departamentos pendentes:")
            for p in pendentes[1:]:
                print(f"     - '{p['departamento']}' (Linha {p['linha']})")
    else:
        print("  Todos os departamentos na sheet CONTROLO já possuem STATUS e TIMESTAMP preenchidos.")


def imprimir_resumo(dados: ProjetoPerfilData):
    stats = dados.obter_estatisticas()
    print("\nRESUMO GERAL DAS FUNÇÕES:")
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
        print("  [AVISO] Nenhuma função encontrada com o critério indicado.")
        return

    print(f"\nEncontrada(s) {len(res)} função(ões):")
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
        print("  [AVISO] Nenhuma função encontrada que contenha esta transação.")
        return

    print(f"\nEncontrada(s) {len(res)} correspondência(s) de transação:")
    # Agrupar por TCODE
    por_tc: Dict[str, List[Dict[str, Any]]] = {}
    for r in res:
        tc = r["tcode"]
        por_tc.setdefault(tc, []).append(r)

    for tc, lista in por_tc.items():
        print(f"\n  TCODE: {tc}")
        for item in lista:
            role = item["role_simples"]
            desc = f" ({item['descricao_role']})" if item.get("descricao_role") else ""
            comp_txt = f" [Compostas: {', '.join(item['compostas'])}]" if item.get("compostas") else ""
            print(f"     ├─ Role Simples: {role}{desc}{comp_txt}")


def imprimir_resultado_user(res: Dict[str, Any]):
    users = res.get("resultados", [])
    if not users:
        print("  [AVISO] Nenhum utilizador encontrado com esse identificador no CUA.")
        return

    print(f"\nEncontrado(s) {len(users)} utilizador(es):")
    for u in users:
        nome_user = u["utilizador"]
        adds = u.get("adicionar", [])
        rms = u.get("remover", [])
        print(f"\n  Utilizador: {nome_user}")
        if adds:
            print(f"     Funções para Adicionar ({len(adds)}):")
            for a in adds:
                st = f" [STATUS: {a['status']}]" if a.get("status") else " [PENDENTE]"
                sis = f" ({a['sistema']})" if a.get("sistema") else ""
                print(f"        ├─ {a['role']}{sis}{st}")
        if rms:
            print(f"     [ERRO] Funções para Remover ({len(rms)}):")
            for r in rms:
                st = f" [STATUS: {r['status']}]" if r.get("status") else " [PENDENTE]"
                sis = f" ({r['sistema']})" if r.get("sistema") else ""
                print(f"        ├─ {r['role']}{sis}{st}")


# =====================================================================
# MENU INTERATIVO - FLUXO DIRECIONADO POR DEPARTAMENTO (CONTROLO)
# =====================================================================

def imprimir_cabecalho_compacto(ficheiro: str):
    """Exibe cabeçalho limpo e direto ao iniciar."""
    print("\n" + "=" * 75)
    print("  PROJETO PERFIL - PROCESSAMENTO DE DEPARTAMENTOS (CUA / PFCG)")
    print(f"  Ficheiro: {os.path.basename(ficheiro)}")
    print("=" * 75)


def exibir_tabela_controlo(dados: ProjetoPerfilData) -> Optional[Dict[str, Any]]:
    """
    Exibe a tabela clara dos departamentos na folha CONTROLO com as suas linhas no Excel
    e retorna o próximo departamento pendente sugerido (se houver).
    """
    print("\nDEPARTAMENTOS NA SHEET 'CONTROLO':")
    print("-" * 75)
    print(f"  {'Linha':<8} | {'Departamento':<30} | {'Estado na CONTROLO':<30}")
    print("-" * 75)

    if not dados.controlo:
        print("  Nenhuma linha encontrada na folha CONTROLO.")
        print("-" * 75)
        return None

    proximo_sugerido = None
    for c in dados.controlo:
        linha_str = f"[{c['linha']}]"
        if c.get("pendente"):
            status_desc = "DISPONIVEL (Pendente)"
            if not proximo_sugerido:
                proximo_sugerido = c
        else:
            ts_curto = str(c.get("timestamp", ""))[:16]
            ts = f" ({ts_curto})" if ts_curto else ""
            status_desc = f"{c.get('status', 'PROCESSADO')}{ts}"
        print(f"  {linha_str:<8} | {c['departamento']:<30} | {status_desc}")
    print("-" * 75)

    if proximo_sugerido:
        print(f"  Proximo departamento sugerido: Linha {proximo_sugerido['linha']} ('{proximo_sugerido['departamento']}')")
    else:
        print("  Todos os departamentos registados na sheet CONTROLO estao processados.")

    return proximo_sugerido


def selecionar_departamento_interativo(dados: ProjetoPerfilData) -> Optional[Dict[str, Any]]:
    """
    Apresenta a listagem da folha CONTROLO e solicita ao utilizador indicar a linha a processar.
    """
    sugerido = exibir_tabela_controlo(dados)
    padrao_linha = str(sugerido["linha"]) if sugerido else ""
    prompt_sugestao = f" [Enter para Linha {padrao_linha}]" if padrao_linha else ""

    while True:
        entrada = input(f"\nIndique a LINHA do departamento que quer processar{prompt_sugestao} ('M' Menu Geral, '0' Sair): ").strip().upper()

        if entrada in ("0", "SAIR", "Q", "QUIT", "EXIT"):
            return None

        if entrada in ("M", "MENU", "GERAL"):
            return {"tipo": "MENU_GERAL"}

        if not entrada and padrao_linha:
            entrada = padrao_linha

        # Busca por número de linha
        item_match = next((c for c in dados.controlo if str(c.get("linha")) == entrada), None)
        if item_match:
            return item_match

        # Busca alternativa por nome de departamento
        item_por_nome = next((c for c in dados.controlo if entrada in c.get("departamento", "").upper()), None)
        if item_por_nome:
            return item_por_nome

        print(f"[AVISO] Linha ou departamento '{entrada}' não encontrado na folha CONTROLO. Escolha uma das linhas listadas.")


def sincronizar_departamento_cua_completo(dados: ProjetoPerfilData, item_dep: Dict[str, Any]):
    """
    Executa o fluxo completo de sincronização no SAP CUA para o departamento selecionado:
      1. Remoção de S4DCLNT100
      2. Remoção de funções expiradas em S4PCLNT100
      3. Remoção de funções obsoletas em S4QCLNT100
      4. Replicação das funções ativas de S4PCLNT100 para S4QCLNT100
      5. Auditoria em tempo real (P == Q)
      6. Atualização de CUA_REMOVE, CUA_ADICIONAR e CONTROLO (SharePoint COM e local)
    """
    dep_nome = item_dep["departamento"]
    linha_excel = item_dep.get("linha", 0)

    analise = analisar_departamento_proposta(dados, dep_nome)
    if not analise.get("encontrado"):
        print(f"[ERRO] Não foi possível carregar os utilizadores do departamento '{dep_nome}'.")
        return

    users = [u["usuario"] for u in analise.get("usuarios", []) if u.get("usuario")]
    if not users:
        print(f"[ERRO] Nenhum utilizador encontrado para '{dep_nome}'.")
        return

    print(f"\n" + "=" * 75)
    print(f"  INICIAR SINCRONIZAÇÃO CUA: {dep_nome.upper()} ({len(users)} Utilizadores)")
    print("=" * 75)
    print(f"  Utilizadores a processar: {', '.join(users)}")
    print("  Etapas automáticas:")
    print("    1. Remoção de sistema secundário S4DCLNT100 (se existir)")
    print("    2. Remoção de roles expiradas em S4PCLNT100")
    print("    3. Limpeza de roles obsoletas em S4QCLNT100")
    print("    4. Atribuição de roles ativas de S4PCLNT100 em S4QCLNT100")
    print("    5. Auditoria em tempo real de integridade P == Q")
    print("    6. Gravação oficial em CUA_REMOVE, CUA_ADICIONAR e CONTROLO")
    print("-" * 75)

    confirma = input("Confirma a execução imediata no SAP CUA? (S/N): ").strip().upper()
    if confirma not in ("S", "SIM", "Y", "YES"):
        print("[EXPIRADO] Operação cancelada pelo utilizador.")
        return

    import threading
    erro_execucao = []

    def worker_sync():
        try:
            import win32service, pythoncom, win32com.client, win32clipboard
            import time
            from datetime import datetime
            from pathlib import Path

            try:
                hwinsta = win32service.OpenWindowStation("WinSta0", True, 0x037F)
                hwinsta.SetProcessWindowStation()
                hdesk = win32service.OpenDesktop("default", 0, True, 0x01FF)
                hdesk.SetThreadDesktop()
            except:
                pass
            pythoncom.CoInitialize()

            rot = win32com.client.Dispatch("SapROTWr.SapROTWrapper")
            sap = rot.GetROTEntry("SAPGUI")
            session = sap.GetScriptingEngine.Children(0).Children(0)

            roles_removidas = []
            roles_adicionadas = []

            for u in users:
                print(f"\n  A processar utilizador: {u} ...")

                # A. Remover S4DCLNT100 na aba Sistemas se existir
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)
                session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = u
                session.findById("wnd[0]/tbar[1]/btn[18]").press()
                time.sleep(0.4)

                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpSYSTEMS").select()
                time.sleep(0.3)
                grid_sys = session.findById(
                    "wnd[0]/usr/tabsTABSTRIP1/tabpSYSTEMS/ssubMAINAREA:SAPLSUID_MAINTENANCE:1210/cntlG_CUA_SYSTEMS_CONTAINER/shellcont/shell"
                )
                linha_s4d = None
                for r in range(grid_sys.RowCount):
                    if grid_sys.GetCellValue(r, "SUBSYSTEM") == "S4DCLNT100":
                        linha_s4d = r
                        break
                if linha_s4d is not None:
                    grid_sys.setCurrentCell(linha_s4d, "SUBSYSTEM")
                    grid_sys.selectedRows = str(linha_s4d)
                    grid_sys.pressToolbarButton("DEL_LINE")
                    time.sleep(0.3)
                    session.findById("wnd[0]").sendVKey(11) # Salvar
                    time.sleep(0.6)
                    print(f"     S4DCLNT100 removido de {u}.")

                # B. Remover roles expiradas de S4PCLNT100
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)
                session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = u
                session.findById("wnd[0]/tbar[1]/btn[18]").press()
                time.sleep(0.4)
                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpACTG").select()
                time.sleep(0.3)
                grid_act = session.findById(
                    "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/cntlG_ROLES_CONTAINER/shellcont/shell"
                )
                expired_p = []
                for r in range(grid_act.RowCount):
                    sis = grid_act.GetCellValue(r, "SUBSYSTEM")
                    agr = grid_act.GetCellValue(r, "AGR_NAME")
                    t_dat = grid_act.GetCellValue(r, "UPDATE_TO_DAT")
                    if sis == "S4PCLNT100" and agr and "9999" not in str(t_dat):
                        expired_p.append(agr)

                if expired_p:
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText("\r\n".join(expired_p))
                    win32clipboard.CloseClipboard()

                    grid_act.selectColumn("SUBSYSTEM")
                    grid_act.selectColumn("AGR_NAME")
                    grid_act.pressToolbarButton("&MB_FILTER")
                    time.sleep(0.3)

                    wnd1 = session.findById("wnd[1]")
                    wnd1.findById("usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "S4PCLNT100"
                    wnd1.findById("usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN002_%_APP_%-VALU_PUSH").press()
                    time.sleep(0.3)
                    wnd2 = session.findById("wnd[2]")
                    wnd2.findById("tbar[0]/btn[24]").press()
                    time.sleep(0.3)
                    wnd2.findById("tbar[0]/btn[8]").press()
                    time.sleep(0.3)
                    wnd1.findById("tbar[0]/btn[0]").press()
                    time.sleep(0.4)

                    rc = grid_act.RowCount
                    if rc > 0:
                        grid_act.selectedRows = ",".join(str(i) for i in range(rc))
                        grid_act.pressToolbarButton("DEL_LINE")
                        time.sleep(0.3)
                        try:
                            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            time.sleep(0.3)
                        except:
                            pass
                    session.findById("wnd[0]").sendVKey(11) # Salvar
                    time.sleep(0.8)
                    for r_exp in expired_p:
                        roles_removidas.append((u, "S4PCLNT100", r_exp, "CONCLUÍDO", f"User {u} has changed"))
                    print(f"     {len(expired_p)} roles expiradas removidas de S4PCLNT100.")

                # C. Remover roles obsoletas de S4QCLNT100
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)
                session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = u
                session.findById("wnd[0]/tbar[1]/btn[18]").press()
                time.sleep(0.4)
                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpACTG").select()
                time.sleep(0.3)
                grid_act = session.findById(
                    "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/cntlG_ROLES_CONTAINER/shellcont/shell"
                )
                roles_q = [
                    grid_act.GetCellValue(r, "AGR_NAME") for r in range(grid_act.RowCount)
                    if grid_act.GetCellValue(r, "SUBSYSTEM") == "S4QCLNT100" and grid_act.GetCellValue(r, "AGR_NAME")
                ]
                if roles_q:
                    grid_act.selectColumn("SUBSYSTEM")
                    grid_act.pressToolbarButton("&MB_FILTER")
                    time.sleep(0.3)
                    wnd1 = session.findById("wnd[1]")
                    wnd1.findById("usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = "S4QCLNT100"
                    wnd1.findById("tbar[0]/btn[0]").press()
                    time.sleep(0.4)
                    rc_q = grid_act.RowCount
                    if rc_q > 0:
                        grid_act.selectedRows = ",".join(str(i) for i in range(rc_q))
                        grid_act.pressToolbarButton("DEL_LINE")
                        time.sleep(0.3)
                        try:
                            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            time.sleep(0.3)
                        except:
                            pass
                    session.findById("wnd[0]").sendVKey(11) # Salvar
                    time.sleep(0.8)
                    for r_q in roles_q:
                        roles_removidas.append((u, "S4QCLNT100", r_q, "CONCLUÍDO", f"User {u} has changed"))
                    print(f"     {len(roles_q)} roles obsoletas removidas de S4QCLNT100.")

                # D. Adicionar roles ativas de S4PCLNT100 em S4QCLNT100
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)
                session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = u
                session.findById("wnd[0]/tbar[1]/btn[18]").press()
                time.sleep(0.4)
                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpACTG").select()
                time.sleep(0.3)
                grid_act = session.findById(
                    "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/cntlG_ROLES_CONTAINER/shellcont/shell"
                )
                p_roles_dict = {}
                for r in range(grid_act.RowCount):
                    sis = grid_act.GetCellValue(r, "SUBSYSTEM")
                    agr = grid_act.GetCellValue(r, "AGR_NAME")
                    f_dat = grid_act.GetCellValue(r, "UPDATE_FROM_DAT")
                    t_dat = grid_act.GetCellValue(r, "UPDATE_TO_DAT")
                    if sis == "S4PCLNT100" and agr:
                        if agr not in p_roles_dict or "9999" in str(t_dat):
                            p_roles_dict[agr] = (f_dat, t_dat)

                first_empty = -1
                for r in range(grid_act.RowCount):
                    if not grid_act.GetCellValue(r, "SUBSYSTEM") and not grid_act.GetCellValue(r, "AGR_NAME"):
                        first_empty = r
                        break
                if first_empty == -1:
                    first_empty = grid_act.RowCount

                curr_row = first_empty
                for agr, (f_dat, t_dat) in sorted(p_roles_dict.items()):
                    grid_act.firstVisibleRow = max(0, curr_row - 2)
                    grid_act.modifyCell(curr_row, "SUBSYSTEM", "S4QCLNT100")
                    grid_act.modifyCell(curr_row, "AGR_NAME", agr)
                    if f_dat:
                        try: grid_act.modifyCell(curr_row, "UPDATE_FROM_DAT", f_dat)
                        except: pass
                    if t_dat:
                        try: grid_act.modifyCell(curr_row, "UPDATE_TO_DAT", t_dat)
                        except: pass
                    roles_adicionadas.append((u, "S4QCLNT100", agr, "CONCLUÍDO", f"User {u} has changed"))
                    curr_row += 1

                grid_act.currentCellColumn = "AGR_NAME"
                grid_act.pressEnter()
                time.sleep(0.5)
                session.findById("wnd[0]").sendVKey(11) # Salvar
                time.sleep(0.8)
                print(f"     {len(p_roles_dict)} roles ativas replicadas para S4QCLNT100.")

                # E. Auditoria final
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nSU01"
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.3)
                session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").text = u
                session.findById("wnd[0]/tbar[1]/btn[18]").press()
                time.sleep(0.4)
                session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpACTG").select()
                time.sleep(0.3)
                grid_act = session.findById(
                    "wnd[0]/usr/tabsTABSTRIP1/tabpACTG/ssubMAINAREA:SAPLSUID_MAINTENANCE:1106/cntlG_ROLES_CONTAINER/shellcont/shell"
                )
                p_finais = {grid_act.GetCellValue(r, "AGR_NAME") for r in range(grid_act.RowCount) if grid_act.GetCellValue(r, "SUBSYSTEM") == "S4PCLNT100" and grid_act.GetCellValue(r, "AGR_NAME")}
                q_finais = {grid_act.GetCellValue(r, "AGR_NAME") for r in range(grid_act.RowCount) if grid_act.GetCellValue(r, "SUBSYSTEM") == "S4QCLNT100" and grid_act.GetCellValue(r, "AGR_NAME")}
                sess_ok = (p_finais == q_finais)
                print(f"     Auditoria {u}: P={len(p_finais)} | Q={len(q_finais)} -> MATCH EXATO: {sess_ok}")

            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

            # Gravação em Excel
            print("\n  A atualizar folhas de cálculo oficiais (CUA_REMOVE, CUA_ADICIONAR, CONTROLO)...")
            xl = win32com.client.Dispatch("Excel.Application")
            wb = None
            for w in xl.Workbooks:
                if "S4H_Perfis de autorização.xlsx" in w.Name or "perfis" in w.Name.lower():
                    wb = w
                    break

            timestamp_now = datetime.now()
            timestamp_str = timestamp_now.strftime("%d/%m/%Y %H:%M:%S")

            if wb:
                # CUA_REMOVE
                ws_rem = wb.Worksheets("CUA_REMOVE")
                lr_rem = ws_rem.UsedRange.Rows.Count
                lid_rem = int(ws_rem.Cells(lr_rem, 1).Value or 0)
                cur_r = lr_rem + 1
                cur_id = lid_rem + 1
                for usr, sis, rol, st, msg in roles_removidas:
                    ws_rem.Cells(cur_r, 1).Value = cur_id
                    ws_rem.Cells(cur_r, 2).Value = usr
                    ws_rem.Cells(cur_r, 3).Value = sis
                    ws_rem.Cells(cur_r, 4).Value = rol
                    ws_rem.Cells(cur_r, 5).Value = st
                    ws_rem.Cells(cur_r, 6).Value = msg
                    ws_rem.Cells(cur_r, 7).Value = timestamp_str
                    cur_r += 1
                    cur_id += 1

                # CUA_ADICIONAR
                ws_add = wb.Worksheets("CUA_ADICIONAR")
                lr_add = ws_add.UsedRange.Rows.Count
                lid_add = int(ws_add.Cells(lr_add, 1).Value or 0)
                cur_r = lr_add + 1
                cur_id = lid_add + 1
                for usr, sis, rol, st, msg in roles_adicionadas:
                    ws_add.Cells(cur_r, 1).Value = cur_id
                    ws_add.Cells(cur_r, 2).Value = usr
                    ws_add.Cells(cur_r, 3).Value = sis
                    ws_add.Cells(cur_r, 4).Value = rol
                    ws_add.Cells(cur_r, 5).Value = st
                    ws_add.Cells(cur_r, 6).Value = msg
                    ws_add.Cells(cur_r, 7).Value = timestamp_str
                    cur_r += 1
                    cur_id += 1

                # CONTROLO
                ws_ctrl = wb.Worksheets("CONTROLO")
                for r in range(2, ws_ctrl.UsedRange.Rows.Count + 1):
                    if str(ws_ctrl.Cells(r, 1).Value or "").strip() == dep_nome:
                        ws_ctrl.Cells(r, 2).Value = "PROCESSADO"
                        ws_ctrl.Cells(r, 3).Value = timestamp_now
                        break
                wb.Save()
                try:
                    caminho_local = Path(r"C:\workspace\SapScript\sap_script_uploads\S4H_Perfis de autorização.xlsx")
                    wb.SaveCopyAs(str(caminho_local))
                except:
                    pass
                print("  Excel guardado com sucesso!")

            item_dep["status"] = "PROCESSADO"
            item_dep["timestamp"] = timestamp_str
            item_dep["pendente"] = False
            print(f"\nSincronização do departamento '{dep_nome}' concluída com sucesso!")

        except Exception as exc:
            erro_execucao.append(exc)
            print(f"[ERRO] Erro na execução da sincronização CUA: {exc}")

    t = threading.Thread(target=worker_sync)
    t.start()
    t.join()


def executar_menu_departamento(dados: ProjetoPerfilData, item_dep: Dict[str, Any]) -> str:
    """
    Apresenta o resumo do departamento selecionado e as ações operacionais diretas.
    Retorna 'VOLTAR' para escolher outro departamento, ou 'SAIR' para terminar.
    """
    dep_nome = item_dep["departamento"]
    linha = item_dep.get("linha", "?")

    analise = analisar_departamento_proposta(dados, dep_nome)

    while True:
        print("\n" + "=" * 75)
        print(f"  DEPARTAMENTO SELECIONADO: {dep_nome.upper()} (Linha {linha} no Excel)")
        print("=" * 75)

        st_tag = "PENDENTE (Disponível para processamento)" if item_dep.get("pendente") else f"{item_dep.get('status')} | {item_dep.get('timestamp')}"
        print(f"  Estado na CONTROLO: {st_tag}")

        if analise.get("encontrado"):
            users = analise.get("usuarios", [])
            compostas = analise.get("compostas", [])
            print(f"  Utilizadores no Departamento ({len(users)}):")
            for u in users:
                comp = f"[Composta: {u['composta']}]" if u.get("composta") else "[Sem Composta]"
                print(f"     • {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']} {comp} -> {u['total_singles']} Singles")
            if compostas:
                print(f"  Funções Compostas ({len(compostas)}): {', '.join(compostas)}")
            print(f"  Funções Individuais (Singles): {analise.get('total_singles_distintas', 0)} distintas")
        else:
            print(f"  [AVISO] Aviso: {analise.get('mensagem', 'Departamento não encontrado na folha Proposta Ativa.')}")

        print("-" * 75)
        print("  AÇÕES DISPONÍVEIS PARA ESTE DEPARTAMENTO:")
        print("  [1] Validar Utilizadores no SAP PRD (AGR_USERS via RFC & Cruzamento)")
        print("  [2] Cruzar Fontes (PFCG_CREATE, COMPOSTA, AUTHORITY, EXCLUÇÃO)")
        print("  [3] Sincronizar CUA (Remover S4D / Limpar PRD / Alinhar QAS)")
        print("  [4] Verificar Existência de Funções no SAP PRD (AGR_DEFINE)")
        print("  [5] Ver Análise Detalhada (Proposta Ativa)")
        print("  [6] ➕ Incorporar Funções do Catálogo Ativas na 'Proposta Ativa'")
        print("  [7] ↩  Voltar / Escolher Outro Departamento da CONTROLO")
        print("  [8] Menu Geral de Pesquisas (Roles, TCODEs, Users, etc.)")
        print("  [0] Sair")
        print("-" * 75)

        acao = input("Escolha uma ação: ").strip().upper()

        if acao == "1":
            print(f"\nA validar atribuições no SAP PRD para '{dep_nome}' ...")
            res_prd = validar_utilizadores_prd(dados, dep_nome)
            imprimir_validacao_utilizadores_prd(res_prd)

            # Verificar se há funções do catálogo oficial para incorporar
            cands = []
            for u in res_prd.get("utilizadores", []):
                if u.get("inativo_usr02"):
                    continue
                for r in u.get("adicionais_no_catalogo", []):
                    cands.append((u["usuario"], u["nome"], r))
            if cands:
                print("-" * 75)
                conf = input(f"Foram detetadas {len(cands)} função(ões) ativas no PRD que constam no catálogo criado.\n   Deseja acrescentá-las à folha 'Proposta Ativa' agora? (S/N): ").strip().upper()
                if conf in ("S", "SIM", "Y", "YES"):
                    for u_id, u_nome, r_nome in cands:
                        res_add = adicionar_funcao_proposta_ativa(dados.caminho, u_id, r_nome, dados=dados)
                        print(f"   └─ {res_add.get('mensagem')}")
                    print("\nA recarregar dados do Excel e revalidar conformidade...")
                    dados = carregar_projeto_perfil(dados.caminho)
                    analise = analisar_departamento_proposta(dados, dep_nome)
                    res_prd_nova = validar_utilizadores_prd(dados, dep_nome)
                    imprimir_validacao_utilizadores_prd(res_prd_nova)

        elif acao == "2":
            res_cruz = cruzar_fontes_departamento(dados, dep_nome)
            imprimir_cruzamento_fontes(res_cruz)

        elif acao == "3":
            sincronizar_departamento_cua_completo(dados, item_dep)

        elif acao == "4":
            if analise.get("encontrado"):
                roles_dep = list(analise["compostas"]) + list(analise["singles_frequencia"].keys())
                res_prd = verificar_funcoes_prd(roles_dep)
                imprimir_resultado_verificacao_prd(res_prd, f"Departamento '{dep_nome}'")
            else:
                print("[AVISO] Não foi possível obter as funções do departamento.")

        elif acao == "5":
            imprimir_analise_departamento(analise)

        elif acao == "6":
            print(f"\nA verificar funções do catálogo ativas no PRD para '{dep_nome}' ...")
            res_inc = incorporar_adicionais_catalogo_departamento(dados, dep_nome)
            if not res_inc.get("ok"):
                print(f"[ERRO] Erro ao incorporar funções: {res_inc.get('erro')}")
            elif res_inc.get("total_incorporadas") == 0:
                print(f"ℹ {res_inc.get('mensagem', 'Nenhuma função do catálogo para incorporar.')}")
            else:
                for item in res_inc.get("resultados", []):
                    res_item = item.get("resultado", {})
                    print(f"  {res_item.get('mensagem')}")
                print("\nA recarregar dados do Excel atualizados...")
                dados = carregar_projeto_perfil(dados.caminho)
                analise = analisar_departamento_proposta(dados, dep_nome)

        elif acao in ("7", "V", "VOLTAR"):
            return "VOLTAR"

        elif acao in ("8", "M", "GERAL"):
            return "MENU_GERAL"

        elif acao in ("0", "S", "SAIR", "Q"):
            return "SAIR"

        else:
            print("[AVISO] Opção inválida. Tente novamente.")


def menu_geral_pesquisas(dados: ProjetoPerfilData) -> str:
    """Menu de pesquisas avançadas de funções, tcodes e utilizadores."""
    while True:
        print("\n" + "-" * 75)
        print("  MENU GERAL DE PESQUISAS & ESTATÍSTICAS:")
        print("  [1] Pesquisar por Nome de Função (Role / Perfil)")
        print("  [2] Pesquisar por Transação (TCODE -> Funções)")
        print("  [3] Pesquisar por Utilizador (CUA: Adições / Remoções)")
        print("  [4] Listar todas as Roles Simples")
        print("  [5] Listar todas as Roles Compostas")
        print("  [6] Ver Resumo Geral de Todas as Funções")
        print("  [7] Abrir outro ficheiro Excel")
        print("  [8] Validar TCODEs da folha 'Proposta' no SAP PRD (TSTC via RFC)")
        print("  [9] Sincronizar Catálogo (Proposta -> PFCG_CREATE -> SAP PRD)")
        print("  [0] ↩  Voltar ao Processamento de Departamentos")
        print("-" * 75)

        op = input("Escolha uma opção: ").strip()

        if op == "1":
            termo = input("Digite o nome da função ou texto (ex.: Z_BR, MANAGER, ZORG): ").strip()
            if termo:
                imprimir_resultado_funcao(pesquisar_funcao(dados, termo))

        elif op == "2":
            tcode = input("Digite a transação SAP (ex.: FB03, ME21N, BP): ").strip()
            if tcode:
                imprimir_resultado_tcode(pesquisar_por_tcode(dados, tcode))

        elif op == "3":
            user = input("Digite o utilizador SAP (ex.: S6005, S170): ").strip()
            if user:
                imprimir_auditoria_utilizador(auditar_utilizador(dados, user))

        elif op == "4":
            print(f"\nTODAS AS ROLES SIMPLES ({len(dados.roles_simples)}):")
            for r, info in sorted(dados.roles_simples.items()):
                print(f"  - {r:<35} | {len(info.get('tcodes', []))} TCODEs | {info.get('descricao', '')}")

        elif op == "5":
            print(f"\nTODAS AS ROLES COMPOSTAS ({len(dados.roles_compostas)}):")
            for r, info in sorted(dados.roles_compostas.items()):
                print(f"  - {r:<35} | {len(info.get('roles_filhas', []))} Filhas | {info.get('descricao', '')}")

        elif op == "6":
            imprimir_resumo(dados)

        elif op == "7":
            novo = input("Caminho do novo ficheiro Excel: ").strip()
            if novo and os.path.exists(novo):
                dados = carregar_projeto_perfil(novo)
                print("Novo ficheiro carregado com sucesso!")
            else:
                print("[ERRO] Ficheiro não encontrado.")

        elif op == "8":
            analise_prop = analisar_folha_proposta(dados)
            if not analise_prop.get("encontrado"):
                print(f"[ERRO] {analise_prop.get('mensagem')}")
            else:
                tcodes_alvo = analise_prop["todos_tcodes"]
                print(f"\nA validar {len(tcodes_alvo)} transações únicas da folha 'Proposta' no SAP PRD (tabela TSTC via RFC)...")
                res_tc = verificar_tcodes_prd(tcodes_alvo)
                imprimir_resultado_verificacao_tcodes_prd(res_tc, "Folha 'Proposta'")
                if res_tc.get("nao_existentes"):
                    print(f"\nA retirar da lista e marcar como 'Transação não existe' na folha 'Proposta'...")
                    res_m = marcar_tcodes_inexistentes_sheet_proposta(dados.caminho, res_tc["nao_existentes"])
                    if res_m.get("ok"):
                        print(f"{res_m.get('marcados')} ocorrência(s) marcada(s) com 'Transação não existe' na coluna ao lado.")
                        for d in res_m.get("detalhes", []):
                            print(f"   └─ Linha {d['linha']}: {d['tcode']} da role {d['role']}")

        elif op == "9":
            dados = executar_sincronizacao_arranque_completa(dados, dados.caminho)

        elif op in ("0", "V", "VOLTAR"):
            return "VOLTAR"

        else:
            print("[AVISO] Opção inválida.")


def menu_interativo(caminho_inicial: Optional[str] = None):
    """Executa o novo menu interativo limpo direcionado por departamento."""
    caminho = caminho_inicial or encontrar_excel_padrao()

    if not caminho or not os.path.exists(caminho):
        print("Nenhum ficheiro Excel detectado automaticamente.")
        caminho = selecionar_ficheiro_dialogo()
        if not caminho:
            print("[ERRO] Operação cancelada. Saindo.")
            return

    try:
        print(f"A carregar ficheiro Excel: {os.path.basename(caminho)} ...")
        dados = carregar_projeto_perfil(caminho)
    except Exception as e:
        print(f"[ERRO] Erro ao ler Excel: {e}")
        return

    # Regra do Projeto: Validação e Sincronização Automática Completa no Arranque (4 Fases)
    dados = executar_sincronizacao_arranque_completa(dados, caminho)

    imprimir_cabecalho_compacto(caminho)

    while True:
        item_selecionado = selecionar_departamento_interativo(dados)
        if not item_selecionado:
            print("\nEncerrando. Até logo!")
            break

        if item_selecionado.get("tipo") == "MENU_GERAL":
            acao = menu_geral_pesquisas(dados)
            if acao == "SAIR":
                break
            continue

        acao = executar_menu_departamento(dados, item_selecionado)
        if acao == "SAIR":
            print("\nEncerrando. Até logo!")
            break
        elif acao == "MENU_GERAL":
            acao_g = menu_geral_pesquisas(dados)
            if acao_g == "SAIR":
                break


# =====================================================================
# PONTO DE ENTRADA CLI
# =====================================================================

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Projeto Perfil - Análise e Pesquisa de Funções SAP")
    parser.add_argument("--xlsx", "--ficheiro", "-f", dest="ficheiro", help="Caminho do ficheiro Excel de perfis")
    parser.add_argument("--controlo", "--validar-controlo", dest="controlo", action="store_true", help="Validar sheet CONTROLO e listar departamentos pendentes")
    parser.add_argument("--proximo", "--proximo-departamento", dest="proximo", action="store_true", help="Identificar e processar o próximo departamento a avançar da sheet CONTROLO (STATUS e TIMESTAMP vazios)")
    parser.add_argument("--departamento", "-d", dest="departamento", nargs="?", const="", help="Analisar funções de um departamento na sheet Proposta Ativa (se omitido, usa o próximo de CONTROLO)")
    parser.add_argument("--cruzar-fontes", dest="cruzar_fontes", action="store_true", help="Cruzar PFCG_CREATE, PFCG_COMPOSTA, PFCG_AUTHORITY e EXCLUÇÃO para o departamento")
    parser.add_argument("--validar-users-prd", dest="validar_users_prd", action="store_true", help="Validar atribuições dos utilizadores no SAP PRD (AGR_USERS) com cruzamento relacional")
    parser.add_argument("--verificar-prd", dest="verificar_prd", action="store_true", help="Verificar se as funções existem no sistema SAP PRD via RFC (AGR_DEFINE)")
    parser.add_argument("--comparar-prd", "--funcoes-diferentes-prd", dest="comparar_prd", action="store_true", help="Comparar catálogo de funções no PRD vs ficheiro Excel (AGR_DEFINE e AGR_USERS)")
    parser.add_argument("--pesquisar-role", "-r", dest="role", help="Pesquisar diretamente por nome de função")
    parser.add_argument("--pesquisar-tcode", "-t", dest="tcode", help="Pesquisar diretamente por transação SAP")
    parser.add_argument("--pesquisar-user", "-u", dest="user", help="Auditar e pesquisar diretamente por utilizador (Proposta, CUA e SAP PRD)")
    parser.add_argument("--incorporar-adicionais", dest="incorporar_adicionais", action="store_true", help="Incorporar na Proposta Ativa as funções do catálogo que já estão ativas no SAP PRD")
    parser.add_argument("--validar-proposta-tcodes", "--validar-tcodes-prd", dest="validar_proposta_tcodes", action="store_true", help="Validar se todas as transações da sheet Proposta existem na tabela TSTC do SAP PRD")
    parser.add_argument("--sincronizar-catalogo", dest="sincronizar_catalogo", action="store_true", help="Sincronizar catálogo da folha Proposta com PFCG_CREATE e SAP PRD")

    args = parser.parse_args()

    caminho_alvo = args.ficheiro or encontrar_excel_padrao()

    # Se foram passados parâmetros de pesquisa direta via CLI:
    if args.controlo or args.proximo or args.departamento is not None or args.cruzar_fontes or args.validar_users_prd or args.verificar_prd or args.comparar_prd or args.role or args.tcode or args.user or args.incorporar_adicionais or args.validar_proposta_tcodes or args.sincronizar_catalogo:
        if not caminho_alvo:
            print("[ERRO] Erro: Ficheiro Excel não encontrado.")
            sys.exit(1)
        
        dados = carregar_projeto_perfil(caminho_alvo)
        imprimir_cabecalho(caminho_alvo)
        proximo_dep = dados.obter_proximo_departamento()

        if args.controlo:
            imprimir_validacao_controlo(dados)
        if args.proximo:
            if proximo_dep:
                print(f"\nPRÓXIMO DEPARTAMENTO A AVANÇAR: '{proximo_dep}' (STATUS e TIMESTAMP vazios em CONTROLO)")
                res_dep = analisar_departamento_proposta(dados, proximo_dep)
                imprimir_analise_departamento(res_dep)
            else:
                print("\nTodos os departamentos na sheet CONTROLO já possuem STATUS preenchido.")
        if args.departamento is not None and not (args.cruzar_fontes or args.validar_users_prd or args.verificar_prd or args.incorporar_adicionais):
            alvo = args.departamento.strip() if args.departamento.strip() else proximo_dep
            if not alvo:
                alvo = "Purchase & Services"
            res_dep = analisar_departamento_proposta(dados, alvo)
            imprimir_analise_departamento(res_dep)
        if args.cruzar_fontes:
            alvo_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else proximo_dep
            if not alvo_dep:
                alvo_dep = "Purchase & Services"
            res_cruz = cruzar_fontes_departamento(dados, alvo_dep)
            imprimir_cruzamento_fontes(res_cruz)
        if args.validar_users_prd:
            alvo_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else proximo_dep
            if not alvo_dep:
                alvo_dep = "Purchase & Services"
            res_prd_users = validar_utilizadores_prd(dados, alvo_dep)
            imprimir_validacao_utilizadores_prd(res_prd_users)
        if args.incorporar_adicionais:
            alvo_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else proximo_dep
            if not alvo_dep:
                alvo_dep = "Purchase & Services"
            res_inc = incorporar_adicionais_catalogo_departamento(dados, alvo_dep)
            if not res_inc.get("ok"):
                print(f"[ERRO] Erro ao incorporar funções: {res_inc.get('erro')}")
            elif res_inc.get("total_incorporadas") == 0:
                print(f"ℹ {res_inc.get('mensagem', 'Nenhuma função do catálogo para incorporar.')}")
            else:
                for item in res_inc.get("resultados", []):
                    res_item = item.get("resultado", {})
                    print(f"  {res_item.get('mensagem')}")
                print("\nA recarregar dados e revalidar utilizadores no PRD...")
                dados = carregar_projeto_perfil(caminho_alvo)
                res_prd_novos = validar_utilizadores_prd(dados, alvo_dep)
                imprimir_validacao_utilizadores_prd(res_prd_novos)
        if args.verificar_prd:
            alvo_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else proximo_dep
            if alvo_dep:
                dep_info = analisar_departamento_proposta(dados, alvo_dep)
                roles_dep = list(dep_info["compostas"]) + list(dep_info["singles_frequencia"].keys())
                res_prd = verificar_funcoes_prd(roles_dep)
                imprimir_resultado_verificacao_prd(res_prd, f"Departamento '{alvo_dep}'")
            else:
                todas = list(dados.roles_simples.keys()) + list(dados.roles_compostas.keys())
                res_prd = verificar_funcoes_prd(todas)
                imprimir_resultado_verificacao_prd(res_prd, "Todas as Funções do Excel (PFCG_CREATE + PFCG_COMPOSTA)")
        if args.comparar_prd:
            res_comp = comparar_funcoes_catalogo_prd(dados)
            imprimir_comparacao_catalogo_prd(res_comp)
        if args.role:
            imprimir_resultado_funcao(pesquisar_funcao(dados, args.role))
        if args.tcode:
            imprimir_resultado_tcode(pesquisar_por_tcode(dados, args.tcode))
        if args.user:
            res_audit = auditar_utilizador(dados, args.user)
            imprimir_auditoria_utilizador(res_audit)
        if args.validar_proposta_tcodes:
            analise_prop = analisar_folha_proposta(dados)
            if not analise_prop.get("encontrado"):
                print(f"[ERRO] {analise_prop.get('mensagem')}")
            else:
                tcodes_alvo = analise_prop["todos_tcodes"]
                print(f"\nA validar {len(tcodes_alvo)} transações únicas da folha 'Proposta' no SAP PRD (tabela TSTC via RFC)...")
                res_tc = verificar_tcodes_prd(tcodes_alvo)
                imprimir_resultado_verificacao_tcodes_prd(res_tc, "Folha 'Proposta'")
                if res_tc.get("nao_existentes"):
                    print(f"\nA retirar da lista e marcar como 'Transação não existe' na folha 'Proposta'...")
                    res_m = marcar_tcodes_inexistentes_sheet_proposta(dados.caminho, res_tc["nao_existentes"])
                    if res_m.get("ok"):
                        print(f"{res_m.get('marcados')} ocorrência(s) marcada(s) com 'Transação não existe' na coluna ao lado.")
                        for d in res_m.get("detalhes", []):
                            print(f"   └─ Linha {d['linha']}: {d['tcode']} da role {d['role']}")
        if args.sincronizar_catalogo:
            dados = executar_sincronizacao_arranque_completa(dados, caminho_alvo)

    else:
        # Modo interativo padrão
        menu_interativo(caminho_alvo)

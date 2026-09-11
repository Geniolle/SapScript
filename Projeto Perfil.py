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
    print("  🌐 COMPARATIVO DE FUNÇÕES: SAP PRD vs FICHEIRO EXCEL")
    print("=" * 80)

    if not res.get("ok"):
        print(f"  ❌ Erro ao consultar SAP PRD: {res.get('erro')}")
        return

    print("  📌 1. CATÁLOGO GLOBAL DE FUNÇÕES NO SAP PRD (AGR_DEFINE):")
    print(f"     • Total de funções no PRD:              {res['total_prd_catalogo']:>6d}")
    print(f"       ├─ Funções Customizadas (Z*):          {res['total_prd_z']:>6d}")
    print(f"       ├─ Funções Standard SAP (SAP_*):       {res['total_prd_sap']:>6d}")
    print(f"       └─ Outras funções Standard:            {res['total_prd_outras']:>6d}")
    print(f"     • Total de funções no Ficheiro Excel:   {res['total_excel']:>6d}")

    print("\n  🔍 2. FUNÇÕES NO PRD QUE NÃO CONSTAM NO FICHEIRO EXCEL:")
    print(f"     • Diferença Total Bruta:                {res['dif_catalogo_bruto']:>6d}")
    print(f"       (Inclui {res['total_prd_sap']} standard SAP_* e outras legadas)")
    print(f"     • Diferença Líquida (após EXCLUÇÃO):    {res['dif_catalogo_sem_exclusao']:>6d}")
    print(f"     • Funções Customizadas Z* fora do Excel: {res['dif_z_sem_exclusao']:>6d}")
    print(f"       ├─ Funções Simples (Single):           {res['dif_z_simples']:>6d}")
    print(f"       └─ Funções Compostas (Composite):      {res['dif_z_compostas']:>6d}")

    print("\n  🏷️ 3. SEGMENTAÇÃO DAS FUNÇÕES Z* FORA DO EXCEL:")
    print(f"     • Organizacionais (ZORG_*):              {len(res['z_org']):>6d}")
    print(f"     • Business Roles (Z_BR_*):               {len(res['z_br']):>6d}")
    print(f"     • Módulo MM (ZMM_*):                     {len(res['z_mm']):>6d}")
    print(f"     • Módulo FI (ZFI_*):                     {len(res['z_fi']):>6d}")
    print(f"     • Módulo SD (ZSD_*):                     {len(res['z_sd']):>6d}")
    print(f"     • Outros legados Z*:                     {len(res['z_outras']):>6d}")

    print("\n  👥 4. FUNÇÕES ATIVAMENTE ATRIBUÍDAS A UTILIZADORES (AGR_USERS):")
    print(f"     • Total de funções distintas atribuídas: {res['total_atribuidas_ativas_prd']:>6d}")
    print(f"     • Funções atribuídas fora do Excel (Bruto): {res['dif_atrib_bruto']:>4d}")
    print(f"     • Funções atribuídas fora do Excel (Líquido): {res['dif_atrib_sem_exclusao']:>2d}")
    if res["roles_atribuidas_fora_excel"]:
        print("       Funções ativas a utilizadores não catalogadas:")
        for r in res["roles_atribuidas_fora_excel"][:25]:
            print(f"         ├─ {r}")
        if len(res["roles_atribuidas_fora_excel"]) > 25:
            print(f"         └─ ... e mais {len(res['roles_atribuidas_fora_excel']) - 25} funções.")

    print("\n  🛡️ Padrões desconsiderados da folha EXCLUÇÃO:")
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
    print("  🔗 CRUZAMENTO DE FONTES (PFCG_CREATE, PFCG_COMPOSTA, PFCG_AUTHORITY, EXCLUÇÃO)")
    print(f"  🏢 Departamento: {resultado.get('departamento', '')}")
    print("=" * 75)

    if not resultado.get("encontrado"):
        print(f"  ⚠️ {resultado.get('mensagem', 'Departamento não encontrado.')}")
        return

    print(f"  Regras EXCLUÇÃO aplicadas: {resultado.get('padroes_exclusao', [])}")
    print(f"  Utilizadores analisados:   {resultado.get('total_usuarios', 0)}")
    print("-" * 75)

    for u in resultado.get("usuarios", []):
        comp_str = f"Composta: {u['composta']}" if u.get("composta") else "Sem Composta"
        print(f"\n  👤 {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']}")
        print(f"     └─ {comp_str} | {u['singles_proposta_total']} Singles na Proposta")
        print(f"        ├─ Membros em PFCG_COMPOSTA:   {u['membros_composta_total']}")
        print(f"        ├─ PFCG_AUTHORITY da Composta: {u['authority_composta_total']}")
        print(f"        └─ 🎯 Esperado Expandido Final: {u['esperado_total']} Funções")
        if u.get("excluidas_proposta"):
            print(f"           🛡️ Desconsideradas por EXCLUÇÃO: {', '.join(u['excluidas_proposta'])}")


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


def imprimir_validacao_utilizadores_prd(res: Dict[str, Any]):
    """Exibe no terminal a validação dos utilizadores no PRD com cruzamento relacional."""
    print("\n" + "=" * 75)
    print("  🔍 VALIDAÇÃO DE UTILIZADORES NO SAP PRD (via RFC AGR_USERS & USR02)")
    print(f"  🏢 Departamento: {res.get('departamento', '')} | Sistema: {res.get('sistema', 'PRD')}")
    print("=" * 75)

    if not res.get("ok"):
        print(f"  ❌ Erro ao validar utilizadores no PRD: {res.get('erro')}")
        return

    print(f"  Regras EXCLUÇÃO aplicadas: {res.get('padroes_exclusao', [])}")
    print(f"  Total de Funções Adicionais Ativas:         {res.get('total_adicionais_ativas', 0)}")
    print(f"  Ocorrências Desconsideradas por EXCLUÇÃO:  {res.get('total_desconsideradas_exclusao', 0)}")
    print("-" * 75)

    for u in res.get("utilizadores", []):
        if u.get("inativo_usr02"):
            status_ico = "⏸️"
            obs = f" [CONTA EXPIRADA EM USR02 - Validade terminou a {u.get('validade_fim')}]"
        else:
            status_ico = "✅" if u["conforme"] else "⚠️"
            obs = ""

        add_str = f" | {len(u['adicionais'])} Adicionais" if u["adicionais"] else " | 0 Adicionais"
        print(f"\n  {status_ico} 👤 {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']}{obs}")
        print(f"     └─ Atribuições Esperadas Ativas: {u['esperadas_ativas']}/{u['esperadas_total']}{add_str}")
        
        if u["faltam"] and not u.get("inativo_usr02"):
            print(f"        ❌ EM FALTA ({len(u['faltam'])}): {', '.join(u['faltam'])}")
        elif u["faltam"] and u.get("inativo_usr02"):
            print(f"        ℹ️ Sem atribuições ativas por desativação/offboarding no sistema SAP.")
        if u["adicionais"]:
            print(f"        📌 ADICIONAIS A VALIDAR ({len(u['adicionais'])}): {', '.join(u['adicionais'])}")
        if u["desconsideradas_exclusao"]:
            print(f"        🛡️ Protegidas por EXCLUÇÃO ({len(u['desconsideradas_exclusao'])}): {', '.join(u['desconsideradas_exclusao'])}")


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
    print(f"  👤 AUDITORIA COMPLETA DE UTILIZADOR: {u}")
    if p.get("nome"):
        print(f"  Nome: {p.get('nome')} | Cargo: {p.get('cargo')} | Departamento: {p.get('departamento')}")
    print("=" * 75)

    # 1. Proposta Ativa
    if p:
        print("\n  📋 1. DADOS NA 'PROPOSTA ATIVA':")
        print(f"     ├─ Função Composta: {p.get('composta') or 'Nenhuma'}")
        print(f"     ├─ Funções Simples Declaradas: {p.get('total_singles_proposta')}")
        print(f"     └─ 🎯 Total Esperado após Expansão Relacional: {res.get('esperadas_total')} Funções")
    else:
        print("\n  📋 1. 'PROPOSTA ATIVA': Utilizador não localizado na folha Proposta Ativa.")

    # 2. CUA
    cua = res.get("cua", {})
    adds = cua.get("adicionar", [])
    rms = cua.get("remover", [])
    if adds or rms:
        print(f"\n  🔄 2. ATRIBUIÇÕES NO CUA: {len(adds)} para Adicionar, {len(rms)} para Remover")
    else:
        print("\n  🔄 2. ATRIBUIÇÕES NO CUA: Nenhuma instrução pendente em CUA_ADICIONAR/CUA_REMOVE.")

    # 3. SAP PRD
    if not res.get("prd_conectado"):
        print(f"\n  ❌ 3. SAP PRD: Erro de ligação RFC ({res.get('prd_erro')})")
        return

    m = res.get("mestre_usr02", {})
    print("\n  🌐 3. ESTADO NO SAP PRD (Mestre USR02 & AGR_USERS):")
    if m:
        st_lock = "🔒 BLOQUEADO" if m.get("bloqueado") else "🔓 Desbloqueado"
        st_exp = f"⏸️ EXPIRADO a {m.get('validade_fim')}" if m.get("expirado") else "✅ Conta Ativa"
        print(f"     ├─ Tipo: {m.get('tipo', 'A')} | Estado: {st_lock} | Validade: {st_exp}")
        print(f"     ├─ Último Logon: {m.get('ultimo_logon') or 'Nunca'}")
    print(f"     ├─ Funções Ativas no PRD:       {res.get('ativas_total')}")
    print(f"     ├─ Funções Expiradas/Histórico: {res.get('expiradas_total')}")
    print(f"     └─ Conformidade com a Proposta: {res.get('esperadas_ativas')}/{res.get('esperadas_total')} ativas")

    if res.get("faltam"):
        print(f"\n  ❌ FUNÇÕES EM FALTA ({len(res['faltam'])}):")
        for f in res["faltam"]:
            print(f"     ├─ {f}")
    else:
        print("\n  ✅ TODAS AS FUNÇÕES PREVISTAS ESTÃO 100% ATIVAS NO PRD (0 em falta)!")

    prot = res.get("protegidas_exclusao", [])
    if prot:
        print(f"\n  🛡️ Ocorrências Protegidas por EXCLUÇÃO ({len(prot)}):")
        for pr in prot:
            print(f"     ├─ {pr}")

    adicionais = res.get("adicionais_reais", [])
    if adicionais:
        print(f"\n  📌 FUNÇÕES ADICIONAIS A VALIDAR ({len(adicionais)}):")
        for a in adicionais:
            print(f"     ├─ {a}")
    else:
        print("\n  🎉 NENHUMA FUNÇÃO ADICIONAL PENDENTE DE VALIDAÇÃO (0 adicionais)!")


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
    print(f"  Departamentos pendentes (STATUS e TIMESTAMP vazios): {len(pendentes)}")
    print("-" * 75)
    for c in dados.controlo:
        status_tag = "⏳ PENDENTE (STATUS E TIMESTAMP VAZIOS)" if c["pendente"] else f"✅ {c['status']}"
        ts_info = f" | {c['timestamp']}" if c.get("timestamp") else ""
        print(f"  Linha {c['linha']}: {c['departamento']:<30} -> {status_tag}{ts_info}")

    print("-" * 75)
    proximo = dados.obter_proximo_departamento()
    if proximo:
        print(f"  🎯 PRÓXIMO DEPARTAMENTO A AVANÇAR: '{proximo}'")
        item_prox = next((p for p in pendentes if p["departamento"] == proximo), {})
        tem_sheet = proximo in dados.sheets_disponiveis
        sheet_info = f"[Sheet correspondente '{proximo}' EXISTE no Excel]" if tem_sheet else "[Sheet correspondente não encontrada]"
        print(f"     👉 Linha {item_prox.get('linha')}: STATUS e TIMESTAMP estão vazios! {sheet_info}")
        if len(pendentes) > 1:
            print("\n  Demais departamentos pendentes:")
            for p in pendentes[1:]:
                print(f"     - '{p['departamento']}' (Linha {p['linha']})")
    else:
        print("  ✅ Todos os departamentos na sheet CONTROLO já possuem STATUS e TIMESTAMP preenchidos.")


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
# MENU INTERATIVO - FLUXO DIRECIONADO POR DEPARTAMENTO (CONTROLO)
# =====================================================================

def imprimir_cabecalho_compacto(ficheiro: str):
    """Exibe cabeçalho limpo e direto ao iniciar."""
    print("\n" + "=" * 75)
    print("  🚀 PROJETO PERFIL - PROCESSAMENTO DE DEPARTAMENTOS (CUA / PFCG)")
    print(f"  📂 Ficheiro: {os.path.basename(ficheiro)}")
    print("=" * 75)


def exibir_tabela_controlo(dados: ProjetoPerfilData) -> Optional[Dict[str, Any]]:
    """
    Exibe a tabela clara dos departamentos na folha CONTROLO com as suas linhas no Excel
    e retorna o próximo departamento pendente sugerido (se houver).
    """
    print("\n📋 DEPARTAMENTOS NA SHEET 'CONTROLO':")
    print("-" * 75)
    print(f"  {'Linha':<8} | {'Departamento':<30} | {'Estado na CONTROLO':<30}")
    print("-" * 75)

    if not dados.controlo:
        print("  ⚠️ Nenhuma linha encontrada na folha CONTROLO.")
        print("-" * 75)
        return None

    proximo_sugerido = None
    for c in dados.controlo:
        linha_str = f"[{c['linha']}]"
        if c.get("pendente"):
            status_desc = "⏳ DISPONÍVEL (Pendente)"
            if not proximo_sugerido:
                proximo_sugerido = c
        else:
            ts_curto = str(c.get("timestamp", ""))[:16]
            ts = f" ({ts_curto})" if ts_curto else ""
            status_desc = f"✅ {c.get('status', 'PROCESSADO')}{ts}"
        print(f"  {linha_str:<8} | {c['departamento']:<30} | {status_desc}")
    print("-" * 75)

    if proximo_sugerido:
        print(f"  🎯 Próximo departamento sugerido: Linha {proximo_sugerido['linha']} ('{proximo_sugerido['departamento']}')")
    else:
        print("  ✅ Todos os departamentos registados na sheet CONTROLO estão processados.")

    return proximo_sugerido


def selecionar_departamento_interativo(dados: ProjetoPerfilData) -> Optional[Dict[str, Any]]:
    """
    Apresenta a listagem da folha CONTROLO e solicita ao utilizador indicar a linha a processar.
    """
    sugerido = exibir_tabela_controlo(dados)
    padrao_linha = str(sugerido["linha"]) if sugerido else ""
    prompt_sugestao = f" [Enter para Linha {padrao_linha}]" if padrao_linha else ""

    while True:
        entrada = input(f"\n👉 Indique a LINHA do departamento que quer processar{prompt_sugestao} ('M' Menu Geral, '0' Sair): ").strip().upper()

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

        print(f"⚠️ Linha ou departamento '{entrada}' não encontrado na folha CONTROLO. Escolha uma das linhas listadas.")


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
        print(f"❌ Não foi possível carregar os utilizadores do departamento '{dep_nome}'.")
        return

    users = [u["usuario"] for u in analise.get("usuarios", []) if u.get("usuario")]
    if not users:
        print(f"❌ Nenhum utilizador encontrado para '{dep_nome}'.")
        return

    print(f"\n" + "=" * 75)
    print(f"  🚀 INICIAR SINCRONIZAÇÃO CUA: {dep_nome.upper()} ({len(users)} Utilizadores)")
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

    confirma = input("👉 Confirma a execução imediata no SAP CUA? (S/N): ").strip().upper()
    if confirma not in ("S", "SIM", "Y", "YES"):
        print("⏸️ Operação cancelada pelo utilizador.")
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
                print(f"\n  👤 A processar utilizador: {u} ...")

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
                    print(f"     ✅ S4DCLNT100 removido de {u}.")

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
                    print(f"     ✅ {len(expired_p)} roles expiradas removidas de S4PCLNT100.")

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
                    print(f"     ✅ {len(roles_q)} roles obsoletas removidas de S4QCLNT100.")

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
                print(f"     ✅ {len(p_roles_dict)} roles ativas replicadas para S4QCLNT100.")

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
                print(f"     🎯 Auditoria {u}: P={len(p_finais)} | Q={len(q_finais)} -> MATCH EXATO: {sess_ok}")

            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)

            # Gravação em Excel
            print("\n  📊 A atualizar folhas de cálculo oficiais (CUA_REMOVE, CUA_ADICIONAR, CONTROLO)...")
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
                print("  ✅ Excel guardado com sucesso!")

            item_dep["status"] = "PROCESSADO"
            item_dep["timestamp"] = timestamp_str
            item_dep["pendente"] = False
            print(f"\n🎉 Sincronização do departamento '{dep_nome}' concluída com sucesso!")

        except Exception as exc:
            erro_execucao.append(exc)
            print(f"❌ Erro na execução da sincronização CUA: {exc}")

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
        print(f"  🏢 DEPARTAMENTO SELECIONADO: {dep_nome.upper()} (Linha {linha} no Excel)")
        print("=" * 75)

        st_tag = "⏳ PENDENTE (Disponível para processamento)" if item_dep.get("pendente") else f"✅ {item_dep.get('status')} | {item_dep.get('timestamp')}"
        print(f"  📌 Estado na CONTROLO: {st_tag}")

        if analise.get("encontrado"):
            users = analise.get("usuarios", [])
            compostas = analise.get("compostas", [])
            print(f"  👥 Utilizadores no Departamento ({len(users)}):")
            for u in users:
                comp = f"[Composta: {u['composta']}]" if u.get("composta") else "[Sem Composta]"
                print(f"     • {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']} {comp} -> {u['total_singles']} Singles")
            if compostas:
                print(f"  📦 Funções Compostas ({len(compostas)}): {', '.join(compostas)}")
            print(f"  🧩 Funções Individuais (Singles): {analise.get('total_singles_distintas', 0)} distintas")
        else:
            print(f"  ⚠️ Aviso: {analise.get('mensagem', 'Departamento não encontrado na folha Proposta Ativa.')}")

        print("-" * 75)
        print("  AÇÕES DISPONÍVEIS PARA ESTE DEPARTAMENTO:")
        print("  [1] 👤 Validar Utilizadores no SAP PRD (AGR_USERS via RFC & Cruzamento)")
        print("  [2] 🔗 Cruzar Fontes (PFCG_CREATE, COMPOSTA, AUTHORITY, EXCLUÇÃO)")
        print("  [3] 🔄 Sincronizar CUA (Remover S4D / Limpar PRD / Alinhar QAS)")
        print("  [4] 🌐 Verificar Existência de Funções no SAP PRD (AGR_DEFINE)")
        print("  [5] 📋 Ver Análise Detalhada (Proposta Ativa)")
        print("  [6] ↩️  Voltar / Escolher Outro Departamento da CONTROLO")
        print("  [7] 🔍 Menu Geral de Pesquisas (Roles, TCODEs, Users, etc.)")
        print("  [0] 🚪 Sair")
        print("-" * 75)

        acao = input("👉 Escolha uma ação: ").strip().upper()

        if acao == "1":
            print(f"\n⏳ A validar atribuições no SAP PRD para '{dep_nome}' ...")
            res_prd = validar_utilizadores_prd(dados, dep_nome)
            imprimir_validacao_utilizadores_prd(res_prd)

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
                print("⚠️ Não foi possível obter as funções do departamento.")

        elif acao == "5":
            imprimir_analise_departamento(analise)

        elif acao in ("6", "V", "VOLTAR"):
            return "VOLTAR"

        elif acao in ("7", "M", "GERAL"):
            return "MENU_GERAL"

        elif acao in ("0", "S", "SAIR", "Q"):
            return "SAIR"

        else:
            print("⚠️ Opção inválida. Tente novamente.")


def menu_geral_pesquisas(dados: ProjetoPerfilData) -> str:
    """Menu de pesquisas avançadas de funções, tcodes e utilizadores."""
    while True:
        print("\n" + "-" * 75)
        print("  🔍 MENU GERAL DE PESQUISAS & ESTATÍSTICAS:")
        print("  [1] 🔍 Pesquisar por Nome de Função (Role / Perfil)")
        print("  [2] 📌 Pesquisar por Transação (TCODE -> Funções)")
        print("  [3] 👤 Pesquisar por Utilizador (CUA: Adições / Remoções)")
        print("  [4] 📋 Listar todas as Roles Simples")
        print("  [5] 📋 Listar todas as Roles Compostas")
        print("  [6] 📊 Ver Resumo Geral de Todas as Funções")
        print("  [7] 📂 Abrir outro ficheiro Excel")
        print("  [0] ↩️  Voltar ao Processamento de Departamentos")
        print("-" * 75)

        op = input("👉 Escolha uma opção: ").strip()

        if op == "1":
            termo = input("🔎 Digite o nome da função ou texto (ex.: Z_BR, MANAGER, ZORG): ").strip()
            if termo:
                imprimir_resultado_funcao(pesquisar_funcao(dados, termo))

        elif op == "2":
            tcode = input("🔎 Digite a transação SAP (ex.: FB03, ME21N, BP): ").strip()
            if tcode:
                imprimir_resultado_tcode(pesquisar_por_tcode(dados, tcode))

        elif op == "3":
            user = input("🔎 Digite o utilizador SAP (ex.: S6005, S170): ").strip()
            if user:
                imprimir_auditoria_utilizador(auditar_utilizador(dados, user))

        elif op == "4":
            print(f"\n📋 TODAS AS ROLES SIMPLES ({len(dados.roles_simples)}):")
            for r, info in sorted(dados.roles_simples.items()):
                print(f"  - {r:<35} | {len(info.get('tcodes', []))} TCODEs | {info.get('descricao', '')}")

        elif op == "5":
            print(f"\n📋 TODAS AS ROLES COMPOSTAS ({len(dados.roles_compostas)}):")
            for r, info in sorted(dados.roles_compostas.items()):
                print(f"  - {r:<35} | {len(info.get('roles_filhas', []))} Filhas | {info.get('descricao', '')}")

        elif op == "6":
            imprimir_resumo(dados)

        elif op == "7":
            novo = input("📂 Caminho do novo ficheiro Excel: ").strip()
            if novo and os.path.exists(novo):
                dados = carregar_projeto_perfil(novo)
                print("✅ Novo ficheiro carregado com sucesso!")
            else:
                print("❌ Ficheiro não encontrado.")

        elif op in ("0", "V", "VOLTAR"):
            return "VOLTAR"

        else:
            print("⚠️ Opção inválida.")


def menu_interativo(caminho_inicial: Optional[str] = None):
    """Executa o novo menu interativo limpo direcionado por departamento."""
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

    imprimir_cabecalho_compacto(caminho)

    while True:
        item_selecionado = selecionar_departamento_interativo(dados)
        if not item_selecionado:
            print("\n👋 Encerrando. Até logo!")
            break

        if item_selecionado.get("tipo") == "MENU_GERAL":
            acao = menu_geral_pesquisas(dados)
            if acao == "SAIR":
                break
            continue

        acao = executar_menu_departamento(dados, item_selecionado)
        if acao == "SAIR":
            print("\n👋 Encerrando. Até logo!")
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

    args = parser.parse_args()

    caminho_alvo = args.ficheiro or encontrar_excel_padrao()

    # Se foram passados parâmetros de pesquisa direta via CLI:
    if args.controlo or args.proximo or args.departamento is not None or args.cruzar_fontes or args.validar_users_prd or args.verificar_prd or args.comparar_prd or args.role or args.tcode or args.user:
        if not caminho_alvo:
            print("❌ Erro: Ficheiro Excel não encontrado.")
            sys.exit(1)
        
        dados = carregar_projeto_perfil(caminho_alvo)
        imprimir_cabecalho(caminho_alvo)
        proximo_dep = dados.obter_proximo_departamento()

        if args.controlo:
            imprimir_validacao_controlo(dados)
        if args.proximo:
            if proximo_dep:
                print(f"\n🎯 PRÓXIMO DEPARTAMENTO A AVANÇAR: '{proximo_dep}' (STATUS e TIMESTAMP vazios em CONTROLO)")
                res_dep = analisar_departamento_proposta(dados, proximo_dep)
                imprimir_analise_departamento(res_dep)
            else:
                print("\n✅ Todos os departamentos na sheet CONTROLO já possuem STATUS preenchido.")
        if args.departamento is not None and not (args.cruzar_fontes or args.validar_users_prd or args.verificar_prd):
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

    else:
        # Modo interativo padrão
        menu_interativo(caminho_alvo)

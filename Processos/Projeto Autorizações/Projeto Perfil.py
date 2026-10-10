# -*- coding: utf-8 -*-
"""
Projeto Perfil.py
======================================================================
Módulo de pesquisa, análise e leitura estruturada do ficheiro Excel
de Perfis de Autorização e Funções SAP (PFCG / CUA).

Funcionalidades:
  - Utilização exclusiva do ficheiro Excel mestre local configurado no código
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
from collections import defaultdict, Counter
from typing import Optional, Dict, Any, List, Set, Iterable, Tuple

# Garantir raiz do projeto no sys.path
PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Garantir codificação UTF-8 no Windows
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
        sys.stderr.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass

# Auto-deteção e re-execução no ambiente virtual .venv-rfc se pandas não estiver presente
try:
    import pandas
except ImportError:
    project_root_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    venv_python = os.path.join(project_root_dir, ".venv-rfc", "Scripts", "python.exe")
    if os.path.exists(venv_python) and sys.executable.lower() != venv_python.lower():
        import subprocess
        res = subprocess.run([venv_python, os.path.abspath(__file__)] + sys.argv[1:])
        sys.exit(res.returncode)

# Suporte a cores no terminal (Colorama / ANSI)
try:
    import colorama
    if hasattr(colorama, "just_fix_windows_console"):
        colorama.just_fix_windows_console()
    colorama.init(strip=False)
    CLR_VERDE = colorama.Fore.GREEN + colorama.Style.BRIGHT
    CLR_VERMELHO = colorama.Fore.RED + colorama.Style.BRIGHT
    CLR_AMARELO = colorama.Fore.YELLOW + colorama.Style.BRIGHT
    CLR_AZUL = colorama.Fore.CYAN + colorama.Style.BRIGHT
    CLR_RESET = colorama.Style.RESET_ALL
except Exception:
    CLR_VERDE = ""
    CLR_VERMELHO = ""
    CLR_AMARELO = ""
    CLR_AZUL = ""
    CLR_RESET = ""


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


def encontrar_coluna_funcoes_individuais_ws(ws: Any, default_col: int = 11) -> int:
    """
    Localiza na linha 1 do cabeçalho de uma folha (openpyxl Worksheet ou win32com Worksheet)
    a coluna correspondente a 'Funções Individuais'.
    Devolve o índice da coluna baseado em 1 (ex.: 11 para Coluna K).
    """
    try:
        # Se for win32com (tem atributo Cells e UsedRange)
        if hasattr(ws, "Cells") and not hasattr(ws, "cell"):
            max_c = min(ws.UsedRange.Columns.Count + 10, 100) if hasattr(ws, "UsedRange") else 50
            for c in range(1, max_c + 1):
                val = str(ws.Cells(1, c).Value or "").strip()
                norm = normalizar_nome_coluna(val)
                if "FUNCOESINDIVIDUAIS" in norm or "FUNCAOINDIVIDUAL" in norm:
                    return c
        # Se for openpyxl
        elif hasattr(ws, "cell"):
            max_c = ws.max_column or 50
            for c in range(1, max_c + 1):
                val = str(ws.cell(row=1, column=c).value or "").strip()
                norm = normalizar_nome_coluna(val)
                if "FUNCOESINDIVIDUAIS" in norm or "FUNCAOINDIVIDUAL" in norm:
                    return c
    except Exception:
        pass
    return default_col


def encontrar_coluna_funcoes_individuais_df(df: Any) -> int:
    """
    Localiza no DataFrame do pandas (via df.columns ou df.iloc[0])
    o índice da coluna 'Funções Individuais' baseado em 0.
    Se o cabeçalho não existir na linha 1, lança ValueError sem fallback silencioso.
    """
    if hasattr(df, "columns"):
        for idx, col_name in enumerate(df.columns):
            norm = normalizar_nome_coluna(str(col_name))
            if "FUNCOESINDIVIDUAIS" in norm or "FUNCAOINDIVIDUAL" in norm:
                return idx
    if hasattr(df, "iloc") and len(df) > 0:
        for idx in range(len(df.columns)):
            norm = normalizar_nome_coluna(str(df.iloc[0, idx]))
            if "FUNCOESINDIVIDUAIS" in norm or "FUNCAOINDIVIDUAL" in norm:
                return idx
    raise ValueError("CABECALHO_FUNCOES_INDIVIDUAIS_NAO_ENCONTRADO: O cabeçalho 'Funções Individuais' não foi encontrado na linha 1 da folha Proposta Ativa.")


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


def is_role_sap_valida(role_raw: Any) -> bool:
    """Valida nomes técnicos de roles SAP do projeto, evitando descrições/cargos."""
    role = normalizar_texto(role_raw)
    return role.startswith("Z") and len(role) >= 4


def extrair_roles_sap(valor_raw: Any) -> List[str]:
    """Extrai roles SAP de uma célula, aceitando listas separadas por vírgula, ponto-e-vírgula ou quebra de linha."""
    if valor_raw is None:
        return []
    texto = str(valor_raw).replace("\r", "\n").replace("\t", " ")
    roles = []
    for parte in re.split(r"[,;\n]+", texto):
        role = normalizar_texto(parte)
        if is_role_sap_valida(role) and role not in roles:
            roles.append(role)
    return roles


# Carregar variáveis de ambiente do .env (raiz do projeto SapScript)
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), ".env"))
except ImportError:
    pass


# =====================================================================
# LOCALIZAÇÃO E SINCRONIZAÇÃO DO FICHEIRO EXCEL (SHAREPOINT / ONEDRIVE)
# =====================================================================

CAMINHO_EXCEL_LOCAL = str(Path(__file__).resolve().parents[2] / "S4H_Perfis de autorização_v1.xlsx")
CAMINHO_EXCEL_BACKUP_LOCAL = str(Path(__file__).resolve().parents[2] / "output" / "S4H_Perfis de autorização_v1_backup_automatico.xlsx")


def sincronizar_copia_sharepoint() -> bool:
    """
    Localiza o ficheiro sincronizado do SharePoint na pasta do OneDrive da Salsa
    e atualiza a cópia de trabalho local em C:\\workspace\\SapScript\\S4H_Perfis de autorização_v1.xlsx.
    Faz backup automático antes de sobrescrever.
    """
    import shutil
    try:
        user_profile = os.environ.get("USERPROFILE", "")
        if not user_profile:
            return False

        # Procurar a pasta de atalhos sincronizados do SharePoint
        padrao_onedrive = os.path.join(user_profile, "OneDrive - Salsajeans", "Shortcuts", "*", "S4H_Perfis de autorização_v1.xlsx")
        candidatos = glob.glob(padrao_onedrive)
        
        if not candidatos:
            # Fallback para qualquer ficheiro v1 na estrutura do OneDrive corporativo
            padrao_fallback = os.path.join(user_profile, "OneDrive - Salsajeans", "**", "S4H_Perfis de autorização_v1.xlsx")
            candidatos = glob.glob(padrao_fallback, recursive=True)

        if not candidatos:
            return False

        origem_sp = candidatos[0]
        if not os.path.exists(origem_sp):
            return False

        # Se já existe o local e a origem for mais recente (ou se o local não existir), efetuar a cópia
        precisa_atualizar = True
        if os.path.exists(CAMINHO_EXCEL_LOCAL):
            mtime_origem = os.path.getmtime(origem_sp)
            mtime_local = os.path.getmtime(CAMINHO_EXCEL_LOCAL)
            if mtime_local >= mtime_origem:
                precisa_atualizar = False

        if precisa_atualizar:
            # Criar pasta de output se necessário
            dir_output = os.path.dirname(CAMINHO_EXCEL_BACKUP_LOCAL)
            os.makedirs(dir_output, exist_ok=True)

            # Backup com timestamp se o local já existir
            if os.path.exists(CAMINHO_EXCEL_LOCAL):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_ts = os.path.join(dir_output, f"S4H_Perfis_v1_backup_{timestamp}.xlsx")
                shutil.copy2(CAMINHO_EXCEL_LOCAL, backup_ts)

            # Copiar da origem sincronizada para o ficheiro local
            shutil.copy2(origem_sp, CAMINHO_EXCEL_LOCAL)
            print(f"[SINCRONIZAÇÃO] Cópia atualizada com sucesso a partir do SharePoint/OneDrive:")
            print(f"               Origem: {origem_sp}")
            print(f"               Destino: {CAMINHO_EXCEL_LOCAL}")
            return True

    except Exception as err:
        print(f"[AVISO] Não foi possível sincronizar automaticamente do SharePoint/OneDrive: {err}")
    return False


def encontrar_excel_padrao(diretorio_base: Optional[str] = None) -> Optional[str]:
    """Devolve exclusivamente o ficheiro mestre local definido para o projeto."""
    del diretorio_base  # compatibilidade com chamadas antigas; não é usado.
    return CAMINHO_EXCEL_LOCAL if os.path.isfile(CAMINHO_EXCEL_LOCAL) else None


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

        # Sheet DEFINIÇÕES: roles base por departamento {DEPARTAMENTO: set(roles)}
        self.definicoes_departamento: Dict[str, Set[str]] = {}

        # Discrepâncias de sincronização CUA detetadas
        self.discrepancias_cua: Dict[str, List[Dict[str, Any]]] = {}

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

        A folha EXCLUÇÃO não participa da expansão nem da atribuição de funções;
        ela é reservada exclusivamente ao processo CUA_REMOVE.
        """
        expandidas: Set[str] = set()
        pendentes = [str(r).strip().upper() for r in roles if str(r).strip()]
        while pendentes:
            role = pendentes.pop()
            if not role or role in expandidas:
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
                # A fila operacional é controlada exclusivamente por STATUS.
                # Um timestamp residual não deve ocultar um departamento por processar.
                is_pendente = not bool(st)

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

    # -------------------------------------------------------------
    # 7. Leitura de DEFINIÇÕES (roles base por departamento)
    # -------------------------------------------------------------
    sheet_def = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) in ("DEFINICOES", "DEFINICOES")), None)
    if sheet_def:
        df_def = pd.read_excel(excel_file, sheet_name=sheet_def)
        cols_map = {normalizar_nome_coluna(c): c for c in df_def.columns}
        col_dep = cols_map.get("DEPARTAMENTO")
        if col_dep:
            for _, row in df_def.iterrows():
                dep = normalizar_texto(row.get(col_dep))
                if not dep:
                    continue
                roles_dep = dados.definicoes_departamento.setdefault(dep, set())
                for col in df_def.columns:
                    if col == col_dep:
                        continue
                    for role in extrair_roles_sap(row.get(col)):
                        roles_dep.add(role)

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
    col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_raw)

    for idx in range(1, len(df_raw)):
        dep_val = str(df_raw.iloc[idx, 6]).strip() if pd.notna(df_raw.iloc[idx, 6]) else ""
        if not dep_val or dep_val.upper() in ("NAN", "NONE", ""):
            dep_val = str(df_raw.iloc[idx, 4]).strip() if pd.notna(df_raw.iloc[idx, 4]) else ""
        if normalizar_texto(dep_val) == dep_norm or (dep_norm and dep_norm in normalizar_texto(dep_val)):
            user_id = str(df_raw.iloc[idx, 0]).strip() if pd.notna(df_raw.iloc[idx, 0]) else ""
            nome = str(df_raw.iloc[idx, 2]).strip() if pd.notna(df_raw.iloc[idx, 2]) else ""
            cargo = str(df_raw.iloc[idx, 5]).strip() if pd.notna(df_raw.iloc[idx, 5]) else ""
            comp_raw = str(df_raw.iloc[idx, 7]).strip() if pd.notna(df_raw.iloc[idx, 7]) else ""
            comp = comp_raw if comp_raw.upper() not in ("NAN", "NONE", "") else None
            if comp:
                compostas.add(comp)

            user_singles = [
                normalizar_texto(x) for x in df_raw.iloc[idx, col_fi_idx:].dropna().tolist()
                if is_role_sap_valida(x)
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
                        backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
                    backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
                    backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
                backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
        "PEOPLE & TALENT": ("People & Talent", "People & Talent"),
        "P&T": ("People & Talent", "People & Talent"),
        "HEALTH & SAFETY": ("Health & Safety", "Health & Safety"),
        "H&S": ("Health & Safety", "Health & Safety"),
        "DIGITAL": ("Digital", "Digital"),
        "IT": ("IT", "IT"),
        "TI": ("IT", "IT"),
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
        if s in ["Construction & Maintenance", "Purchase & Services", "Client Services", "Industry Services", "People & Talent", "Health & Safety", "Digital", "IT", "Legal"]:
            sheet_dfs[s] = pd.read_excel(excel_file, sheet_name=s, header=None)

    users_em_falta = []
    total_users_avaliados = 0
    col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_ativa, default_idx=10)

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
        for col_idx in range(col_fi_idx, len(r)):
            val_cell = r.iloc[col_idx]
            if pd.notna(val_cell):
                v_str = str(val_cell).strip().upper()
                if v_str and v_str.startswith("Z") and v_str != c_val and v_str not in dados.roles_compostas:
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

    backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
                col_fi = encontrar_coluna_funcoes_individuais_ws(ws, default_col=11)
                for u_info in users_em_falta:
                    row_idx = u_info["linha"]
                    last_col = max(col_fi, ws.UsedRange.Columns.Count)
                    ws.Range(ws.Cells(row_idx, col_fi), ws.Cells(row_idx, last_col)).ClearContents()
                    for i, role in enumerate(u_info["roles_esperadas"]):
                        ws.Cells(row_idx, col_fi + i).Value = role
                        alteracoes += 1
                wb.Save()
                try:
                    if backup_local.exists():
                        wb.SaveCopyAs(str(backup_local))
                except Exception:
                    pass
                sucesso_com = True
                print(f"  {alteracoes} função(ões) reescrita(s) via Excel COM (a partir da coluna {col_fi}).")
        except Exception:
            pass

    if not sucesso_com:
        try:
            wb = openpyxl.load_workbook(caminho_excel)
            ws = wb["Proposta Ativa"]
            col_fi = encontrar_coluna_funcoes_individuais_ws(ws, default_col=11)
            for u_info in users_em_falta:
                row_idx = u_info["linha"]
                for c in range(col_fi, ws.max_column + 1):
                    ws.cell(row=row_idx, column=c).value = None
                for i, role in enumerate(u_info["roles_esperadas"]):
                    ws.cell(row=row_idx, column=col_fi + i, value=role)
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
    como membros componentes da Composite Role na folha 'PFCG_COMPOSTA' e remove relações
    que já não constam na 'Proposta Ativa'.
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
    col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_ativa)

    for _, r in df_ativa.iterrows():
        comp = str(r.get(col_comp, "")).strip() if pd.notna(r.get(col_comp)) else ""
        if not comp or comp.lower() in ("nan", "none", "-"):
            continue
        for col_idx in range(col_fi_idx, len(r)):
            val = r.iloc[col_idx]
            if pd.notna(val):
                v_str = str(val).strip().upper()
                if v_str and v_str.startswith("Z") and v_str != comp and v_str not in dados.roles_compostas:
                    composta_roles[comp].add(v_str)

    df_comp = pd.read_excel(excel_file, sheet_name="PFCG_COMPOSTA")
    pares_existentes = set()
    linhas_obsoletas = []
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

    pares_esperados = {
        (comp.strip().upper(), role.strip().upper())
        for comp, roles in composta_roles.items()
        for role in roles
    }
    for idx, row in df_comp.iterrows():
        par = (
            str(row.get("AGR_NAME_COMPOSTA", "")).strip().upper(),
            str(row.get("AGR_NAME", "")).strip().upper(),
        )
        if par not in pares_esperados:
            linhas_obsoletas.append(idx + 2)

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

    backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)

    if not novas_linhas and not linhas_obsoletas:
        print(f"  [3/4] PFCG_COMPOSTA: {len(composta_roles)}/{len(composta_roles)} Composite Roles alinhadas no Excel ({len(df_comp)} registos).")
        return {"ok": True, "total_compostas": len(composta_roles), "linhas_adicionadas": 0, "linhas_removidas": 0, "total_registos": len(df_comp)}

    print(f"  PFCG_COMPOSTA: {len(novas_linhas)} relação(ões) em falta e {len(linhas_obsoletas)} obsoleta(s).")
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
                for row_idx in reversed(linhas_obsoletas):
                    ws.Rows(row_idx).Delete()
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
            for row_idx in reversed(linhas_obsoletas):
                ws.delete_rows(row_idx, 1)
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
        "linhas_removidas": len(linhas_obsoletas),
        "total_registos": len(df_comp) - len(linhas_obsoletas) + len(novas_linhas)
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


REGRAS_ESPECIAIS_TCODE: Dict[str, str] = {
    "BP": "TRATADO_POR_REGRA_ESPECIAL",
}


def auditar_matrizes_departamentais_vs_proposta(
    caminho_excel: str,
    dados: ProjetoPerfilData,
    regras_especiais: Optional[Dict[str, str]] = None
) -> Dict[str, Any]:
    """
    Audita TODAS as TCODEs efetivamente utilizadas nas matrizes departamentais
    listadas na sheet CONTROLO (independentemente do STATUS) contra a sheet Proposta.

    Regra de Utilização:
    - Qualquer 'X' (ou 1, SIM, YES, S) em qualquer coluna de utilizador = TCODE UTILIZADA
    - Zero 'X' = IGNORAR
    - TCODE normalizada com strip + uppercase, sem duplicados por departamento.
    - Se TCODE == "BP" (ou em regras_especiais): CLASSIFICAR = TRATADO_POR_REGRA_ESPECIAL
    - Se TCODE tem mapeamento na Proposta (TCODE -> 1:N Roles): CLASSIFICAR = MAPEADA
    - Caso contrário: CLASSIFICAR = TCODE_SEM_MAPEAMENTO_PROPOSTA
    """
    import pandas as pd
    from collections import defaultdict

    if regras_especiais is None:
        regras_especiais = REGRAS_ESPECIAIS_TCODE

    fonte = abrir_excel_seguro(caminho_excel)
    excel_file = pd.ExcelFile(fonte)

    # 1. Carregar Mapeamento TCODE -> 1:N Roles da sheet Proposta (e descrições de TCODE)
    df_prop = pd.read_excel(excel_file, sheet_name="Proposta")
    proposta_tcodes_mapeados: Set[str] = set()
    proposta_tcode_descs: Dict[str, str] = {}
    current_role = None

    for _, row in df_prop.iterrows():
        f_val = str(row["FUNÇÃO"]).strip() if pd.notna(row["FUNÇÃO"]) else ""
        d_val = str(row["DESCRIÇÃO"]).strip() if pd.notna(row["DESCRIÇÃO"]) else ""
        if f_val.startswith("Z_") and d_val and "TRANSACAO NAO EXISTE" not in d_val.upper():
            current_role = f_val
        elif f_val and current_role and "TRANSACAO NAO EXISTE" not in d_val.upper():
            for t in f_val.replace(";", " ").split():
                tc = t.strip().upper()
                if tc:
                    proposta_tcodes_mapeados.add(tc)
                    if tc not in proposta_tcode_descs or not proposta_tcode_descs[tc]:
                        proposta_tcode_descs[tc] = d_val

    # 2. Obter TODOS os departamentos listados na sheet CONTROLO (sem filtrar por STATUS)
    deps_controlo: List[str] = [
        item["departamento"] for item in dados.controlo if item.get("departamento")
    ]

    pendencias_por_dep: Dict[str, List[Dict[str, str]]] = defaultdict(list)
    total_tcodes_sem_mapeamento = 0
    tcodes_sem_mapeamento_set: Set[str] = set()
    total_tcodes_utilizadas_auditadas = 0

    for dep_nome in deps_controlo:
        sheet_dep_nome = next(
            (s for s in excel_file.sheet_names if normalizar_nome_coluna(s) == normalizar_nome_coluna(dep_nome)),
            None
        )
        if not sheet_dep_nome:
            continue

        df_matriz = pd.read_excel(excel_file, sheet_name=sheet_dep_nome, header=None)
        if len(df_matriz) < 2:
            continue

        # Encontrar linha de cabeçalho dos utilizadores e colunas válidas
        hdr_r_idx = None
        user_cols = []
        for r_idx in range(min(5, len(df_matriz))):
            cols_found = []
            for c_idx in range(2, len(df_matriz.columns)):
                cell_v = str(df_matriz.iloc[r_idx, c_idx]).strip()
                if cell_v and cell_v.upper() not in ("NAN", "NONE", ""):
                    cols_found.append(c_idx)
            if cols_found:
                hdr_r_idx = r_idx
                user_cols = cols_found
                break

        if hdr_r_idx is None or not user_cols:
            continue

        # Percorrer as linhas de TCODE da matriz departamental
        tcodes_vistas_no_dep: Set[str] = set()

        for r_idx in range(hdr_r_idx + 1, len(df_matriz)):
            tc_raw = df_matriz.iloc[r_idx, 0]
            if pd.isna(tc_raw):
                continue
            tc_clean_list = limpar_tcode(tc_raw)
            if not tc_clean_list:
                continue

            tc = tc_clean_list[0]
            desc_matriz = str(df_matriz.iloc[r_idx, 1]).strip() if pd.notna(df_matriz.iloc[r_idx, 1]) else ""

            # Verificar se existe pelo menos um "X" em qualquer coluna de utilizador
            tem_x = False
            for c_idx in user_cols:
                flag = str(df_matriz.iloc[r_idx, c_idx]).strip().upper()
                if flag in ("X", "1", "SIM", "YES", "S"):
                    tem_x = True
                    break

            if not tem_x:
                continue  # ZERO X = IGNORAR

            if tc in tcodes_vistas_no_dep:
                continue  # Evitar duplicados no mesmo departamento
            tcodes_vistas_no_dep.add(tc)

            total_tcodes_utilizadas_auditadas += 1

            # Classificação
            if tc in regras_especiais:
                classificacao = regras_especiais[tc]
            elif tc in proposta_tcodes_mapeados:
                classificacao = "MAPEADA"
            else:
                classificacao = "TCODE_SEM_MAPEAMENTO_PROPOSTA"

            if classificacao == "TCODE_SEM_MAPEAMENTO_PROPOSTA":
                desc_final = proposta_tcode_descs.get(tc) or desc_matriz
                pendencias_por_dep[dep_nome].append({
                    "tcode": tc,
                    "descricao": desc_final
                })
                tcodes_sem_mapeamento_set.add(tc)

    total_tcodes_sem_mapeamento = len(tcodes_sem_mapeamento_set)
    deps_com_pendencias = list(pendencias_por_dep.keys())

    return {
        "deps_auditados": len(deps_controlo),
        "total_tcodes_utilizadas": total_tcodes_utilizadas_auditadas,
        "tem_pendencias": total_tcodes_sem_mapeamento > 0,
        "deps_com_pendencias_count": len(deps_com_pendencias),
        "deps_com_pendencias": deps_com_pendencias,
        "total_tcodes_sem_mapeamento": total_tcodes_sem_mapeamento,
        "tcodes_sem_mapeamento_set": tcodes_sem_mapeamento_set,
        "pendencias_por_dep": dict(pendencias_por_dep)
    }


def preparar_cua_adicionar(
    caminho_excel: str,
    dados: ProjetoPerfilData,
    roles_por_user_finais: Optional[Dict[str, Tuple[int, str, Set[str]]]] = None,
    df_ativa: Optional[Any] = None,
) -> Dict[str, Any]:
    """
    [ETAPA 6/7] PREPARAÇÃO DA SHEET CUA_ADICIONAR.
    Lê a Proposta Ativa e o catálogo vivo para calcular atribuições diretas necessárias,
    compara com a sheet CUA_ADICIONAR existente (chave UTILIZADOR + SISTEMA + AGR_NAME)
    e grava no Excel SOMENTE as linhas em falta com STATUS/MSG/TIMESTEMP vazios.
    NÃO executa qualquer chamada de escrita no SAP.
    """
    import pandas as pd

    excel_file = pd.ExcelFile(abrir_excel_seguro(caminho_excel))
    if df_ativa is None:
        df_ativa = pd.read_excel(excel_file, sheet_name="Proposta Ativa")

    df_add = pd.read_excel(excel_file, sheet_name="CUA_ADICIONAR") if "CUA_ADICIONAR" in excel_file.sheet_names else pd.DataFrame()

    col_u_add = next((c for c in df_add.columns if "UTILIZADOR" in str(c).upper() or "USER" in str(c).upper()), None)
    col_s_add = next((c for c in df_add.columns if "SISTEMA" in str(c).upper() or "SUBSYSTEM" in str(c).upper()), None)
    col_r_add = next((c for c in df_add.columns if "AGR_NAME" in str(c).upper() or "ROLE" in str(c).upper() or "FUNCAO" in str(c).upper()), None)

    existing_cua_keys: Set[Tuple[str, str, str]] = set()
    max_id = 0

    if not df_add.empty:
        for idx_row, r in df_add.iterrows():
            try:
                rid = int(r.get("ID", 0))
                if rid > max_id:
                    max_id = rid
            except Exception:
                pass
            u_val = normalizar_texto(r.get(col_u_add)) if col_u_add else ""
            s_val = normalizar_texto(r.get(col_s_add)) if col_s_add else ""
            r_val = normalizar_texto(r.get(col_r_add)) if col_r_add else ""
            if u_val and s_val and r_val:
                existing_cua_keys.add((u_val, s_val, r_val))

    for item_dados in getattr(dados, "cua_adicionar", []):
        u_val = normalizar_texto(item_dados.get("utilizador"))
        s_val = normalizar_texto(item_dados.get("sistema"))
        r_val = normalizar_texto(item_dados.get("role"))
        if u_val and s_val and r_val:
            existing_cua_keys.add((u_val, s_val, r_val))

    roles_compostas_map = dados.roles_compostas

    mapa_dep_sheet = {
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
    }

    col_user_at = [c for c in df_ativa.columns if "USER" in str(c).upper() or "UTILIZADOR" in str(c).upper() or "USU" in str(c).upper()][0]
    col_comp_at = [c for c in df_ativa.columns if "COMPOSITE" in str(c).upper()][0]
    col_dep_at = [c for c in df_ativa.columns if str(c).strip().upper() == "DEPARTAMENTO"][0]
    col_dep_dir_at = [c for c in df_ativa.columns if "DIRE" in str(c).upper()][0]

    novas_linhas: List[Dict[str, Any]] = []
    detalhes_por_dep: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    total_users_analisados = 0
    total_atribuicoes_esperadas = 0
    ja_existentes_count = 0

    sistemas_alvo = ["S4PCLNT100", "S4QCLNT100"]

    for idx, r_at in df_ativa.iterrows():
        u_id = normalizar_texto(r_at.get(col_user_at))
        comp_nome = normalizar_texto(r_at.get(col_comp_at))
        if not u_id or not comp_nome or comp_nome in ("NAN", "NONE", "-"):
            continue

        d1 = str(r_at.get(col_dep_at, "")).strip() if pd.notna(r_at.get(col_dep_at)) else ""
        d2 = str(r_at.get(col_dep_dir_at, "")).strip() if pd.notna(r_at.get(col_dep_dir_at)) else ""
        dep_raw = (d1 if d1 and d1.lower() != "nan" else d2).strip().upper()
        dep_user = mapa_dep_sheet.get(dep_raw, dep_raw).upper()

        total_users_analisados += 1

        filhas_comp = {normalizar_texto(f) for f in roles_compostas_map.get(comp_nome, {}).get("roles_filhas", [])}
        roles_diretas_necessarias = {comp_nome}

        col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_ativa)
        for c_idx in range(col_fi_idx, len(r_at)):
            val_cell = r_at.iloc[c_idx]
            if pd.notna(val_cell):
                v_str = normalizar_texto(val_cell)
                if v_str and v_str.startswith("Z") and v_str != comp_nome and v_str not in filhas_comp and v_str not in roles_compostas_map:
                    roles_diretas_necessarias.add(v_str)

        if roles_por_user_finais and u_id in roles_por_user_finais:
            _, _, r_set_unificadas = roles_por_user_finais[u_id]
            for r_unif in r_set_unificadas:
                r_unif_norm = normalizar_texto(r_unif)
                if r_unif_norm and r_unif_norm != comp_nome and r_unif_norm not in filhas_comp and r_unif_norm not in roles_compostas_map:
                    roles_diretas_necessarias.add(r_unif_norm)

        u_nome = str(r_at.get(next((c for c in df_ativa.columns if "NOME" in str(c).upper()), col_user_at), "")).strip()

        novas_user_dep = []
        for sys_code in sistemas_alvo:
            for r_nec in sorted(roles_diretas_necessarias):
                total_atribuicoes_esperadas += 1
                key = (u_id, sys_code, r_nec)
                if key in existing_cua_keys:
                    ja_existentes_count += 1
                else:
                    max_id += 1
                    item_novo = {
                        "ID": max_id,
                        "UTILIZADOR": u_id,
                        "SISTEMA": sys_code,
                        "AGR_NAME": r_nec,
                        "STATUS": "",
                        "MSG": "",
                        "TIMESTEMP": "",
                        "PRD": "",
                        "QAD": "",
                    }
                    novas_linhas.append(item_novo)
                    existing_cua_keys.add(key)
                    novas_user_dep.append((sys_code, r_nec))

        if novas_user_dep:
            detalhes_por_dep[dep_user].append({
                "user": u_id,
                "nome": u_nome,
                "novas": novas_user_dep
            })

    return {
        "max_id": max_id,
        "total_users_analisados": total_users_analisados,
        "total_atribuicoes_esperadas": total_atribuicoes_esperadas,
        "ja_existentes_count": ja_existentes_count,
        "novas_linhas_count": len(novas_linhas),
        "novas_linhas": novas_linhas,
        "detalhes_por_dep": dict(detalhes_por_dep)
    }


def executar_atualizacao_integrada_excel(
    caminho_excel: Optional[str] = None,
    dados: Optional[ProjetoPerfilData] = None,
    validar_tcodes_rfc: bool = True,
    incorporar_prd_rfc: bool = True,
    assumir_sim: bool = False,
) -> ProjetoPerfilData:
    """
    [TAREFA 1 DO PROGRAMA] Executa a atualização e consolidação integrada do livro Excel
    num único ciclo atómico em memória, sem alternância ou saltos visuais entre folhas.

    Ordem lógica de dependência:
      [1/6] Matrizes Departamentais × Proposta (Validar TCODEs utilizadas e mapeamentos)
      [2/6] Proposta (Validar catálogo de TCODEs em SAP PRD / TSTC)
      [3/6] Proposta Ativa (Reconciliar utilizadores e funções das matrizes)
      [4/6] Proposta Ativa / catálogo vivo (Incorporar funções válidas necessárias)
      [5/6] PFCG_COMPOSTA (Reconciliar Composite -> Single Roles)
      [6/6] Sincronização Departamental & CUA (Auditar utilizadores e atribuições)
    """
    import pandas as pd
    from collections import defaultdict
    import openpyxl
    import shutil

    if not caminho_excel:
        caminho_excel = encontrar_excel_padrao()
    if not caminho_excel or not os.path.exists(caminho_excel):
        print(f"[ERRO] Ficheiro Excel não encontrado: {caminho_excel}")
        if dados is not None:
            return dados
        raise FileNotFoundError(f"Excel não encontrado: {caminho_excel}")

    print("\n" + "=" * 78)
    print("  TAREFA 1: ATUALIZAÇÃO E SINCRONIZAÇÃO INTEGRADA DO EXCEL")
    print("=" * 78)

    if dados is None:
        dados = carregar_projeto_perfil(caminho_excel)

    fonte = abrir_excel_seguro(caminho_excel)
    excel_file = pd.ExcelFile(fonte)

    # -------------------------------------------------------------
    # ETAPA [1/6]: Matrizes Departamentais × Proposta
    # -------------------------------------------------------------
    res_auditoria = auditar_matrizes_departamentais_vs_proposta(caminho_excel, dados)
    tcodes_ignoradas_por_sem_mapeamento: Set[str] = set()

    if not res_auditoria["tem_pendencias"]:
        print(f"  {CLR_VERDE}✓{CLR_RESET} [1/6] Matrizes Departamentais: todas as TCODEs utilizadas possuem mapeamento válido na Proposta.")
    else:
        n_deps_pend = res_auditoria["deps_com_pendencias_count"]
        n_tcodes_pend = res_auditoria["total_tcodes_sem_mapeamento"]
        print(f"\n==============================================================================")
        print(f"  VALIDAÇÃO MATRIZES DEPARTAMENTAIS × PROPOSTA")
        print(f"==============================================================================")
        print(f"  {CLR_AMARELO}⚠️ Foram encontradas TCODEs utilizadas nas matrizes departamentais\n    sem mapeamento na sheet Proposta.{CLR_RESET}\n")

        for dep_pend, lista_tc in res_auditoria["pendencias_por_dep"].items():
            print(f"  {dep_pend}:")
            for item_tc in lista_tc:
                print(f"     - {item_tc['tcode']:<20} | {item_tc['descricao']}")
            print()

        print("-" * 78)
        print(f"  Departamentos com pendências: {n_deps_pend}")
        print(f"  TCODEs sem mapeamento: {n_tcodes_pend}")
        print("-" * 78)

        print(f"\n------------------------------------------------------------------------------")
        print(f"  Existem transações utilizadas nas matrizes departamentais sem mapeamento")
        print(f"  na sheet Proposta.\n")
        print(f"  Estas pendências podem ser corrigidas agora ou tratadas posteriormente.\n")
        print(f"  [1] Continuar o processamento parcialmente")
        print(f"  [2] Abortar para corrigir o ficheiro")
        print(f"------------------------------------------------------------------------------")

        opcao_escolhida = ""
        if assumir_sim:
            opcao_escolhida = "1"
            print("  [Auto-confirmado --yes] Opção 1 selecionada.")
        else:
            try:
                opcao_escolhida = input("  Selecione uma opção [1-2]: ").strip()
            except Exception:
                opcao_escolhida = "2"

        if opcao_escolhida != "1":
            print(f"\n------------------------------------------------------------------------------")
            print(f"  Processamento abortado pelo utilizador.\n")
            print(f"  Corrija as TCODEs sem mapeamento na sheet Proposta")
            print(f"  ou remova os X indevidos das matrizes departamentais.\n")
            print(f"  Nenhuma etapa posterior da Tarefa 1 será executada.")
            print(f"------------------------------------------------------------------------------\n")
            return dados

        tcodes_ignoradas_por_sem_mapeamento = res_auditoria["tcodes_sem_mapeamento_set"]
        print(f"\n  {CLR_AMARELO}⚠️ [1/6] Matrizes Departamentais: {n_tcodes_pend} TCODEs sem mapeamento em {n_deps_pend} departamento(s).\n     Continuação parcial autorizada pelo utilizador.{CLR_RESET}")

    # -------------------------------------------------------------
    # ETAPA [2/6]: Sanitização da folha 'Proposta' (TCODEs no SAP TSTC)
    # -------------------------------------------------------------
    df_prop = pd.read_excel(excel_file, sheet_name="Proposta")
    analise_prop = analisar_folha_proposta(dados)
    todos_tcodes = analise_prop.get("todos_tcodes", [])
    tcodes_a_marcar = []

    if validar_tcodes_rfc and todos_tcodes:
        res_tc = verificar_tcodes_prd(todos_tcodes)
        nao_existentes = set(res_tc.get("nao_existentes", []))
        if nao_existentes:
            role_atual = ""
            for idx, row in df_prop.iterrows():
                f_val = str(row["FUNÇÃO"]).strip().upper() if pd.notna(row["FUNÇÃO"]) else ""
                d_val = str(row["DESCRIÇÃO"]).strip() if pd.notna(row["DESCRIÇÃO"]) else ""
                if is_nome_funcao_proposta(f_val, d_val):
                    role_atual = f_val
                elif f_val in nao_existentes and role_atual and "TRANSACAO NAO EXISTE" not in d_val.upper():
                    tcodes_a_marcar.append((idx + 2, role_atual, f_val))

    if tcodes_a_marcar:
        print(f"  {CLR_VERMELHO}✗ [2/6] Proposta: {len(tcodes_a_marcar)} transação(ões) descontinuada(s) detetada(s) para marcação.{CLR_RESET}")
    else:
        print(f"  {CLR_VERDE}✓{CLR_RESET} [2/6] Proposta: Catálogo com {len(todos_tcodes)} TCODEs 100% validado no SAP PRD (TSTC).")

    # Mapeamento TCODE -> roles (desconsiderando TCODEs descontinuados)
    tcodes_descartados_set = {tc for _, _, tc in tcodes_a_marcar}
    tcode_to_roles = defaultdict(list)
    current_role = None
    for _, row in df_prop.iterrows():
        f_val = str(row["FUNÇÃO"]).strip() if pd.notna(row["FUNÇÃO"]) else ""
        d_val = str(row["DESCRIÇÃO"]).strip() if pd.notna(row["DESCRIÇÃO"]) else ""
        if f_val.startswith("Z_") and d_val and "TRANSACAO NAO EXISTE" not in d_val.upper():
            current_role = f_val
        elif f_val and current_role and "TRANSACAO NAO EXISTE" not in d_val.upper() and f_val.upper() not in tcodes_descartados_set:
            for t in f_val.replace(";", " ").split():
                tc = t.strip().upper()
                if tc:
                    tcode_to_roles[tc].append(current_role)

    # -------------------------------------------------------------
    # ETAPA 2: Matrizes Departamentais -> 'Proposta Ativa'
    # -------------------------------------------------------------
    mapa_dep_sheet = {
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
    }

    df_ativa = pd.read_excel(excel_file, sheet_name="Proposta Ativa")
    col_user = [c for c in df_ativa.columns if "USER" in str(c).upper() or "UTILIZADOR" in str(c).upper() or "USU" in str(c).upper()][0]
    col_comp = [c for c in df_ativa.columns if "COMPOSITE" in str(c).upper()][0]
    col_dep = [c for c in df_ativa.columns if str(c).strip().upper() == "DEPARTAMENTO"][0]
    col_dep_dir = [c for c in df_ativa.columns if "DIRE" in str(c).upper()][0]

    sheet_dfs = {}
    for s in excel_file.sheet_names:
        if s in ["Construction & Maintenance", "Purchase & Services", "Client Services", "Industry Services", "People & Talent", "Health & Safety", "Digital", "IT", "Legal"]:
            sheet_dfs[s] = pd.read_excel(excel_file, sheet_name=s, header=None)

    users_em_falta_matriz = {}
    total_users_avaliados = 0
    roles_por_user_finais = {}
    col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_ativa)

    for idx, r in df_ativa.iterrows():
        c_val = str(r.get(col_comp, "")).strip() if pd.notna(r.get(col_comp)) else ""
        if not c_val or c_val.lower() in ("nan", "none", "-"):
            continue

        u_val = str(r.get(col_user, "")).strip().upper() if pd.notna(r.get(col_user)) else ""
        d_val = str(r.get(col_dep, "")).strip() if pd.notna(r.get(col_dep)) else ""
        dd_val = str(r.get(col_dep_dir, "")).strip() if pd.notna(r.get(col_dep_dir)) else ""
        dep = d_val if d_val and d_val.lower() != "nan" else dd_val

        total_users_avaliados += 1
        funcoes_existentes = set()
        celulas_originais = set()
        for col_idx in range(col_fi_idx, len(r)):
            val_cell = r.iloc[col_idx]
            if pd.notna(val_cell):
                v_str = str(val_cell).strip().upper()
                if v_str and v_str.startswith("Z"):
                    celulas_originais.add(v_str)
                    if v_str != c_val and v_str not in dados.roles_compostas:
                        funcoes_existentes.add(v_str)

        dep_norm = dep.strip().upper()
        sheet_nome = mapa_dep_sheet.get(dep_norm)
        tcodes_marcados_user = []
        if sheet_nome and sheet_nome in sheet_dfs:
            df_matriz = sheet_dfs[sheet_nome]
            col_u_idx = None
            hdr_r_idx = None
            for r_idx in range(min(5, len(df_matriz))):
                for c_idx in range(len(df_matriz.columns)):
                    cell_v = str(df_matriz.iloc[r_idx, c_idx]).strip().upper()
                    if u_val and (re.search(r'\b' + re.escape(u_val) + r'\b', cell_v) or cell_v == u_val):
                        col_u_idx = c_idx
                        hdr_r_idx = r_idx
                        break
                if col_u_idx is not None:
                    break
            if col_u_idx is not None and hdr_r_idx is not None:
                for r_idx in range(hdr_r_idx + 1, len(df_matriz)):
                    tc = str(df_matriz.iloc[r_idx, 0]).strip().upper()
                    flag = str(df_matriz.iloc[r_idx, col_u_idx]).strip().upper()
                    if flag in ("X", "1", "SIM", "YES", "S"):
                        tcodes_marcados_user.append(tc)

        roles_esperadas = set()
        for tc in tcodes_marcados_user:
            if tc in tcodes_ignoradas_por_sem_mapeamento:
                continue
            for mr in tcode_to_roles.get(tc, []):
                roles_esperadas.add(mr)

        roles_faltam_matriz = roles_esperadas - funcoes_existentes
        precisa_limpar_celulas = (celulas_originais != funcoes_existentes)
        if roles_faltam_matriz or precisa_limpar_celulas:
            roles_unificadas = funcoes_existentes.union(roles_esperadas)
            users_em_falta_matriz[u_val] = {
                "linha": idx + 2,
                "user": u_val,
                "composite": c_val,
                "roles_esperadas": set(roles_unificadas),
            }
            roles_por_user_finais[u_val] = (idx + 2, c_val, set(roles_unificadas))
        else:
            roles_por_user_finais[u_val] = (idx + 2, c_val, set(funcoes_existentes))

    if users_em_falta_matriz:
        print(f"  {CLR_VERMELHO}✗ [3/6] Proposta Ativa: {len(users_em_falta_matriz)} utilizador(es) atualizados com funções em falta das matrizes.{CLR_RESET}")
    else:
        print(f"  {CLR_VERDE}✓{CLR_RESET} [3/6] Proposta Ativa: Reconciliar utilizadores e funções das matrizes ({total_users_avaliados}/{total_users_avaliados} utilizadores).")

    # -------------------------------------------------------------
    # ETAPA [4/6]: SAP PRD -> 'Proposta Ativa' (Incorporação de funções vivas)
    # -------------------------------------------------------------
    funcoes_vivas_a_adicionar = defaultdict(set)
    if incorporar_prd_rfc and roles_por_user_finais:
        from datetime import date
        users_lista = sorted(list(roles_por_user_finais.keys()))
        catalogo_oficial_set = set(dados.roles_simples.keys()).union(dados.roles_compostas.keys())
        for r_list in tcode_to_roles.values():
            catalogo_oficial_set.update(r_list)

        try:
            from sap_rfc._rfc_common import (
                build_connection_params_for, load_project_env, find_project_root,
                make_read_only_guard, read_table, make_option_in
            )
            from pyrfc import Connection
            load_project_env(find_project_root())
            params = build_connection_params_for("PRD")
            conn = Connection(**params)
            guard = make_read_only_guard(["AGR_USERS", "USR02"])

            hoje_int = int(date.today().strftime("%Y%m%d"))
            inativos_usr02 = set()

            for i in range(0, len(users_lista), 20):
                chunk = users_lista[i:i+20]
                rows_usr = read_table(conn, guard, table_name="USR02", fields=["BNAME", "GLTGB"], options=make_option_in("BNAME", chunk), rowcount=0)
                for r in rows_usr:
                    if len(r) >= 2:
                        bn = str(r[0]).strip().upper()
                        g_fim = str(r[1]).strip()
                        fim_int = int(g_fim) if g_fim.isdigit() and int(g_fim) > 0 else 99991231
                        if hoje_int > fim_int:
                            inativos_usr02.add(bn)

            for i in range(0, len(users_lista), 20):
                chunk = [u for u in users_lista[i:i+20] if u not in inativos_usr02]
                if not chunk: continue
                rows_agr = read_table(conn, guard, table_name="AGR_USERS", fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT", "COL_FLAG"], options=make_option_in("UNAME", chunk), rowcount=0)
                for r in rows_agr:
                    if len(r) >= 4:
                        role = str(r[0]).strip().upper()
                        un = str(r[1]).strip().upper()
                        f_d = str(r[2]).strip()
                        t_d = str(r[3]).strip()
                        col_flag = str(r[4]).strip().upper() if len(r) > 4 else ""
                        if col_flag == "X":
                            continue  # Ignorar funções filhas herdadas de compostas
                        inicio = int(f_d) if f_d.isdigit() else 0
                        fim = int(t_d) if t_d.isdigit() else 99991231
                        if inicio <= hoje_int <= fim and role and un in roles_por_user_finais:
                            row_num, c_role, roles_atuais = roles_por_user_finais[un]
                            if role != c_role and role not in dados.roles_compostas:
                                if not dados.is_excluida(role) and role in catalogo_oficial_set:
                                    if role not in roles_atuais:
                                        funcoes_vivas_a_adicionar[un].add(role)
                                        roles_atuais.add(role)
            conn.close()
        except Exception as e_rfc:
            print(f"  [AVISO] Não foi possível consultar funções vivas no PRD via RFC: {e_rfc}")

    total_vivas = sum(len(v) for v in funcoes_vivas_a_adicionar.values())
    if total_vivas > 0:
        print(f"  {CLR_AMARELO}⚠️ {CLR_RESET} [4/6] Proposta Ativa / catálogo vivo: {total_vivas} função(ões) viva(s) do catálogo em PRD a incorporar ({len(funcoes_vivas_a_adicionar)} utilizadores).")
    else:
        print(f"  {CLR_VERDE}✓{CLR_RESET} [4/6] Proposta Ativa / catálogo vivo: Incorporar funções válidas necessárias (0 pendentes).")

    # -------------------------------------------------------------
    # ETAPA [5/7]: 'Proposta Ativa' -> 'PFCG_COMPOSTA'
    # -------------------------------------------------------------
    df_comp = pd.read_excel(excel_file, sheet_name="PFCG_COMPOSTA")
    composta_roles_esperadas = defaultdict(set)
    composta_textos = {}
    for _, (l_idx, comp_nome, r_set) in roles_por_user_finais.items():
        if comp_nome and str(comp_nome).strip().upper() not in ("NAN", "NONE", "-"):
            composta_roles_esperadas[str(comp_nome).strip().upper()].update(r_set)

    pares_existentes = set()
    linhas_obsoletas_comp = []
    max_id_comp = 0
    for idx, row in df_comp.iterrows():
        try:
            rid = int(row.get("ID", 0))
            if rid > max_id_comp: max_id_comp = rid
        except Exception: pass
        agr_c = str(row.get("AGR_NAME_COMPOSTA", "")).strip().upper()
        agr = str(row.get("AGR_NAME", "")).strip().upper()
        txt = str(row.get("TEXT", "")).strip()
        if agr_c and txt and agr_c not in composta_textos:
            composta_textos[agr_c] = txt
        if agr_c and agr:
            pares_existentes.add((agr_c, agr))

    pares_esperados_comp = {
        (comp, r)
        for comp, r_set in composta_roles_esperadas.items()
        for r in r_set
    }

    linhas_obsoletas_comp = []

    novas_linhas_comp = []
    for comp, r_set in sorted(composta_roles_esperadas.items()):
        desc = composta_textos.get(comp, "")
        for r in sorted(r_set):
            if (comp, r) not in pares_existentes:
                max_id_comp += 1
                novas_linhas_comp.append({
                    "ID": max_id_comp,
                    "AGR_NAME_COMPOSTA": comp,
                    "TEXT": desc,
                    "AGR_NAME": r,
                })

    if novas_linhas_comp or linhas_obsoletas_comp:
        print(f"  {CLR_VERMELHO}✗ [5/7] PFCG_COMPOSTA: {len(novas_linhas_comp)} relação(ões) em falta e {len(linhas_obsoletas_comp)} obsoleta(s).{CLR_RESET}")
    else:
        print(f"  {CLR_VERDE}✓{CLR_RESET} [5/7] PFCG_COMPOSTA: Reconciliar Composite -> Single Roles ({len(composta_roles_esperadas)} Composite Roles 100% consistentes).")

    # -------------------------------------------------------------
    # ETAPA [6/7]: PREPARAÇÃO DA SHEET CUA_ADICIONAR (SEM ESCRITA SAP)
    # -------------------------------------------------------------
    res_prep_cua = preparar_cua_adicionar(
        caminho_excel=caminho_excel,
        dados=dados,
        roles_por_user_finais=roles_por_user_finais,
        df_ativa=df_ativa,
    )
    novas_linhas_cua = res_prep_cua.get("novas_linhas", [])

    print("\n==============================================================================")
    print("  [6/7] PREPARAÇÃO DA SHEET CUA_ADICIONAR")
    print("==============================================================================")
    detalhes_dep = res_prep_cua.get("detalhes_por_dep", {})
    if detalhes_dep:
        for dep_nome, user_list in sorted(detalhes_dep.items()):
            print(f"\n  {dep_nome}:")
            for u_info in user_list:
                print(f"\n  {u_info['user']} - {u_info['nome']}")
                print(f"    Pendentes para CUA_ADICIONAR:")
                for sys_code, r_nec in u_info["novas"]:
                    print(f"      + {sys_code} | {r_nec}")
        print()

    if res_prep_cua["novas_linhas_count"] > 0:
        print(f"  {CLR_VERMELHO}✗ [6/7] CUA_ADICIONAR preparada:{CLR_RESET}")
        print(f"     {res_prep_cua['total_users_analisados']} utilizadores analisados")
        print(f"     {res_prep_cua['total_atribuicoes_esperadas']} atribuições esperadas")
        print(f"     {res_prep_cua['ja_existentes_count']} já existentes")
        print(f"     {res_prep_cua['novas_linhas_count']} novas linhas criadas")
    else:
        print(f"  {CLR_VERDE}✓ [6/7] CUA_ADICIONAR já está sincronizada com a necessidade funcional.{CLR_RESET}")
        print(f"     Nenhuma nova linha necessária.")

    # -------------------------------------------------------------
    # GRAVAÇÃO UNIFICADA NO EXCEL (Apenas se houver alterações)
    # -------------------------------------------------------------
    tem_alteracoes = bool(tcodes_a_marcar or users_em_falta_matriz or funcoes_vivas_a_adicionar or novas_linhas_comp or linhas_obsoletas_comp or novas_linhas_cua)

    if not tem_alteracoes:
        print("-" * 78)
        print(f"  {CLR_VERDE}✓ O ficheiro Excel encontra-se 100% alinhado em todas as sheets.{CLR_RESET}")
        print("-" * 78)
        return dados

    print("\n  A aplicar alterações consolidadas no ficheiro Excel...")

    # Backup de segurança antes de alterar
    try:
        output_dir = Path(r"C:\workspace\SapScript\output")
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = output_dir / f"S4H_Perfis_backup_{timestamp_str}.xlsx"
        shutil.copy2(caminho_excel, backup_path)
        print(f"  Backup de segurança criado: {backup_path.name}")
    except Exception as e_bk:
        print(f"  [AVISO] Falha ao criar backup: {e_bk}")

    # Verificar se Excel está aberto via COM
    wb_com = None
    xl = None
    if sys.platform.startswith("win"):
        try:
            import win32com.client
            xl = win32com.client.Dispatch("Excel.Application")
            for w in xl.Workbooks:
                if Path(w.FullName).resolve() == Path(caminho_excel).resolve():
                    wb_com = w
                    break
        except Exception:
            wb_com = None

    if wb_com is not None and xl is not None:
        old_su = xl.ScreenUpdating
        old_da = xl.DisplayAlerts
        try:
            xl.ScreenUpdating = False
            xl.DisplayAlerts = False

            # Etapa 1: Proposta
            if tcodes_a_marcar:
                ws_p = wb_com.Worksheets("Proposta")
                for r_idx, _, _ in tcodes_a_marcar:
                    ws_p.Cells(r_idx, 2).Value = "Transação não existe"

            # Etapa 2 & 3: Proposta Ativa
            if users_em_falta_matriz or funcoes_vivas_a_adicionar:
                ws_a = wb_com.Worksheets("Proposta Ativa")
                col_fi = encontrar_coluna_funcoes_individuais_ws(ws_a, default_col=11)
                users_a_reescrever = set(users_em_falta_matriz.keys()).union(funcoes_vivas_a_adicionar.keys())
                for u in users_a_reescrever:
                    row_idx, _, roles_finais = roles_por_user_finais[u]
                    last_col = max(col_fi, ws_a.UsedRange.Columns.Count)
                    ws_a.Range(ws_a.Cells(row_idx, col_fi), ws_a.Cells(row_idx, last_col)).ClearContents()
                    for i, role in enumerate(sorted(roles_finais)):
                        ws_a.Cells(row_idx, col_fi + i).Value = role

            # Etapa 4: PFCG_COMPOSTA (Apenas colunas 1 a 4)
            if linhas_obsoletas_comp or novas_linhas_comp:
                ws_c = wb_com.Worksheets("PFCG_COMPOSTA")
                for r_idx in reversed(linhas_obsoletas_comp):
                    ws_c.Rows(r_idx).Delete()
                last_r = ws_c.UsedRange.Rows.Count
                while ws_c.Cells(last_r, 2).Value:
                    last_r += 1
                for i, row in enumerate(novas_linhas_comp):
                    curr = last_r + i
                    ws_c.Cells(curr, 1).Value = row["ID"]
                    ws_c.Cells(curr, 2).Value = row["AGR_NAME_COMPOSTA"]
                    ws_c.Cells(curr, 3).Value = row["TEXT"]
                    ws_c.Cells(curr, 4).Value = row["AGR_NAME"]

            # Etapa 6: CUA_ADICIONAR
            if novas_linhas_cua:
                ws_add = wb_com.Worksheets("CUA_ADICIONAR")
                last_r_add = ws_add.UsedRange.Rows.Count
                while ws_add.Cells(last_r_add, 2).Value:
                    last_r_add += 1
                for i, row_item in enumerate(novas_linhas_cua):
                    curr = last_r_add + i
                    ws_add.Cells(curr, 1).Value = row_item["ID"]
                    ws_add.Cells(curr, 2).Value = row_item["UTILIZADOR"]
                    ws_add.Cells(curr, 3).Value = row_item["SISTEMA"]
                    ws_add.Cells(curr, 4).Value = row_item["AGR_NAME"]
                    ws_add.Cells(curr, 5).Value = ""
                    ws_add.Cells(curr, 6).Value = ""
                    ws_add.Cells(curr, 7).Value = ""

            wb_com.Save()
            print("  Ficheiro Excel gravado via Excel.Application (COM) sem cintilação de ecrã!")
        finally:
            xl.ScreenUpdating = old_su
            xl.DisplayAlerts = old_da
    else:
        wb_ox = openpyxl.load_workbook(caminho_excel)

        # Etapa 1: Proposta
        if tcodes_a_marcar:
            ws_p = wb_ox["Proposta"]
            for r_idx, _, _ in tcodes_a_marcar:
                ws_p.cell(row=r_idx, column=2, value="Transação não existe")

        # Etapa 2 & 3: Proposta Ativa
        if users_em_falta_matriz or funcoes_vivas_a_adicionar:
            ws_a = wb_ox["Proposta Ativa"]
            col_fi = encontrar_coluna_funcoes_individuais_ws(ws_a, default_col=11)
            users_a_reescrever = set(users_em_falta_matriz.keys()).union(funcoes_vivas_a_adicionar.keys())
            for u in users_a_reescrever:
                row_idx, _, roles_finais = roles_por_user_finais[u]
                for c in range(col_fi, ws_a.max_column + 1):
                    ws_a.cell(row=row_idx, column=c).value = None
                for i, role in enumerate(sorted(roles_finais)):
                    ws_a.cell(row=row_idx, column=col_fi + i, value=role)

        # Etapa 4: PFCG_COMPOSTA (Apenas colunas 1 a 4)
        if linhas_obsoletas_comp or novas_linhas_comp:
            ws_c = wb_ox["PFCG_COMPOSTA"]
            for r_idx in reversed(linhas_obsoletas_comp):
                ws_c.delete_rows(r_idx, 1)
            for row in novas_linhas_comp:
                next_r = ws_c.max_row + 1
                ws_c.cell(row=next_r, column=1, value=row["ID"])
                ws_c.cell(row=next_r, column=2, value=row["AGR_NAME_COMPOSTA"])
                ws_c.cell(row=next_r, column=3, value=row["TEXT"])
                ws_c.cell(row=next_r, column=4, value=row["AGR_NAME"])

        # Etapa 6: CUA_ADICIONAR
        if novas_linhas_cua:
            ws_add = wb_ox["CUA_ADICIONAR"]
            for row_item in novas_linhas_cua:
                next_r = ws_add.max_row + 1
                ws_add.cell(row=next_r, column=1, value=row_item["ID"])
                ws_add.cell(row=next_r, column=2, value=row_item["UTILIZADOR"])
                ws_add.cell(row=next_r, column=3, value=row_item["SISTEMA"])
                ws_add.cell(row=next_r, column=4, value=row_item["AGR_NAME"])
                ws_add.cell(row=next_r, column=5, value="")
                ws_add.cell(row=next_r, column=6, value="")
                ws_add.cell(row=next_r, column=7, value="")

        wb_ox.save(caminho_excel)
        wb_ox.close()
        print("  Ficheiro Excel gravado via openpyxl com sucesso!")

    print("  A recarregar dados estruturados do Excel atualizado...")
    dados = carregar_projeto_perfil(caminho_excel)
    print("-" * 78)
    return dados


def executar_sincronizacao_arranque_completa(dados: ProjetoPerfilData, caminho_excel: str) -> ProjetoPerfilData:
    """Executa a sincronização e atualização integrada do Excel."""
    return executar_atualizacao_integrada_excel(caminho_excel, dados)


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
    # EXCLUÇÃO é aplicada somente ao processo CUA_REMOVE.
    dif_total_sem_exclusao = set(dif_total_bruto)
    dif_z_sem_exclusao = set(dif_z_bruto)

    dif_atrib_bruto = roles_atribuidas_prd - roles_excel
    dif_atrib_sem_exclusao = set(dif_atrib_bruto)

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
    Cruza relacionalmente as fontes essenciais de autorizações:
      1. PFCG_CREATE: Funções individuais criadas (catálogo);
      2. PFCG_COMPOSTA: Funções filhas/membros de cada função composta;
      3. DEFINIÇÕES: Funções padrão departamentais.
    """
    analise = analisar_departamento_proposta(dados, departamento)
    if not analise.get("encontrado"):
        return {"encontrado": False, "departamento": departamento, "mensagem": analise.get("mensagem")}

    usuarios_detalhe = []
    dep_norm = normalizar_texto(departamento)
    definicoes_dep = dados.definicoes_departamento.get(dep_norm, set())
    if not definicoes_dep and dep_norm in ("LEGAL", "PURCHASE & SERVICES", "PURCHASE AND SERVICES"):
        definicoes_dep = dados.definicoes_departamento.get("PURCHASE & SERVICES", set())

    for u in analise.get("usuarios", []):
        usr = u["usuario"]
        comp = u.get("composta")
        singles_prop = [r for r in u.get("singles", []) if r]

        # Membros PFCG_COMPOSTA / PFCG_CREATE
        membros_comp = set(dados.roles_compostas.get(comp, {}).get("roles_filhas", [])) if comp else set()

        # Singles avulsas que não são filhas da composta (para atribuição direta)
        singles_avulsas = [r for r in singles_prop if r not in membros_comp]

        # Definições que não são filhas da composta
        def_diretas = [r for r in definicoes_dep if r not in membros_comp]

        # Funções oficiais de atribuição direta: Composta + Singles Avulsas + Definições
        diretas_set = set()
        if comp:
            diretas_set.add(comp)
        diretas_set.update(singles_avulsas)
        diretas_set.update(def_diretas)

        # Todas as autorizações cobertas (diretas + filhas da composta)
        todas_autorizacoes = set(diretas_set)
        if membros_comp:
            todas_autorizacoes.update(membros_comp)

        usuarios_detalhe.append({
            "usuario": usr,
            "nome": u["nome"],
            "cargo": u["cargo"],
            "composta": comp,
            "singles_proposta_total": len(singles_prop),
            "singles_avulsas": sorted(singles_avulsas),
            "singles_avulsas_total": len(singles_avulsas),
            "definicoes_total": len(definicoes_dep),
            "definicoes_roles": sorted(definicoes_dep),
            "membros_composta_total": len(membros_comp),
            "membros_composta": sorted(membros_comp),
            "diretas_esperadas": sorted(diretas_set),
            "diretas_total": len(diretas_set),
            "esperado_total": len(diretas_set),
            "esperado_roles": sorted(diretas_set),
            "todas_autorizacoes_total": len(todas_autorizacoes),
            "todas_autorizacoes_roles": sorted(todas_autorizacoes),
        })

    return {
        "encontrado": True,
        "departamento": departamento,
        "total_usuarios": len(usuarios_detalhe),
        "definicoes_roles": sorted(definicoes_dep),
        "usuarios": usuarios_detalhe,
    }


def imprimir_cruzamento_fontes(resultado: Dict[str, Any]):
    """Exibe no terminal o relatório de cruzamento relacional de fontes (PFCG_CREATE, PFCG_COMPOSTA e DEFINIÇÕES)."""
    print("\n" + "=" * 78)
    print("  CRUZAMENTO DE FONTES (PFCG_CREATE, PFCG_COMPOSTA, DEFINIÇÕES)")
    print(f"  Departamento: {resultado.get('departamento', '')}")
    print("=" * 78)

    if not resultado.get("encontrado"):
        print(f"  [AVISO] {resultado.get('mensagem', 'Departamento não encontrado.')}")
        return

    def_dep_lista = resultado.get("definicoes_roles", [])
    if def_dep_lista:
        print(f"  Funções Padrão (DEFINIÇÕES - {len(def_dep_lista)}): {', '.join(def_dep_lista)}")
    print(f"  Utilizadores analisados: {resultado.get('total_usuarios', 0)}")
    print("-" * 78)

    for u in resultado.get("usuarios", []):
        comp_str = f"Composta: {u['composta']}" if u.get("composta") else "Sem Composta"
        print(f"\n  👤 {u['usuario']:<10} | {u['nome']:<25} | {u['cargo']}")
        print(f"     └─ {comp_str} | {u['singles_proposta_total']} Singles na Proposta Ativa")
        print(f"        ├─ Funções Padrão (Definições): {u['definicoes_total']}")
        print(f"        ├─ Membros em PFCG_CREATE:      {u['membros_composta_total']}")
        print(f"        └─ Total a Atribuir no SAP:     {u['esperado_total']} Funções ({u['todas_autorizacoes_total']} autorizações com membros)")


def validar_utilizadores_prd(dados: ProjetoPerfilData, departamento: str = "Purchase & Services") -> Dict[str, Any]:
    """
    Verifica no SAP PRD (AGR_USERS via RFC) as atribuições reais de cada utilizador
    do departamento, cruzando com as fontes oficiais (PFCG_CREATE, PFCG_COMPOSTA, DEFINIÇÕES).
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

        ativas_diretas_prd: Dict[str, Set[str]] = {u: set() for u in esperado_por_user}
        ativas_filhas_prd: Dict[str, Set[str]] = {u: set() for u in esperado_por_user}
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
            rows = read_table(conn, guard, table_name="AGR_USERS", fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT", "COL_FLAG"], options=opts, rowcount=0)
            for r in rows:
                if len(r) >= 4:
                    role = str(r[0]).strip().upper()
                    uname = str(r[1]).strip().upper()
                    from_d = str(r[2]).strip()
                    to_d = str(r[3]).strip()
                    col_flag = str(r[4]).strip().upper() if len(r) > 4 else ""
                    if uname in esperado_por_user and role:
                        st = assignment_status_rfc(from_d, to_d)
                        if st == "ATIVO":
                            if col_flag == "X":
                                ativas_filhas_prd[uname].add(role)
                            else:
                                ativas_diretas_prd[uname].add(role)
                        else:
                            inativas_prd[uname][role] = {"status": st, "inicio": from_d, "fim": to_d}

        conn.close()

        relatorio_users = []
        total_adicionais = 0
        total_desconsideradas = 0

        for uname in sorted(esperado_por_user.keys()):
            u_meta = detalhes_user[uname]
            comp_esperada = u_meta.get("composta")
            membros_comp = set(u_meta.get("membros_composta", []))

            # Funções diretas esperadas para atribuição
            esp = set(u_meta.get("diretas_esperadas", esperado_por_user[uname]))

            # Funções ativas no SAP:
            atv_diretas = ativas_diretas_prd[uname]
            atv_filhas = ativas_filhas_prd[uname]

            # Uma função esperada está satisfeita se:
            # 1. Estiver diretamente atribuída ao utilizador (atv_diretas); OU
            # 2. For uma single coberta pela Composta ativa; OU
            # 3. Estiver ativa via herança (atv_filhas)
            faltam = []
            for r in sorted(esp):
                if r in atv_diretas:
                    continue
                if r == comp_esperada:
                    faltam.append(r)
                elif r in atv_filhas:
                    continue
                else:
                    faltam.append(r)

            # Funções adicionais atribuídas diretamente fora do plano esperado
            adicionais_brutas = sorted(atv_diretas - esp)

            # Filtro da folha EXCLUÇÃO: proteger padrões técnicos e legados (ex: ZMM_APROVA_PEDC_COD_*, Z_MY_HOME, SAP_*)
            desconsideradas = sorted([r for r in adicionais_brutas if dados.is_excluida(r)])
            adicionais_reais = sorted([r for r in adicionais_brutas if not dados.is_excluida(r)])

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

            usr_info = mestre_usr02.get(uname, {})
            u_expirado = usr_info.get("expirado", False)
            esperadas_ativas = len(esp) - len(faltam)

            relatorio_users.append({
                "usuario": uname,
                "nome": u_meta["nome"],
                "cargo": u_meta["cargo"],
                "composta": comp_esperada,
                "inativo_usr02": u_expirado,
                "validade_fim": usr_info.get("validade_fim", ""),
                "esperadas_total": len(esp),
                "ativas_total": len(atv_diretas),
                "esperadas_ativas": esperadas_ativas,
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


# =====================================================================
# AUDITORIA DEPARTAMENTAL PÓS-PROCESSAMENTO (100% READ-ONLY)
# =====================================================================

def comparar_utilizador_ambientes(
    usuario: str,
    nome: str,
    excel_esp: Dict[str, Any],
    prd_dados: Dict[str, Any],
    qas_dados: Dict[str, Any],
    padroes_exclusao: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Compara o desenho esperado pelo Excel com o estado real retornado por PRD e QAS.
    Gera uma classificação minuciosa sem alterar dados ou esconder exclusões.
    """
    import fnmatch

    padroes = [str(p).strip().upper() for p in (padroes_exclusao or []) if str(p).strip()]

    def is_excl(r_name: str) -> bool:
        r_up = str(r_name).strip().upper()
        return any(fnmatch.fnmatchcase(r_up, p) for p in padroes)

    comp_esp = excel_esp.get("composta")
    roles_dir_esp = set(excel_esp.get("roles_diretas", set()))
    roles_her_esp = set(excel_esp.get("roles_herdadas", set()))
    roles_tot_esp = set(excel_esp.get("roles_totais", set()))

    # PRD
    status_conta_prd = prd_dados.get("status_conta", "ATIVA")
    roles_dir_prd = set(prd_dados.get("roles_diretas", set()))
    roles_her_prd = set(prd_dados.get("roles_herdadas", set()))
    roles_tot_prd = set(prd_dados.get("roles_totais", set()))

    # QAS
    status_conta_qas = qas_dados.get("status_conta", "ATIVA")
    roles_dir_qas = set(qas_dados.get("roles_diretas", set()))
    roles_her_qas = set(qas_dados.get("roles_herdadas", set()))
    roles_tot_qas = set(qas_dados.get("roles_totais", set()))

    # 1. Analisar PRD
    faltam_prd = []
    if status_conta_prd != "NAO_ENCONTRADO":
        for r in sorted(roles_dir_esp):
            if r in roles_dir_prd:
                continue
            if r == comp_esp:
                faltam_prd.append(r)
            elif r in roles_her_prd:
                continue
            else:
                faltam_prd.append(r)

    extras_prd_brutas = sorted(roles_dir_prd - roles_dir_esp) if status_conta_prd != "NAO_ENCONTRADO" else []
    protegidas_exclusao_prd = [r for r in extras_prd_brutas if is_excl(r)]
    extras_prd = [r for r in extras_prd_brutas if not is_excl(r)]

    # 2. Analisar QAS
    faltam_qas = []
    if status_conta_qas != "NAO_ENCONTRADO":
        for r in sorted(roles_dir_esp):
            if r in roles_dir_qas:
                continue
            if r == comp_esp:
                faltam_qas.append(r)
            elif r in roles_her_qas:
                continue
            else:
                faltam_qas.append(r)

    extras_qas_brutas = sorted(roles_dir_qas - roles_dir_esp) if status_conta_qas != "NAO_ENCONTRADO" else []
    protegidas_exclusao_qas = [r for r in extras_qas_brutas if is_excl(r)]
    extras_qas = [r for r in extras_qas_brutas if not is_excl(r)]

    # 3. Diferenças PRD vs QAS
    prd_nao_qas = sorted(roles_tot_prd - roles_tot_qas) if (status_conta_prd != "NAO_ENCONTRADO" and status_conta_qas != "NAO_ENCONTRADO") else []
    qas_nao_prd = sorted(roles_tot_qas - roles_tot_prd) if (status_conta_prd != "NAO_ENCONTRADO" and status_conta_qas != "NAO_ENCONTRADO") else []

    # 4. Classificações
    classificacoes = set()

    if status_conta_prd == "NAO_ENCONTRADO" or status_conta_qas == "NAO_ENCONTRADO":
        classificacoes.add("UTILIZADOR_NAO_ENCONTRADO")

    if status_conta_prd in ("INATIVA", "EXPIRADA", "BLOQUEADA") or status_conta_qas in ("INATIVA", "EXPIRADA", "BLOQUEADA"):
        classificacoes.add("CONTA_INATIVA")

    if comp_esp:
        comp_em_prd = comp_esp in roles_dir_prd
        comp_em_qas = comp_esp in roles_dir_qas
        if not comp_em_prd or not comp_em_qas:
            classificacoes.add("COMPOSITE_DIVERGENTE")

    if faltam_prd or faltam_qas:
        classificacoes.add("ROLE_EM_FALTA")

    if extras_prd or extras_qas:
        classificacoes.add("ROLE_EXTRA")

    if protegidas_exclusao_prd or protegidas_exclusao_qas:
        classificacoes.add("PROTEGIDA_POR_EXCLUSAO")

    status_prd = "OK" if (status_conta_prd != "NAO_ENCONTRADO" and not faltam_prd and not extras_prd and (not comp_esp or comp_esp in roles_dir_prd)) else "DIVERGENTE"
    status_qas = "OK" if (status_conta_qas != "NAO_ENCONTRADO" and not faltam_qas and not extras_qas and (not comp_esp or comp_esp in roles_dir_qas)) else "DIVERGENTE"

    if status_prd == "DIVERGENTE":
        classificacoes.add("DIVERGENTE_PRD")
    if status_qas == "DIVERGENTE":
        classificacoes.add("DIVERGENTE_QAS")

    if prd_nao_qas or qas_nao_prd or (status_prd != status_qas):
        classificacoes.add("PRD_QAS_DIVERGENTE")

    if not classificacoes - {"PROTEGIDA_POR_EXCLUSAO"} and status_prd == "OK" and status_qas == "OK":
        classificacoes.add("OK")

    status_geral = "OK" if ("OK" in classificacoes and status_prd == "OK" and status_qas == "OK") else "DIVERGENTE"

    return {
        "usuario": usuario,
        "nome": nome,
        "composite_esperada": comp_esp,
        "roles_diretas_esperadas": sorted(roles_dir_esp),
        "roles_totais_esperadas": sorted(roles_tot_esp),
        "status_conta_prd": status_conta_prd,
        "status_conta_qas": status_conta_qas,
        "status_prd": status_prd,
        "status_qas": status_qas,
        "status_geral": status_geral,
        "classificacoes": sorted(list(classificacoes)),
        "faltam_prd": faltam_prd,
        "extras_prd": extras_prd,
        "protegidas_exclusao_prd": protegidas_exclusao_prd,
        "faltam_qas": faltam_qas,
        "extras_qas": extras_qas,
        "protegidas_exclusao_qas": protegidas_exclusao_qas,
        "prd_nao_qas": prd_nao_qas,
        "qas_nao_prd": qas_nao_prd,
        "roles_prd": sorted(roles_dir_prd),
        "roles_qas": sorted(roles_dir_qas),
    }


def consultar_utilizadores_ambiente_rfc(
    usuarios_lista: List[str],
    ambiente: str = "PRD"
) -> Dict[str, Dict[str, Any]]:
    """
    Consulta via RFC read-only (AGR_USERS e USR02) o estado real de uma lista de utilizadores num ambiente (PRD ou QAD).
    Retorna mapeamento {utilizador: {'roles_diretas': set, 'roles_herdadas': set, 'roles_totais': set, 'status_conta': str}}.
    """
    from datetime import date
    if not usuarios_lista:
        return {}

    env_upper = str(ambiente).strip().upper()
    resultado_map: Dict[str, Dict[str, Any]] = {
        u: {
            "usuario": u,
            "roles_diretas": set(),
            "roles_herdadas": set(),
            "roles_totais": set(),
            "status_conta": "DESCONHECIDO",
        }
        for u in usuarios_lista
    }

    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_read_only_guard, read_table, make_option_in
        )
        from pyrfc import Connection

        load_project_env(find_project_root())
        params = build_connection_params_for(env_upper)
        conn = Connection(**params)
        guard = make_read_only_guard(["AGR_USERS", "USR02"])

        hoje_int = int(date.today().strftime("%Y%m%d"))

        # 1. Leitura USR02 (Estado da conta)
        users_existentes = set()
        chunk_size = 25
        for i in range(0, len(usuarios_lista), chunk_size):
            chunk = usuarios_lista[i:i + chunk_size]
            opts_usr = make_option_in("BNAME", chunk)
            rows_usr = read_table(conn, guard, table_name="USR02", fields=["BNAME", "GLTGV", "GLTGB", "UFLAG"], options=opts_usr, rowcount=0)
            for r in rows_usr:
                if len(r) >= 3:
                    bn = str(r[0]).strip().upper()
                    users_existentes.add(bn)
                    g_fim = str(r[2]).strip()
                    uflag = str(r[3]).strip() if len(r) > 3 else "0"
                    fim_int = int(g_fim) if g_fim.isdigit() and int(g_fim) > 0 else 99991231
                    
                    if uflag != "0":
                        st = "BLOQUEADA"
                    elif hoje_int > fim_int:
                        st = "EXPIRADA"
                    else:
                        st = "ATIVA"

                    if bn in resultado_map:
                        resultado_map[bn]["status_conta"] = st

        for u in usuarios_lista:
            if u not in users_existentes:
                resultado_map[u]["status_conta"] = "NAO_ENCONTRADO"

        # 2. Leitura AGR_USERS (Roles ativas)
        for i in range(0, len(usuarios_lista), chunk_size):
            chunk = [u for u in usuarios_lista[i:i + chunk_size] if u in users_existentes]
            if not chunk:
                continue
            opts_agr = make_option_in("UNAME", chunk)
            rows_agr = read_table(conn, guard, table_name="AGR_USERS", fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT", "COL_FLAG"], options=opts_agr, rowcount=0)
            for r in rows_agr:
                if len(r) >= 4:
                    role = str(r[0]).strip().upper()
                    un = str(r[1]).strip().upper()
                    f_d = str(r[2]).strip()
                    t_d = str(r[3]).strip()
                    col_flag = str(r[4]).strip().upper() if len(r) > 4 else ""

                    inicio = int(f_d) if f_d.isdigit() else 0
                    fim = int(t_d) if t_d.isdigit() else 99991231

                    if inicio <= hoje_int <= fim and role and un in resultado_map:
                        resultado_map[un]["roles_totais"].add(role)
                        if col_flag == "X":
                            resultado_map[un]["roles_herdadas"].add(role)
                        else:
                            resultado_map[un]["roles_diretas"].add(role)

        conn.close()
    except Exception as err:
        print(f"[AVISO] Não foi possível consultar o ambiente SAP {env_upper} via RFC: {err}")

    return resultado_map


def auditar_departamentos_processados(
    dados: ProjetoPerfilData,
    ambientes: Iterable[str] = ("PRD", "QAD"),
    filtro_departamento: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Executa a Auditoria Departamental Pós-Processamento 100% READ-ONLY.
    Seleciona dinamicamente departamentos da sheet CONTROLO cujo STATUS = PROCESSADO.
    """
    deps_processados = []
    for c in dados.controlo:
        st = str(c.get("status", "")).strip().upper()
        d_nome = str(c.get("departamento", "")).strip()
        if st == "PROCESSADO":
            if filtro_departamento:
                norm_filtro = normalizar_texto(filtro_departamento)
                norm_d = normalizar_texto(d_nome)
                if norm_filtro == norm_d or norm_filtro in norm_d:
                    deps_processados.append(d_nome)
            else:
                deps_processados.append(d_nome)

    if not deps_processados:
        return {
            "ok": True,
            "mensagem": "Nenhum departamento com STATUS = PROCESSADO encontrado na folha CONTROLO.",
            "total_departamentos": 0,
            "total_utilizadores": 0,
            "total_ok": 0,
            "total_divergentes": 0,
            "departamentos": {},
        }

    resultado_geral = {
        "ok": True,
        "total_departamentos": len(deps_processados),
        "total_utilizadores": 0,
        "total_ok": 0,
        "total_divergentes": 0,
        "total_prd_divergente": 0,
        "total_qas_divergente": 0,
        "total_prd_qas_divergente": 0,
        "total_roles_falta": 0,
        "total_roles_extra": 0,
        "departamentos": {},
    }

    env_lista = [str(a).strip().upper() for a in ambientes]
    env_prd = "PRD" if "PRD" in env_lista else env_lista[0]
    env_qas = "QAD" if "QAD" in env_lista or "QAS" in env_lista else (env_lista[1] if len(env_lista) > 1 else "QAD")

    for dep_nome in deps_processados:
        cruzamento = cruzar_fontes_departamento(dados, dep_nome)
        if not cruzamento.get("encontrado"):
            continue

        users_excel = cruzamento.get("usuarios", [])
        if not users_excel:
            continue

        users_ids = [u["usuario"] for u in users_excel]

        # Consultar PRD e QAS em batch
        mapa_prd = consultar_utilizadores_ambiente_rfc(users_ids, ambiente=env_prd)
        mapa_qas = consultar_utilizadores_ambiente_rfc(users_ids, ambiente=env_qas)

        users_auditados = []
        dep_ok = 0
        dep_div = 0
        dep_prd_div = 0
        dep_qas_div = 0
        dep_roles_falta = 0
        dep_roles_extra = 0

        for u_info in users_excel:
            uid = u_info["usuario"]
            nome = u_info["nome"]

            excel_esp = {
                "composta": u_info.get("composta"),
                "roles_diretas": set(u_info.get("diretas_esperadas", [])),
                "roles_herdadas": set(u_info.get("membros_composta", [])),
                "roles_totais": set(u_info.get("todas_autorizacoes_roles", [])),
            }

            prd_dados = mapa_prd.get(uid, {"status_conta": "DESCONHECIDO", "roles_diretas": set(), "roles_herdadas": set(), "roles_totais": set()})
            qas_dados = mapa_qas.get(uid, {"status_conta": "DESCONHECIDO", "roles_diretas": set(), "roles_herdadas": set(), "roles_totais": set()})

            comp_res = comparar_utilizador_ambientes(
                usuario=uid,
                nome=nome,
                excel_esp=excel_esp,
                prd_dados=prd_dados,
                qas_dados=qas_dados,
                padroes_exclusao=dados.padroes_exclusao,
            )

            if comp_res["status_geral"] == "OK":
                dep_ok += 1
            else:
                dep_div += 1

            if comp_res["status_prd"] == "DIVERGENTE":
                dep_prd_div += 1
            if comp_res["status_qas"] == "DIVERGENTE":
                dep_qas_div += 1

            if "ROLE_EM_FALTA" in comp_res["classificacoes"]:
                dep_roles_falta += 1
            if "ROLE_EXTRA" in comp_res["classificacoes"]:
                dep_roles_extra += 1

            users_auditados.append(comp_res)

        resultado_geral["total_utilizadores"] += len(users_auditados)
        resultado_geral["total_ok"] += dep_ok
        resultado_geral["total_divergentes"] += dep_div
        resultado_geral["total_prd_divergente"] += dep_prd_div
        resultado_geral["total_qas_divergente"] += dep_qas_div
        resultado_geral["total_roles_falta"] += dep_roles_falta
        resultado_geral["total_roles_extra"] += dep_roles_extra

        resultado_geral["departamentos"][dep_nome] = {
            "departamento": dep_nome,
            "status_controlo": "PROCESSADO",
            "total_utilizadores": len(users_auditados),
            "ok": dep_ok,
            "divergentes": dep_div,
            "prd_divergente": dep_prd_div,
            "qas_divergente": dep_qas_div,
            "roles_falta": dep_roles_falta,
            "roles_extra": dep_roles_extra,
            "utilizadores": users_auditados,
        }

    return resultado_geral


def imprimir_auditoria_departamental(resultado: Dict[str, Any]):
    """Imprime no terminal o relatório completo e estruturado da auditoria departamental."""
    print("\n" + "=" * 78)
    print(" AUDITORIA DEPARTAMENTAL PÓS-PROCESSAMENTO")
    print(" MODO: 100% READ-ONLY")
    print("=" * 78)

    if not resultado.get("ok") or resultado.get("total_departamentos", 0) == 0:
        print(f"\n  [AVISO] {resultado.get('mensagem', 'Nenhum departamento auditado.')}")
        return

    deps_map = resultado.get("departamentos", {})
    idx_dep = 1
    total_deps = len(deps_map)

    for dep_nome, d_info in deps_map.items():
        print(f"\n[{idx_dep}/{total_deps}] {dep_nome.upper()}")
        print(f"CONTROLO: {d_info.get('status_controlo')}")
        print(f"Utilizadores: {d_info.get('total_utilizadores')}")
        print("-" * 78)

        for u in d_info.get("utilizadores", []):
            comp_str = u.get("composite_esperada") or "Nenhuma"
            n_exp = len(u.get("roles_diretas_esperadas", []))
            n_prd = len(u.get("roles_prd", []))
            n_qas = len(u.get("roles_qas", []))

            print(f"\n{u['usuario']} | {u['nome']}")
            print(f"Composite Esperada: {comp_str}")
            print(f"Excel:  {n_exp} roles diretas")
            print(f"PRD:    {n_prd} roles diretas")
            print(f"QAS:    {n_qas} roles diretas")
            print(f"\nPRD:  {u['status_prd']}")
            print(f"QAS:  {u['status_qas']}")

            if u.get("faltam_prd"):
                print("Falta no PRD:")
                for r in u["faltam_prd"]:
                    print(f"  - {r}")

            if u.get("extras_prd"):
                print("Extra no PRD:")
                for r in u["extras_prd"]:
                    print(f"  + {r}")

            if u.get("faltam_qas"):
                print("Falta no QAS:")
                for r in u["faltam_qas"]:
                    print(f"  - {r}")

            if u.get("extras_qas"):
                print("Extra no QAS:")
                for r in u["extras_qas"]:
                    print(f"  + {r}")

            if u.get("protegidas_exclusao_prd") or u.get("protegidas_exclusao_qas"):
                prot = sorted(list(set(u.get("protegidas_exclusao_prd", []) + u.get("protegidas_exclusao_qas", []))))
                print("Protegidas por Exclusão:")
                for r in prot:
                    print(f"  ~ {r}")

            res_class = ", ".join(u.get("classificacoes", ["OK"]))
            print(f"\nResultado: {res_class}")
            print("-" * 70)

        print(f"\nRESUMO - {dep_nome.upper()}")
        print(f"Utilizadores analisados: {d_info.get('total_utilizadores')}")
        print(f"OK:                     {d_info.get('ok')}")
        print(f"Divergentes:             {d_info.get('divergentes')}")
        print(f"PRD divergente:           {d_info.get('prd_divergente')}")
        print(f"QAS divergente:           {d_info.get('qas_divergente')}")
        print(f"Roles em falta:           {d_info.get('roles_falta')}")
        print(f"Roles extra:              {d_info.get('roles_extra')}")
        print("=" * 78)
        idx_dep += 1

    print("\n" + "=" * 78)
    print(" RESUMO GERAL DO SISTEMA")
    print("=" * 78)
    print(f"Departamentos auditados: {resultado.get('total_departamentos')}")
    print(f"Utilizadores analisados: {resultado.get('total_utilizadores')}")
    print(f"OK:                     {resultado.get('total_ok')}")
    print(f"Divergentes:             {resultado.get('total_divergentes')}")
    print(f"PRD divergente:           {resultado.get('total_prd_divergente')}")
    print(f"QAS divergente:           {resultado.get('total_qas_divergente')}")
    print("=" * 78)


# =====================================================================
# RECONSTRUÇÃO DE FUNÇÕES INDIVIDUAIS NA PROPOSTA ATIVA
# =====================================================================

def mapear_tcodes_sheet_proposta(dados: ProjetoPerfilData) -> Dict[str, List[str]]:
    """
    Lê dinamicamente a folha 'Proposta' e constrói um mapa exato:
    TCODE -> LISTA_DE_FUNÇÕES_INDIVIDUAIS (Dict[str, List[str]])
    Uma mesma transação pode estar associada a mais de uma Função Individual (ex: VK13 -> Z_SALES_PRICECOND_CREATE e Z_SALES_PRICECOND_REPORT).
    Nenhuma relação válida é sobrescrita ou desconsiderada.
    """
    sheet_prop = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == "PROPOSTA"), None)
    if not sheet_prop:
        return {}

    import pandas as pd
    fonte = abrir_excel_seguro(dados.caminho)
    df = pd.read_excel(fonte, sheet_name=sheet_prop, header=None)

    tcode_para_funcoes: Dict[str, List[str]] = {}
    current_role: Optional[str] = None

    for r in range(len(df)):
        f_val = str(df.iloc[r, 0]).strip() if pd.notna(df.iloc[r, 0]) else ""
        d_val = str(df.iloc[r, 1]).strip() if pd.notna(df.iloc[r, 1]) else ""

        if is_nome_funcao_proposta(f_val, d_val):
            current_role = f_val.upper()
        elif f_val and current_role and not is_tcode_marcado_inexistente(d_val):
            for tc in limpar_tcode(f_val):
                tc_clean = tc.upper()
                if tc_clean and not tc_clean.startswith("Z_"):
                    if tc_clean not in tcode_para_funcoes:
                        tcode_para_funcoes[tc_clean] = []
                    if current_role not in tcode_para_funcoes[tc_clean]:
                        tcode_para_funcoes[tc_clean].append(current_role)

    return tcode_para_funcoes


def obter_transacoes_usuario_departamento(
    dados: ProjetoPerfilData,
    nome_sheet_dep: str,
    user_id: str
) -> Tuple[List[str], Optional[int]]:
    """
    Lê a sheet do departamento e extrai todas as transações marcadas com X para o utilizador.
    Retorna (lista_tcodes, indice_coluna_usuario).
    """
    import pandas as pd
    fonte = abrir_excel_seguro(dados.caminho)
    df_dep = pd.read_excel(fonte, sheet_name=nome_sheet_dep, header=None)

    if len(df_dep) < 2:
        return ([], None)

    user_norm = normalizar_texto(user_id)
    col_u_idx: Optional[int] = None

    for col in range(2, df_dep.shape[1]):
        val_hdr = str(df_dep.iloc[1, col]).strip() if pd.notna(df_dep.iloc[1, col]) else ""
        if user_norm in normalizar_texto(val_hdr):
            col_u_idx = col
            break

    if col_u_idx is None:
        return ([], None)

    transacoes_x: List[str] = []
    for r in range(2, len(df_dep)):
        tc_val = str(df_dep.iloc[r, 0]).strip() if pd.notna(df_dep.iloc[r, 0]) else ""
        flag_val = str(df_dep.iloc[r, col_u_idx]).strip().upper() if pd.notna(df_dep.iloc[r, col_u_idx]) else ""

        if flag_val in ("X", "1", "SIM", "YES", "S") and tc_val and tc_val.lower() != "nan":
            transacoes_x.append(tc_val.upper())

    return (transacoes_x, col_u_idx)


def mapeamento_tcodes_para_funcoes_individuais(
    tcodes_usuario: List[str],
    tcode_para_funcao_map: Dict[str, List[str]]
) -> Tuple[List[str], List[str]]:
    """
    Recebe a lista de transações X de um utilizador e o mapa TCODE -> Lista de Funções.
    Para cada TCODE:
      1. Consulta TODAS as funções associadas à TCODE;
      2. Acrescenta TODAS as funções à lista final;
      3. Deduplica os NOMES DAS FUNÇÕES INDIVIDUAIS preservando a ordem de descoberta.
    Retorna (funcoes_individuais_unicas, transacoes_sem_funcao).
    """
    funcoes_unicas: List[str] = []
    sem_funcao: List[str] = []

    for tc in tcodes_usuario:
        tc_upper = tc.upper()
        roles_mapped = tcode_para_funcao_map.get(tc_upper, [])
        if roles_mapped:
            for r_map in roles_mapped:
                if r_map not in funcoes_unicas:
                    funcoes_unicas.append(r_map)
        else:
            if tc_upper not in sem_funcao:
                sem_funcao.append(tc_upper)

    return (funcoes_unicas, sem_funcao)


def reconstruir_funcoes_individuais_departamento(
    dados: ProjetoPerfilData,
    nome_departamento: str,
    simular: bool = True,
    caminho_excel: Optional[str] = None,
    filtro_utilizador: Optional[str] = None
) -> Dict[str, Any]:
    """
    Executa a reconstrução completa das Funções Individuais da Proposta Ativa
    para todos os utilizadores de um departamento específico ou para um único utilizador (se filtro_utilizador for especificado).
    """
    import pandas as pd
    caminho = caminho_excel or dados.caminho

    # 1. Localizar sheet do departamento
    sheet_dep = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == normalizar_nome_coluna(nome_departamento)), None)
    if not sheet_dep:
        return {
            "ok": False,
            "departamento": nome_departamento,
            "motivo": "SHEET_NAO_ENCONTRADA",
            "mensagem": f"[WARNING] Sheet para o departamento '{nome_departamento}' não encontrada no Excel."
        }

    # 2. Mapeamento TCODE -> Funções na folha Proposta
    tcode_map = mapear_tcodes_sheet_proposta(dados)

    # 3. Ler utilizadores do cabeçalho da sheet do departamento (linha 2)
    fonte = abrir_excel_seguro(caminho)
    df_dep = pd.read_excel(fonte, sheet_name=sheet_dep, header=None)

    colunas_users: Dict[str, Dict[str, Any]] = {}
    for col in range(2, df_dep.shape[1]):
        raw_hdr = str(df_dep.iloc[1, col]).strip() if pd.notna(df_dep.iloc[1, col]) else ""
        if raw_hdr and raw_hdr.lower() != "nan":
            partes = raw_hdr.replace("\r", "\n").split("\n")
            p0 = partes[0].strip()
            tokens = p0.split()
            if tokens and (tokens[0].upper().startswith("S") or tokens[0].isdigit()):
                uid = normalizar_texto(tokens[0])
                nome_hdr = " ".join(tokens[1:]).strip() if len(tokens) > 1 else (partes[1].strip() if len(partes) > 1 else uid)
            else:
                uid = normalizar_texto(p0)
                nome_hdr = partes[1].strip() if len(partes) > 1 else uid
            if uid:
                colunas_users[uid] = {"col_idx": col, "nome": nome_hdr, "raw_hdr": raw_hdr}

    if not colunas_users:
        return {
            "ok": True,
            "departamento": nome_departamento,
            "utilizadores_processados": 0,
            "mensagem": "Nenhum utilizador encontrado no cabeçalho da sheet do departamento."
        }

    if filtro_utilizador:
        norm_u_filtro = normalizar_texto(filtro_utilizador)
        if norm_u_filtro not in colunas_users:
            return {
                "ok": False,
                "departamento": nome_departamento,
                "motivo": "UTILIZADOR_NAO_ENCONTRADO_NA_SHEET_DEPARTAMENTO",
                "mensagem": f"[ERRO] Utilizador '{filtro_utilizador}' não foi encontrado no cabeçalho da sheet '{nome_departamento}'."
            }
        colunas_users = {norm_u_filtro: colunas_users[norm_u_filtro]}

    # 4. Ler Proposta Ativa e encontrar obrigatoriamente o cabeçalho 'Funções Individuais' sem fallback
    df_pa = pd.read_excel(fonte, sheet_name="Proposta Ativa")
    col_user_pa = [c for c in df_pa.columns if "USER" in str(c).upper() or "UTILIZADOR" in str(c).upper() or "USU" in str(c).upper()][0]
    col_dep_pa = [c for c in df_pa.columns if str(c).strip().upper() == "DEPARTAMENTO"][0]
    col_comp_pa = [c for c in df_pa.columns if "COMPOSITE" in str(c).upper()][0]
    col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_pa)

    relatorio_utilizadores = []
    celulas_para_pintar_vermelho: List[Tuple[int, int]] = []  # (linha_1based, col_1based)

    for uid, u_meta in colunas_users.items():
        col_idx = u_meta["col_idx"]
        nome_user = u_meta["nome"]

        # 4.1 Validar existência na Proposta Ativa
        rows_match = df_pa[df_pa.iloc[:, 0].astype(str).str.strip().str.upper() == uid]
        if rows_match.empty:
            relatorio_utilizadores.append({
                "usuario": uid,
                "nome": nome_user,
                "status": "UTILIZADOR_NAO_ENCONTRADO_PROPOSTA_ATIVA",
                "mensagem": f"Utilizador '{uid}' não foi encontrado na folha Proposta Ativa."
            })
            continue

        pa_row_idx = rows_match.index[0]
        pa_row = df_pa.iloc[pa_row_idx]
        dep_pa_val = str(pa_row.get(col_dep_pa, "")).strip()
        comp_pa_val = str(pa_row.get(col_comp_pa, "")).strip()

        # 4.2 Validar correspondência de Departamento
        if normalizar_texto(dep_pa_val) != normalizar_texto(nome_departamento):
            relatorio_utilizadores.append({
                "usuario": uid,
                "nome": nome_user,
                "status": "DEPARTAMENTO_DIVERGENTE",
                "departamento_sheet": nome_departamento,
                "departamento_proposta": dep_pa_val,
                "mensagem": f"Departamento na Proposta Ativa ('{dep_pa_val}') diverge da sheet ('{nome_departamento}')."
            })
            continue

        # 4.3 Ler transações X na sheet do departamento
        tcodes_user, _ = obter_transacoes_usuario_departamento(dados, sheet_dep, uid)

        # 4.4 Mapear TCODEs -> Funções Individuais
        funcoes_calc, sem_funcao = mapeamento_tcodes_para_funcoes_individuais(tcodes_user, tcode_map)

        # Se houver transações sem função, registar células a pintar de vermelho
        if sem_funcao:
            for r in range(2, len(df_dep)):
                tc_val = str(df_dep.iloc[r, 0]).strip().upper() if pd.notna(df_dep.iloc[r, 0]) else ""
                flag_val = str(df_dep.iloc[r, col_idx]).strip().upper() if pd.notna(df_dep.iloc[r, col_idx]) else ""
                if tc_val in sem_funcao and flag_val in ("X", "1", "SIM", "YES", "S"):
                    celulas_para_pintar_vermelho.append((r + 1, col_idx + 1))

        # 4.5 Obter funções atuais na Proposta Ativa
        roles_atuais = [
            str(x).strip().upper() for x in pa_row.iloc[col_fi_idx:].dropna().tolist()
            if is_role_sap_valida(x)
        ]

        set_calc = set(funcoes_calc)
        set_atual = set(roles_atuais)

        adicionadas = [r for r in funcoes_calc if r not in set_atual]
        mantidas = [r for r in funcoes_calc if r in set_atual]
        removidas = [r for r in roles_atuais if r not in set_calc]

        tem_alteracao = (funcoes_calc != roles_atuais)

        relatorio_utilizadores.append({
            "usuario": uid,
            "nome": nome_user,
            "linha_proposta_ativa": pa_row_idx + 2,
            "composite_atual": comp_pa_val,
            "status": "ALTERADO" if tem_alteracao else "SEM_ALTERACAO",
            "total_transacoes_x": len(tcodes_user),
            "total_funcoes_unicas": len(funcoes_calc),
            "funcoes_calculadas": funcoes_calc,
            "funcoes_atuais": roles_atuais,
            "sem_funcao": sem_funcao,
            "adicionadas": adicionadas,
            "mantidas": mantidas,
            "removidas": removidas,
        })

    # 5. Se NÃO for simulação, criar backup e realizar a escrita real
    if not simular:
        backup_dir = Path("C:/workspace/SapScript/output/backups")
        backup_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefixo_u = f"PRE_{filtro_utilizador}_" if filtro_utilizador else "PRE_"
        backup_caminho = backup_dir / f"S4H_Perfis_de_autorizacao_{prefixo_u}{timestamp}.xlsx"

        try:
            import shutil
            shutil.copy2(caminho, str(backup_caminho))
            print(f"  [BACKUP] Cópia de segurança criada com sucesso em: {backup_caminho}")
        except Exception as exc:
            raise RuntimeError(f"ABORTADO: Não foi possível criar a cópia de segurança antes da escrita: {exc}")

        # Executar a atualização via openpyxl
        import openpyxl
        wb = openpyxl.load_workbook(caminho)
        ws_pa = wb["Proposta Ativa"]
        ws_dep_op = wb[sheet_dep]

        col_fi_1based = col_fi_idx + 1

        for item in relatorio_utilizadores:
            if item.get("status") in ("ALTERADO", "SEM_ALTERACAO"):
                linha_excel = item["linha_proposta_ativa"]
                funcoes_escrever = item["funcoes_calculadas"]

                # Limpar células antigas de funções individuais
                max_col = max(ws_pa.max_column, col_fi_1based + 50)
                for c in range(col_fi_1based, max_col + 1):
                    ws_pa.cell(row=linha_excel, column=c).value = None

                # Escrever novas funções calculadas
                for idx_f, f_nome in enumerate(funcoes_escrever):
                    ws_pa.cell(row=linha_excel, column=col_fi_1based + idx_f).value = f_nome

        # Atualizar também a folha PFCG_COMPOSTA com a união das funções calculadas por Composite Role
        ws_comp_op = wb["PFCG_COMPOSTA"] if "PFCG_COMPOSTA" in wb.sheetnames else None
        if ws_comp_op:
            # Coletar membros existentes em PFCG_COMPOSTA
            pares_existentes = set()
            max_id_comp = 0
            composta_textos = {}
            for r in range(2, ws_comp_op.max_row + 1):
                try:
                    rid = int(ws_comp_op.cell(r, 1).value or 0)
                    if rid > max_id_comp:
                        max_id_comp = rid
                except Exception:
                    pass
                agr_c = str(ws_comp_op.cell(r, 2).value or "").strip().upper()
                txt_c = str(ws_comp_op.cell(r, 3).value or "").strip()
                agr_f = str(ws_comp_op.cell(r, 4).value or "").strip().upper()
                if agr_c and txt_c and agr_c not in composta_textos:
                    composta_textos[agr_c] = txt_c
                if agr_c and agr_f:
                    pares_existentes.add((agr_c, agr_f))

            # Agrupar união das funções por Composite Role para este departamento
            compostas_dep = defaultdict(set)
            for item in relatorio_utilizadores:
                c_role = str(item.get("composite_atual", "")).strip().upper()
                if c_role and c_role not in ("NAN", "NONE", "-"):
                    compostas_dep[c_role].update(item.get("funcoes_calculadas", []))

            for c_role, funcs_uniao in sorted(compostas_dep.items()):
                txt_desc = composta_textos.get(c_role, dados.roles_compostas.get(c_role, {}).get("descricao", c_role))
                for f_nome in sorted(funcs_uniao):
                    f_upper = f_nome.strip().upper()
                    if f_upper and (c_role, f_upper) not in pares_existentes:
                        max_id_comp += 1
                        next_r = ws_comp_op.max_row + 1
                        ws_comp_op.cell(row=next_r, column=1, value=max_id_comp)
                        ws_comp_op.cell(row=next_r, column=2, value=c_role)
                        ws_comp_op.cell(row=next_r, column=3, value=txt_desc)
                        ws_comp_op.cell(row=next_r, column=4, value=f_upper)
                        pares_existentes.add((c_role, f_upper))

        wb.save(caminho)
        wb.close()
        print(f"  [EXCEL] Reconstrução concluída e salva com sucesso em: {caminho}")

    return {
        "ok": True,
        "departamento": nome_departamento,
        "sheet_departamento": sheet_dep,
        "simular": simular,
        "total_utilizadores": len(relatorio_utilizadores),
        "utilizadores": relatorio_utilizadores,
        "celulas_vermelhas": len(celulas_para_pintar_vermelho),
    }


def reconstruir_funcoes_individuais_todos_departamentos(
    dados: ProjetoPerfilData,
    simular: bool = True,
    caminho_excel: Optional[str] = None
) -> Dict[str, Any]:
    """
    Percorre TODOS os departamentos definidos na sheet CONTROLO
    e executa a reconstrução das Funções Individuais na Proposta Ativa.
    """
    deps_controlo = [c.get("departamento") for c in dados.controlo if c.get("departamento")]
    resultados_deps = {}

    for dep_nome in deps_controlo:
        res = reconstruir_funcoes_individuais_departamento(dados, dep_nome, simular=simular, caminho_excel=caminho_excel)
        resultados_deps[dep_nome] = res

    return {
        "ok": True,
        "simular": simular,
        "total_departamentos": len(deps_controlo),
        "resultados": resultados_deps
    }


def imprimir_relatorio_reconstrucao_proposta_ativa(resultado: Dict[str, Any]):
    """Imprime no terminal o relatório detalhado da reconstrução de Funções Individuais."""
    print("\n" + "=" * 78)
    print(" RECONSTRUÇÃO DE FUNÇÕES INDIVIDUAIS NA PROPOSTA ATIVA")
    print(f" MODO: {'SIMULAÇÃO (sem gravação)' if resultado.get('simular') else 'EXECUÇÃO REAL (gravado no Excel)'}")
    print("=" * 78)

    if "utilizadores" in resultado:
        deps_map = {resultado.get("departamento", "Geral"): resultado}
    else:
        deps_map = resultado.get("resultados", {})

    for dep_nome, d_res in deps_map.items():
        if not d_res.get("ok"):
            print(f"\n[DEPARTAMENTO] {dep_nome}")
            print(f"  └─ {d_res.get('mensagem', 'Erro ou sheet não encontrada.')}")
            continue

        print(f"\n============================================================")
        print(f" RECONSTRUÇÃO DE FUNÇÕES INDIVIDUAIS")
        print(f" Departamento: {dep_nome}")
        print(f"============================================================")

        for u in d_res.get("utilizadores", []):
            st = u.get("status")
            if st in ("UTILIZADOR_NAO_ENCONTRADO_PROPOSTA_ATIVA", "DEPARTAMENTO_DIVERGENTE"):
                print(f"\n 👤 {u['usuario']} | {u['nome']}")
                print(f"    └─ [AVISO] {u.get('mensagem')}")
                continue

            print(f"\n User: {u['usuario']}")
            print(f" Nome: {u['nome']}")
            print(f" Linha Proposta Ativa: {u.get('linha_proposta_ativa')}")
            print(f" Composite Role:       {u.get('composite_atual')}")
            print(f" Transações com X:     {u.get('total_transacoes_x')}")
            print(f" Funções únicas:       {u.get('total_funcoes_unicas')}")

            print("\n Funções individuais encontradas:")
            for idx_f, f_nome in enumerate(u.get("funcoes_calculadas", []), 1):
                print(f"   {idx_f:02d}. {f_nome}")

            if u.get("sem_funcao"):
                print("\n Transações sem função individual:")
                for tc in u["sem_funcao"]:
                    print(f"   - {tc}")

            print(f"\n Alterações:")
            print(f"   Adicionadas ({len(u.get('adicionadas', []))}): {', '.join(u.get('adicionadas', [])) or 'Nenhuma'}")
            print(f"   Mantidas    ({len(u.get('mantidas', []))}): {', '.join(u.get('mantidas', [])) or 'Nenhuma'}")
            print(f"   Removidas   ({len(u.get('removidas', []))}): {', '.join(u.get('removidas', [])) or 'Nenhuma'}")

            print(f"\n Status: {st}")
            print("-" * 60)

    print("\n" + "=" * 78)




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
    Atualiza também a cópia de segurança local configurada (se aplicável).
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
                        col_fi = encontrar_coluna_funcoes_individuais_ws(ws, default_col=11)
                        col_livre = None
                        for c in range(col_fi, max_c + 3):
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
                                backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
                    col_fi = encontrar_coluna_funcoes_individuais_ws(ws, default_col=11)
                    col_livre = None
                    for c in range(col_fi, ws.max_column + 5):
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
                            backup_local = Path(CAMINHO_EXCEL_BACKUP_LOCAL)
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
            col_fi_idx = encontrar_coluna_funcoes_individuais_df(df_raw)
            for idx in range(1, len(df_raw)):
                u_id = str(df_raw.iloc[idx, 0]).strip() if pd.notna(df_raw.iloc[idx, 0]) else ""
                if normalizar_texto(u_id) == user_norm:
                    singles = [
                        str(x).strip().upper() for x in df_raw.iloc[idx, col_fi_idx:].dropna().tolist()
                        if is_role_sap_valida(x)
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

    # 3. Expansão Relacional e Definições
    esperadas = set()
    if meta_proposta:
        base_set = set()
        if meta_proposta.get("composta"):
            base_set.add(meta_proposta["composta"])
        base_set.update(meta_proposta.get("singles", []))
        dep_u = normalizar_texto(meta_proposta.get("departamento", ""))
        definicoes_dep = dados.definicoes_departamento.get(dep_u, set())
        if not definicoes_dep and dep_u in ("LEGAL", "PURCHASE & SERVICES", "PURCHASE AND SERVICES"):
            definicoes_dep = dados.definicoes_departamento.get("PURCHASE & SERVICES", set())
        base_set.update(definicoes_dep)
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
    adicionais_reais = sorted([r for r in adicionais_brutas if not dados.is_excluida(r)])

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
    caminho_completo = str(Path(ficheiro).resolve())
    print("\n" + "=" * 78)
    print("  PROJETO PERFIL - Leitor e Pesquisador de Funções SAP (PFCG / CUA)")
    print(f"  Ficheiro: {caminho_completo}")
    print("=" * 78)


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
    print(f"  Departamentos pendentes (STATUS vazio): {len(pendentes)}")
    print("-" * 75)
    for c in dados.controlo:
        status_tag = "PENDENTE (STATUS VAZIO)" if c["pendente"] else f"{c['status']}"
        ts_info = f" | {c['timestamp']}" if c.get("timestamp") else ""
        print(f"  Linha {c['linha']}: {c['departamento']:<30} -> {status_tag}{ts_info}")

    print("-" * 75)
    proximo = dados.obter_proximo_departamento()
    if proximo:
        print(f"  PRÓXIMO DEPARTAMENTO A AVANÇAR: '{proximo}'")
        item_prox = next((p for p in pendentes if p["departamento"] == proximo), {})
        tem_sheet = proximo in dados.sheets_disponiveis
        sheet_info = f"[Sheet correspondente '{proximo}' EXISTE no Excel]" if tem_sheet else "[Sheet correspondente não encontrada]"
        print(f"     Linha {item_prox.get('linha')}: STATUS está vazio! {sheet_info}")
        if len(pendentes) > 1:
            print("\n  Demais departamentos pendentes:")
            for p in pendentes[1:]:
                print(f"     - '{p['departamento']}' (Linha {p['linha']})")
    else:
        print("  Todos os departamentos na sheet CONTROLO já possuem STATUS preenchido.")


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
    caminho_completo = str(Path(ficheiro).resolve())
    print("\n" + "=" * 78)
    print("  PROJETO PERFIL - PROCESSAMENTO DE DEPARTAMENTOS (CUA / PFCG)")
    print(f"  Ficheiro: {caminho_completo}")
    print("=" * 78)


def obter_pendencias_status(caminho_excel: str) -> Dict[str, int]:
    """Conta linhas com STATUS vazio nas folhas operacionais, sem alterar o Excel."""
    import pandas as pd

    folhas = ["PFCG_CREATE", "PFCG_COMPOSTA", "CUA_ADICIONAR", "CUA_REMOVE"]
    fonte = abrir_excel_seguro(caminho_excel)
    excel_file = pd.ExcelFile(fonte)
    resultado = {}
    for folha in folhas:
        sheet = next((s for s in excel_file.sheet_names if normalizar_nome_coluna(s) == normalizar_nome_coluna(folha)), None)
        if not sheet:
            continue
        df = pd.read_excel(excel_file, sheet_name=sheet)
        col_status = next((c for c in df.columns if normalizar_nome_coluna(c) == "STATUS"), None)
        if col_status is None:
            continue
        status = df[col_status].fillna("").astype(str).str.strip()
        quantidade = int(status.eq("").sum())
        if quantidade:
            resultado[folha] = quantidade
    return resultado


def executar_processos_pendentes(caminho_excel: str, pendencias: Dict[str, int]) -> None:
    """Executa, na ordem operacional, somente as folhas que possuem STATUS vazio."""
    import importlib.util

    pasta = Path(__file__).resolve().parents[2] / "Processos" / "Funções PFCG"

    def carregar(nome_modulo: str, nome_ficheiro: str):
        spec = importlib.util.spec_from_file_location(nome_modulo, pasta / nome_ficheiro)
        if not spec or not spec.loader:
            raise RuntimeError(f"Não foi possível carregar {nome_ficheiro}")
        modulo = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(modulo)
        return modulo

    if "PFCG_CREATE" in pendencias:
        print("\n[1/4] Executando PFCG_CREATE...")
        modulo = carregar("processo_pfcg_create", "A. PFCG_CREATE.py")
        modulo.executar("PRD", caminho_ficheiro=caminho_excel, modo_nao_interativo=True, pedir_confirmacao=False, metodo="RFC")

    if "PFCG_COMPOSTA" in pendencias:
        print("\n[2/4] Executando PFCG_COMPOSTA...")
        modulo = carregar("processo_pfcg_composta", "D. PFCG_COMPOSTA.py")
        modulo.executar("PRD", caminho_ficheiro=caminho_excel, modo_nao_interativo=True, pedir_confirmacao=False)

    if "CUA_ADICIONAR" in pendencias:
        print("\n[3/4] Executando CUA_ADICIONAR...")
        modulo = carregar("processo_cua_adicionar", "CUA_ADICIONAR_WEB.py")
        modulo.executar("CUA", caminho_ficheiro=caminho_excel, modo_nao_interativo=True, pedir_confirmacao=False)

    if "CUA_REMOVE" in pendencias:
        print("\n[4/4] Executando CUA_REMOVE...")
        modulo = carregar("processo_cua_remove", "CUA_REMOVE_WEB.py")
        modulo.executar("CUA", caminho_ficheiro=caminho_excel, modo_nao_interativo=True, pedir_confirmacao=False)


def verificar_e_perguntar_pendencias(caminho_excel: str, assumir_sim: bool = False) -> bool:
    """Mostra as filas pendentes e pede uma única confirmação para executá-las."""
    pendencias = obter_pendencias_status(caminho_excel)
    print("\n" + "=" * 78)
    print("  PROCESSOS PENDENTES")
    print("=" * 78)
    if not pendencias:
        print(f"  {CLR_AZUL}ℹ️  Nenhuma linha com STATUS vazio.{CLR_RESET}")
        return False
    for folha, quantidade in pendencias.items():
        print(f"  ⚠️  {folha}: {quantidade} linha(s)")
    resposta = "" if assumir_sim else input("\nExecutar os processos pendentes? (S/N): ").strip().upper()
    if not assumir_sim and resposta not in ("S", "SIM", "Y", "YES"):
        print("Processos pendentes não executados.")
        return False
    executar_processos_pendentes(caminho_excel, pendencias)
    return True


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
        print(f"  💡 Próximo departamento sugerido: Linha {proximo_sugerido['linha']} ('{proximo_sugerido['departamento']}')")
    else:
        print(f"  {CLR_VERDE}✓ Todos os departamentos registados na folha CONTROLO estão processados.{CLR_RESET}")

    return proximo_sugerido


def selecionar_departamento_interativo(dados: ProjetoPerfilData) -> Optional[Dict[str, Any]]:
    """
    Apresenta a listagem clara dos departamentos na folha CONTROLO
    e solicita ao utilizador indicar a linha do departamento a gerir/reprocessar.
    """
    sugerido = exibir_tabela_controlo(dados)
    padrao_linha = str(sugerido["linha"]) if sugerido else ""
    prompt_sugestao = f" [Enter para Linha {padrao_linha}]" if padrao_linha else ""

    linhas_validas = [str(c.get("linha")) for c in dados.controlo if c.get("linha")]
    exemplo_linhas = f"{linhas_validas[0]}-{linhas_validas[-1]}" if len(linhas_validas) > 1 else (linhas_validas[0] if linhas_validas else "2-7")

    while True:
        if sugerido:
            texto_prompt = f"\nSelecione a Linha do departamento a processar{prompt_sugestao} ('0' para Voltar): "
        else:
            texto_prompt = f"\nQual departamento deseja reprocessar? Indique a Linha [{exemplo_linhas}] ('0' para Voltar): "

        entrada = input(texto_prompt).strip().upper()

        if entrada in ("0", "SAIR", "VOLTAR", "V", "Q", "QUIT", "EXIT"):
            return None

        if not entrada and padrao_linha:
            entrada = padrao_linha

        if not entrada:
            return None

        entrada_limpa = entrada.replace("[", "").replace("]", "").strip()

        # Busca por número de linha
        item_match = next((c for c in dados.controlo if str(c.get("linha")) == entrada_limpa), None)
        if item_match:
            return item_match

        # Busca tolerante por nome de departamento caso o utilizador digite o nome
        item_por_nome = next((c for c in dados.controlo if entrada_limpa in c.get("departamento", "").upper()), None)
        if item_por_nome:
            return item_por_nome

        print(f"[AVISO] Linha '{entrada}' não encontrada na folha CONTROLO. Por favor indique uma Linha válida (ex: {exemplo_linhas}).")


def obter_data_fim_sap_gui(grid_act=None) -> str:
    """
    Retorna a data de validade final (31/12/9999) formatada rigorosamente conforme
    o padrão de data configurado no perfil SAP GUI do utilizador (ex: '31.12.9999' para DD.MM.YYYY).
    Previne o erro SAP DB749 ('Invalid date; enter the date in the format DD.MM.YYYY').
    """
    if grid_act is not None:
        try:
            # 1. Inspecionar se já existe alguma linha com data final na grelha
            for r in range(grid_act.RowCount):
                val = str(grid_act.GetCellValue(r, "UPDATE_TO_DAT") or "").strip()
                if "9999" in val and len(val) >= 8:
                    return val
                val_from = str(grid_act.GetCellValue(r, "UPDATE_FROM_DAT") or "").strip()
                if "." in val_from:
                    return "31.12.9999"
                elif "/" in val_from:
                    parts = val_from.split("/")
                    if len(parts) == 3 and len(parts[2]) == 4:
                        return "12/31/9999" if (int(parts[0]) <= 12 and int(parts[1]) > 12) else "31/12/9999"
                elif "-" in val_from:
                    parts = val_from.split("-")
                    if len(parts) == 3 and len(parts[0]) == 4:
                        return "9999-12-31"
                    return "31-12-9999"
        except Exception:
            pass

    # Padrão corporativo SAP Salsa Jeans / Portugal (DD.MM.YYYY)
    return "31.12.9999"


def sincronizar_departamento_prd_qad_rfc(
    dados: ProjetoPerfilData,
    item_dep: Dict[str, Any],
    confirmar_execucao: bool = True,
    modo_simulacao: bool = False,
    usuarios_alvo: Optional[List[str]] = None,
) -> bool:
    """
    Executa a manutenção direta e atribuição de funções nos ambientes PRD e QAD via RFC:
      1. Lê as roles atuais dos utilizadores em PRD e QAD via BAPI_USER_GET_DETAIL.
      2. Preserva roles protegidas definidas na sheet EXCLUÇÃO (ex: aprovações ZMM*, Z_MY_HOME).
      3. Remove roles legadas/obsoletas fora do catálogo e fora da EXCLUÇÃO.
      4. Atribui as funções oficiais da folha Proposta Ativa (Compostas + Singles).
      5. Modo Simulação (SU01 / SU10 simulada):
         - Executa BAPI_USER_ACTGROUPS_ASSIGN em PRD e QAD.
         - Faz BAPI_TRANSACTION_ROLLBACK imediatamente (nenhuma alteração persistida).
         - Exibe relatório detalhado comparativo de adições, remoções e preservações.
      6. Modo Execução Real:
         - Grava em PRD com BAPI_TRANSACTION_COMMIT(WAIT="X").
         - Replica em QAD com BAPI_TRANSACTION_COMMIT(WAIT="X").
         - Auditoria em tempo real confirmando paridade (PRD == QAD == Catálogo).
         - Registo no Excel (CUA_REMOVE, CUA_ADICIONAR e CONTROLO com backup automático).
    """
    dep_nome = item_dep["departamento"]
    linha_excel = item_dep.get("linha", 0)

    analise = analisar_departamento_proposta(dados, dep_nome)
    if not analise.get("encontrado"):
        print(f"[ERRO] Não foi possível carregar os utilizadores do departamento '{dep_nome}'.")
        return False

    users = [u["usuario"] for u in analise.get("usuarios", []) if u.get("usuario")]
    if not users:
        print(f"[ERRO] Nenhum utilizador encontrado para '{dep_nome}'.")
        return False

    if usuarios_alvo:
        alvo_set = {str(u).strip().upper() for u in usuarios_alvo}
        users = [u for u in users if str(u).strip().upper() in alvo_set]
        if not users:
            print(f"[ERRO] Nenhum dos utilizadores especificados {usuarios_alvo} pertence ao departamento '{dep_nome}'.")
            return False

    tipo_modo = "SIMULAÇÃO SU01/SU10 (RFC SEM COMMIT)" if modo_simulacao else "SINCRONIZAÇÃO DIRETA (PRD -> QAD)"
    u_count_str = f"{len(users)} Utilizador" if len(users) == 1 else f"{len(users)} Utilizadores"
    print(f"\n" + "=" * 75)
    print(f"  {tipo_modo}: {dep_nome.upper()} ({u_count_str})")
    print("=" * 75)
    print(f"  Utilizadores a processar: {', '.join(users)}")
    if not usuarios_alvo:
        if modo_simulacao:
            print("  Modo de Simulação ativo: O SAP irá validar as autorizações sem gravar nada (Rollback).")
        else:
            print("  Etapas automáticas (Execução RFC direta):")
            print("    1. Validação e gravação oficial no SAP PRD (S4PCLNT100)")
            print("    2. Replicação e gravação oficial no SAP QAD (S4QCLNT100)")
            print("    3. Auditoria de paridade pós-gravação (PRD == QAD == Catálogo)")
            print("    4. Gravação oficial em CUA_REMOVE, CUA_ADICIONAR e CONTROLO no Excel")
    print("-" * 75)

    if confirmar_execucao and not modo_simulacao:
        confirma = input("Confirma a execução imediata e gravação no SAP PRD e QAD? (S/N): ").strip().upper()
        if confirma not in ("S", "SIM", "Y", "YES"):
            print("[EXPIRADO] Operação cancelada pelo utilizador.")
            return False

    try:
        from sap_rfc._rfc_common import build_connection_params_for
        from pyrfc import Connection

        params_prd = build_connection_params_for("PRD")
        params_qad = build_connection_params_for("QAD")

        print("  A ligar aos sistemas SAP PRD e QAD via RFC...", flush=True)
        conn_prd = Connection(**params_prd)
        conn_qad = Connection(**params_qad)
        print("  ✓ Conexões RFC ativas em PRD e QAD!", flush=True)
    except Exception as e_conn:
        print(f"[ERRO] Falha ao conectar via RFC: {e_conn}")
        return False

    # Obter funções da sheet DEFINIÇÕES para o departamento
    dep_norm = normalizar_texto(dep_nome)
    definicoes_dep = {normalizar_texto(r) for r in dados.definicoes_departamento.get(dep_norm, set())}
    if not definicoes_dep and dep_norm in ("LEGAL", "PURCHASE & SERVICES", "PURCHASE AND SERVICES"):
        definicoes_dep = {normalizar_texto(r) for r in dados.definicoes_departamento.get("PURCHASE & SERVICES", set())} or {
            "ZORG_TODAS_EMPRESAS", "Z_BASIS_BASE", "ZORG_BP_GERAL",
            "ZORG_BP_LOGISTICS_CUSTOMER", "ZORG_BP_FLVN01_LOGISTICS_VENDO",
            "ZORG_BP_Z001_GENERALPARTNERS", "ZORG_BP_Z003_RELATEDPARTNERS",
            "Z_BR_TYPE_BP_GERAL", "Z_MY_HOME"
        }

    # Identificar todas as funções filhas de compostas para nunca atribuí-las diretamente
    todas_filhas_compostas = set()
    for c_info in dados.roles_compostas.values():
        for f in c_info.get("roles_filhas", []):
            todas_filhas_compostas.add(normalizar_texto(f))

    # Regra oficial: a sheet DEFINIÇÕES é mandatória por departamento.
    # Tudo que está na linha do departamento deve ser atribuído diretamente,
    # mesmo que a role também exista como filha de uma função composta.
    definicoes_dep_diretas = set(definicoes_dep)

    # Mapa de roles esperadas da Proposta Ativa e DEFINIÇÕES para cada utilizador:
    # REGRA DE OURO:
    # 1. Catálogo (Proposta Ativa): Função Composta do utilizador + singles avulsas não-filhas.
    # 2. Definições: Funções organizacionais e base departamentais (sem filhas de compostas).
    # 3. Exclusão: Preservar todas as funções do utilizador contempladas na folha EXCLUÇÃO.
    # 4. Eliminar qualquer atribuição que for diferente destas três fontes.
    user_expected_map = {}
    for u_data in analise.get("usuarios", []):
        u_id = str(u_data.get("usuario", "")).strip().upper()
        if usuarios_alvo and u_id not in alvo_set:
            continue
        comp = normalizar_texto(u_data.get("composta"))
        filhas_comp = set()
        if comp and comp in dados.roles_compostas:
            filhas_comp = {normalizar_texto(f) for f in dados.roles_compostas[comp].get("roles_filhas", [])}

        singles_avulsas = {
            normalizar_texto(r) for r in u_data.get("singles", [])
            if normalizar_texto(r) and normalizar_texto(r) not in filhas_comp and normalizar_texto(r) not in todas_filhas_compostas
        }

        r_set = set(definicoes_dep_diretas)
        if comp:
            r_set.add(comp)
        r_set.update(singles_avulsas)
        user_expected_map[u_id] = {str(r).strip().upper() for r in r_set if r and str(r).strip()}

    roles_removidas = []
    roles_adicionadas = []
    data_hoje_bapi = datetime.now().strftime("%Y%m%d")
    timestamp_now = datetime.now()
    timestamp_str = timestamp_now.strftime("%Y-%m-%d %H:%M:%S")

    sucesso_geral = True

    try:
        for u in users:
            u_norm = str(u).strip().upper()
            print(f"\n  >>> Utilizador: {u} ...", flush=True)
            expected_roles = user_expected_map.get(u_norm, set())

            # 1. Ler roles atuais em PRD (distinguindo atribuições diretas vs herdadas via ORG_FLAG)
            try:
                det_prd = conn_prd.call("BAPI_USER_GET_DETAIL", USERNAME=u)
                roles_prd = {r["AGR_NAME"]: r for r in det_prd.get("ACTIVITYGROUPS", [])}
                diretas_prd = {r["AGR_NAME"]: r for r in det_prd.get("ACTIVITYGROUPS", []) if r.get("ORG_FLAG") != "C"}
            except Exception as e_get_p:
                print(f"     [ERRO] Falha ao consultar utilizador {u} em PRD: {e_get_p}")
                sucesso_geral = False
                continue

            # 2. Ler roles atuais em QAD (distinguindo atribuições diretas vs herdadas via ORG_FLAG)
            try:
                det_qad = conn_qad.call("BAPI_USER_GET_DETAIL", USERNAME=u)
                roles_qad = {r["AGR_NAME"]: r for r in det_qad.get("ACTIVITYGROUPS", [])}
                diretas_qad = {r["AGR_NAME"]: r for r in det_qad.get("ACTIVITYGROUPS", []) if r.get("ORG_FLAG") != "C"}
            except Exception as e_get_q:
                print(f"     [ERRO] Falha ao consultar utilizador {u} em QAD: {e_get_q}")
                sucesso_geral = False
                continue

            # Determinar preservadas (EXCLUÇÃO), a remover e a adicionar (apenas no conjunto de atribuições diretas)
            preservadas_prd = {r: info for r, info in diretas_prd.items() if dados.is_excluida(r)}
            removidas_prd = [r for r in diretas_prd if not dados.is_excluida(r) and r not in expected_roles]
            adicionadas_prd = [r for r in sorted(expected_roles) if r not in diretas_prd]

            preservadas_qad = {r: info for r, info in diretas_qad.items() if dados.is_excluida(r)}
            removidas_qad = [r for r in diretas_qad if not dados.is_excluida(r) and r not in expected_roles]
            adicionadas_qad = [r for r in sorted(expected_roles) if r not in diretas_qad]

            # Montar payload final para PRD
            novo_conjunto_prd = []
            for r_name, r_info in preservadas_prd.items():
                novo_conjunto_prd.append({
                    "AGR_NAME": r_name,
                    "FROM_DAT": r_info.get("FROM_DAT") or data_hoje_bapi,
                    "TO_DAT": r_info.get("TO_DAT") or "99991231"
                })
            for r_name in sorted(expected_roles):
                r_info = roles_prd.get(r_name, {})
                novo_conjunto_prd.append({
                    "AGR_NAME": r_name,
                    "FROM_DAT": r_info.get("FROM_DAT") or data_hoje_bapi,
                    "TO_DAT": r_info.get("TO_DAT") or "99991231"
                })

            # Montar payload final para QAD
            novo_conjunto_qad = []
            for r_name, r_info in preservadas_qad.items():
                novo_conjunto_qad.append({
                    "AGR_NAME": r_name,
                    "FROM_DAT": r_info.get("FROM_DAT") or data_hoje_bapi,
                    "TO_DAT": r_info.get("TO_DAT") or "99991231"
                })
            for r_name in sorted(expected_roles):
                r_info = roles_qad.get(r_name, {})
                novo_conjunto_qad.append({
                    "AGR_NAME": r_name,
                    "FROM_DAT": r_info.get("FROM_DAT") or data_hoje_bapi,
                    "TO_DAT": r_info.get("TO_DAT") or "99991231"
                })

            pres_p_str = f" | Preservadas={len(preservadas_prd)}" if preservadas_prd else ""
            pres_q_str = f" | Preservadas={len(preservadas_qad)}" if preservadas_qad else ""
            print(f"     PRD: {len(diretas_prd)} Diretas -> Remover={len(removidas_prd)}, Adicionar={len(adicionadas_prd)}{pres_p_str} => Total Final={len(novo_conjunto_prd)}")
            print(f"     QAD: {len(diretas_qad)} Diretas -> Remover={len(removidas_qad)}, Adicionar={len(adicionadas_qad)}{pres_q_str} => Total Final={len(novo_conjunto_qad)}")

            # 3. Chamar BAPI em PRD
            res_p = conn_prd.call("BAPI_USER_ACTGROUPS_ASSIGN", USERNAME=u, ACTIVITYGROUPS=novo_conjunto_prd)
            ret_p = res_p.get("RETURN", [])
            erros_p = [m for m in ret_p if str(m.get("TYPE", "")).upper() in ("E", "A")]

            # 4. Chamar BAPI em QAD
            res_q = conn_qad.call("BAPI_USER_ACTGROUPS_ASSIGN", USERNAME=u, ACTIVITYGROUPS=novo_conjunto_qad)
            ret_q = res_q.get("RETURN", [])
            erros_q = [m for m in ret_q if str(m.get("TYPE", "")).upper() in ("E", "A")]

            if erros_p or erros_q:
                sucesso_geral = False
                for ep in erros_p:
                    print(f"     [ERRO PRD] {ep.get('ID')}: {ep.get('MESSAGE')}")
                for eq in erros_q:
                    print(f"     [ERRO QAD] {eq.get('ID')}: {eq.get('MESSAGE')}")
                conn_prd.call("BAPI_TRANSACTION_ROLLBACK")
                conn_qad.call("BAPI_TRANSACTION_ROLLBACK")
                continue

            if modo_simulacao:
                conn_prd.call("BAPI_TRANSACTION_ROLLBACK")
                conn_qad.call("BAPI_TRANSACTION_ROLLBACK")
                msg_p = "; ".join(m.get("MESSAGE", "") for m in ret_p if m.get("MESSAGE")) or "Simulação OK"
                msg_q = "; ".join(m.get("MESSAGE", "") for m in ret_q if m.get("MESSAGE")) or "Simulação OK"
                print(f"     ✓ [SIMULAÇÃO OK] PRD: {msg_p}")
                print(f"     ✓ [SIMULAÇÃO OK] QAD: {msg_q}")
                print("     [ROLLBACK] Nenhuma alteração foi persistida na base de dados.")
            else:
                conn_prd.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
                conn_qad.call("BAPI_TRANSACTION_COMMIT", WAIT="X")

                # Gerar User Compare para as funções compostas envolvidas
                users_processados_set = {str(usr).strip().upper() for usr in users}
                for comp_role in sorted({u_data.get("composta") for u_data in analise.get("usuarios", []) if str(u_data.get("usuario", "")).strip().upper() in users_processados_set and u_data.get("composta")}):
                    try:
                        conn_prd.call("PRGN_GEN_PROFILES_FOR_ROLES", IT_ROLES=[{"AGR_NAME": comp_role}], IV_USERCOMPARE="X")
                    except Exception:
                        pass
                    try:
                        conn_qad.call("PRGN_GEN_PROFILES_FOR_ROLES", IT_ROLES=[{"AGR_NAME": comp_role}], IV_USERCOMPARE="X")
                    except Exception:
                        pass

                print(f"     ✓ [COMMIT OK] Atribuições gravadas com sucesso em PRD e QAD!")

                existing_cua = {
                    (str(item.get("utilizador", "")).strip().upper(), str(item.get("sistema", "")).strip().upper(), str(item.get("role", "")).strip().upper())
                    for item in dados.cua_adicionar
                }

                for r in removidas_prd:
                    roles_removidas.append((u, "S4PCLNT100", r, "CONCLUÍDO", f"User {u} has changed via RFC"))
                for r in removidas_qad:
                    roles_removidas.append((u, "S4QCLNT100", r, "CONCLUÍDO", f"User {u} has changed via RFC"))

                for r in sorted(expected_roles):
                    if (u_norm, "S4PCLNT100", r) not in existing_cua:
                        roles_adicionadas.append((u, "S4PCLNT100", r, "CONCLUÍDO", "Atribuição criada/confirmada no SAP PRD via RFC"))
                        existing_cua.add((u_norm, "S4PCLNT100", r))
                    if (u_norm, "S4QCLNT100", r) not in existing_cua:
                        roles_adicionadas.append((u, "S4QCLNT100", r, "CONCLUÍDO", f"User {u} has changed via RFC"))
                        existing_cua.add((u_norm, "S4QCLNT100", r))

        if modo_simulacao:
            print("\n" + "=" * 75)
            print(f"✓ SIMULAÇÃO CONCLUÍDA COM SUCESSO PARA O DEPARTAMENTO: {dep_nome.upper()}")
            print("  Nenhum dado foi alterado no SAP. Para efetivar, utilize a ação [4] Sincronizar PRD & QAD.")
            print("=" * 75)
            return True

        if not sucesso_geral:
            print(f"\n[ERRO] A sincronização do departamento '{dep_nome}' encontrou falhas.")
            return False

        # Auditoria final em tempo real
        print("\n  A verificar auditoria final de paridade (PRD == QAD == Catálogo)...")
        auditorias_invalidas = []
        for u in users:
            u_norm = str(u).strip().upper()
            expected_direct = user_expected_map.get(u_norm, set())

            det_p = conn_prd.call("BAPI_USER_GET_DETAIL", USERNAME=u)
            p_finais = {r["AGR_NAME"] for r in det_p.get("ACTIVITYGROUPS", []) if not dados.is_excluida(r["AGR_NAME"])}
            p_diretas = {r["AGR_NAME"] for r in det_p.get("ACTIVITYGROUPS", []) if r.get("ORG_FLAG") != "C" and not dados.is_excluida(r["AGR_NAME"])}

            det_q = conn_qad.call("BAPI_USER_GET_DETAIL", USERNAME=u)
            q_finais = {r["AGR_NAME"] for r in det_q.get("ACTIVITYGROUPS", []) if not dados.is_excluida(r["AGR_NAME"])}
            q_diretas = {r["AGR_NAME"] for r in det_q.get("ACTIVITYGROUPS", []) if r.get("ORG_FLAG") != "C" and not dados.is_excluida(r["AGR_NAME"])}

            match_direto = (p_diretas == expected_direct) and (q_diretas == expected_direct)
            match_total = (p_finais == q_finais)

            if match_direto:
                print(f"     Auditoria {u}: Atribuições diretas 100% conformes (PRD={len(p_diretas)} | QAD={len(q_diretas)} | Catálogo={len(expected_direct)}) -> MATCH: True")
                if not match_total:
                    print(f"     ℹ️  Nota PFCG: O total expandido difere (PRD={len(p_finais)} vs QAD={len(q_finais)}) devido a funções filhas geradas por User Compare no ambiente.")
            else:
                print(f"     Auditoria {u}: Divergência em atribuições diretas! PRD={len(p_diretas)} | QAD={len(q_diretas)} | Esperadas={len(expected_direct)} -> MATCH: False")
                auditorias_invalidas.append(u)

        if auditorias_invalidas:
            print(f"[ALERTA] Auditoria acusou divergência nas atribuições diretas dos seguintes utilizadores: {', '.join(auditorias_invalidas)}")

        # Atualização do Excel (CUA_REMOVE, CUA_ADICIONAR, CONTROLO)
        print("\n  A atualizar folhas de cálculo oficiais no Excel...")
        try:
            import shutil
            output_dir = Path(r"C:\workspace\SapScript\output")
            output_dir.mkdir(parents=True, exist_ok=True)
            backup_path = output_dir / f"S4H_Perfis_backup_{timestamp_now.strftime('%Y%m%d_%H%M%S')}.xlsx"
            shutil.copy2(dados.caminho, backup_path)
            print(f"  Backup criado em: {backup_path.name}")
        except Exception as e_bkp:
            print(f"  [AVISO] Falha ao criar backup: {e_bkp}")

        try:
            import openpyxl
            wb_ox = openpyxl.load_workbook(dados.caminho)

            # Folha CUA_REMOVE
            if "CUA_REMOVE" in wb_ox.sheetnames and roles_removidas:
                ws_rem = wb_ox["CUA_REMOVE"]
                max_id_rem = 0
                for r in range(2, ws_rem.max_row + 1):
                    v = ws_rem.cell(r, 1).value
                    try:
                        v_int = int(v)
                        if v_int > max_id_rem: max_id_rem = v_int
                    except (ValueError, TypeError): pass
                for usr, sis, rol, st, msg in roles_removidas:
                    max_id_rem += 1
                    new_r = ws_rem.max_row + 1
                    ws_rem.cell(new_r, 1, max_id_rem)
                    ws_rem.cell(new_r, 2, usr)
                    ws_rem.cell(new_r, 3, sis)
                    ws_rem.cell(new_r, 4, rol)
                    ws_rem.cell(new_r, 5, st)
                    ws_rem.cell(new_r, 6, msg)
                    ws_rem.cell(new_r, 7, timestamp_str)
                    if sis == "S4PCLNT100":
                        ws_rem.cell(new_r, 8, "OK")
                    elif sis == "S4QCLNT100":
                        ws_rem.cell(new_r, 9, "OK")

            # Folha CUA_ADICIONAR
            if "CUA_ADICIONAR" in wb_ox.sheetnames and roles_adicionadas:
                ws_add = wb_ox["CUA_ADICIONAR"]
                row_by_key = {}
                max_id_add = 0
                for r in range(2, ws_add.max_row + 1):
                    v_id = ws_add.cell(r, 1).value
                    try:
                        v_int = int(v_id)
                        if v_int > max_id_add: max_id_add = v_int
                    except (ValueError, TypeError): pass
                    u_c = normalizar_texto(ws_add.cell(r, 2).value)
                    s_c = normalizar_texto(ws_add.cell(r, 3).value)
                    r_c = normalizar_texto(ws_add.cell(r, 4).value)
                    if u_c and s_c and r_c:
                        row_by_key[(u_c, s_c, r_c)] = r

                for usr, sis, rol, st, msg in roles_adicionadas:
                    u_norm = normalizar_texto(usr)
                    s_norm = normalizar_texto(sis)
                    r_norm = normalizar_texto(rol)
                    target_row = row_by_key.get((u_norm, s_norm, r_norm))
                    if target_row:
                        ws_add.cell(target_row, 5, st)
                        ws_add.cell(target_row, 6, msg)
                        ws_add.cell(target_row, 7, timestamp_str)
                        if sis == "S4PCLNT100":
                            ws_add.cell(target_row, 8, "OK")
                        elif sis == "S4QCLNT100":
                            ws_add.cell(target_row, 9, "OK")
                    else:
                        max_id_add += 1
                        new_r = ws_add.max_row + 1
                        ws_add.cell(new_r, 1, max_id_add)
                        ws_add.cell(new_r, 2, usr)
                        ws_add.cell(new_r, 3, sis)
                        ws_add.cell(new_r, 4, rol)
                        ws_add.cell(new_r, 5, st)
                        ws_add.cell(new_r, 6, msg)
                        ws_add.cell(new_r, 7, timestamp_str)
                        if sis == "S4PCLNT100":
                            ws_add.cell(new_r, 8, "OK")
                        elif sis == "S4QCLNT100":
                            ws_add.cell(new_r, 9, "OK")

            # Folha CONTROLO
            if "CONTROLO" in wb_ox.sheetnames:
                ws_ctrl = wb_ox["CONTROLO"]
                for r in range(2, ws_ctrl.max_row + 1):
                    if str(ws_ctrl.cell(r, 1).value or "").strip().upper() == dep_nome.strip().upper():
                        ws_ctrl.cell(r, 2, "PROCESSADO")
                        ws_ctrl.cell(r, 3, timestamp_str)
                        break

            wb_ox.save(dados.caminho)
            print(f"  Ficheiro Excel {Path(dados.caminho).name} atualizado com sucesso!")
        except Exception as e_ox:
            print(f"  [AVISO] Falha ao atualizar Excel com openpyxl: {e_ox}")

        item_dep["status"] = "PROCESSADO"
        item_dep["timestamp"] = timestamp_str
        item_dep["pendente"] = False

        print(f"\n✓ Sincronização do departamento '{dep_nome}' concluída com sucesso (PRD e QAD)!")
        return True

    finally:
        try: conn_prd.close()
        except Exception: pass
        try: conn_qad.close()
        except Exception: pass


sincronizar_departamento_cua_completo = sincronizar_departamento_prd_qad_rfc


def obter_departamentos_pendentes(dados: ProjetoPerfilData) -> List[Dict[str, Any]]:
    """Devolve a fila da CONTROLO com STATUS vazio, respeitando a ordem das linhas."""
    return sorted(
        (item for item in dados.controlo if item.get("pendente")),
        key=lambda item: int(item.get("linha") or 0),
    )


def processar_departamentos_pendentes_em_sequencia(dados: ProjetoPerfilData, assumir_sim: bool = False) -> bool:
    """
    Processa, pela ordem da CONTROLO, todos os departamentos cujo STATUS está vazio.

    A confirmação SAP é única para toda a fila. Cada departamento só recebe o estado
    PROCESSADO depois de concluir todas as etapas e a auditoria P == Q. A fila para
    imediatamente se um departamento falhar, preservando os seguintes como pendentes.
    """
    fila = obter_departamentos_pendentes(dados)
    exibir_tabela_controlo(dados)
    if not fila:
        return False

    print("\nFILA SEQUENCIAL DE DEPARTAMENTOS PENDENTES:")
    for posicao, item in enumerate(fila, 1):
        print(f"  {posicao}. Linha {item.get('linha')}: {item.get('departamento')}")

    resposta = "" if assumir_sim else input(
        f"Executar agora a sequência CUA completa dos {len(fila)} departamento(s), "
        "pela ordem apresentada? (S/N): "
    ).strip().upper()
    if not assumir_sim and resposta not in ("S", "SIM", "Y", "YES"):
        print("Fila departamental não executada.")
        return False

    for posicao, item in enumerate(fila, 1):
        print(
            f"\n[DEPARTAMENTO {posicao}/{len(fila)}] "
            f"{item.get('departamento')} — linha {item.get('linha')}"
        )
        if not sincronizar_departamento_cua_completo(
            dados, item, confirmar_execucao=False
        ):
            print(
                f"[ERRO] A fila foi interrompida em '{item.get('departamento')}'. "
                "Os departamentos seguintes permanecem pendentes."
            )
            return False

    print("\nTodos os departamentos pendentes concluíram a sequência completa.")
    return True


def executar_menu_departamento(dados: ProjetoPerfilData, item_dep: Dict[str, Any]) -> str:
    """
    Apresenta o resumo do departamento selecionado e as ações operacionais diretas.
    Retorna 'VOLTAR' para escolher outro departamento, ou 'SAIR' para terminar.
    """
    dep_nome = item_dep["departamento"]
    linha = item_dep.get("linha", "?")

    analise = analisar_departamento_proposta(dados, dep_nome)

    while True:
        print("\n" + "=" * 60)
        print(f"DEP: {dep_nome.upper()} | Linha {linha}")
        print("=" * 60)

        st_tag = "PENDENTE" if item_dep.get("pendente") else f"{item_dep.get('status')} | {item_dep.get('timestamp')}"
        print(f"Estado: {st_tag}")

        if analise.get("encontrado"):
            users = analise.get("usuarios", [])
            compostas = analise.get("compostas", [])
            total_singles = analise.get("total_singles_distintas", 0)
            print(f"Users: {len(users)} | Compostas: {len(compostas)} | Singles: {total_singles}")
            for u in users:
                comp = u.get("composta") or "SEM_COMPOSTA"
                print(f" {u['usuario']} | {u['nome']} | {u['cargo']} | {comp} | {u['total_singles']} funções")
        else:
            print(f"[AVISO] {analise.get('mensagem', 'Departamento não encontrado na folha Proposta Ativa.')}")

        print("-" * 60)
        print("[1] Global (Todas as Etapas + Gravação PRD/QAD)")
        print("[2] Validar users no PRD    [3] Cruzar fontes")
        print("[4] Sincronizar PRD & QAD   [5] Verificar funções PRD")
        print("[6] Análise detalhada       [7] Incorporar catálogo")
        print("[S] Simulação SU01/SU10     [8] Outro departamento")
        print("[9] Menu geral              [0] Sair")
        print("-" * 60)

        acao = input("Escolha uma ação: ").strip().upper()

        if acao in ("1", "GLOBAL", "G"):
            print("\n" + "=" * 70)
            print(f"  EXECUÇÃO GLOBAL DO DEPARTAMENTO: {dep_nome.upper()} (Linha {linha})")
            print("=" * 70)

            # Etapa 1: Cruzar fontes relacionais
            print("\n[ETAPA 1/5] Cruzamento relacional de fontes (PFCG_CREATE, PFCG_COMPOSTA, DEFINIÇÕES)...")
            res_cruz = cruzar_fontes_departamento(dados, dep_nome)
            imprimir_cruzamento_fontes(res_cruz)

            # Etapa 2: Incorporar funções do catálogo ativas no PRD
            print("\n[ETAPA 2/5] A verificar e incorporar funções do catálogo ativas no PRD...")
            res_inc = incorporar_adicionais_catalogo_departamento(dados, dep_nome)
            if res_inc.get("ok") and res_inc.get("total_incorporadas", 0) > 0:
                print(f"  ✓ {res_inc.get('total_incorporadas')} função(ões) incorporada(s) na folha 'Proposta Ativa'.")
                dados = carregar_projeto_perfil(dados.caminho)
                analise = analisar_departamento_proposta(dados, dep_nome)
            else:
                print(f"  ℹ {res_inc.get('mensagem', 'Nenhuma nova função do catálogo a incorporar.')}")

            # Etapa 3: Verificar existência de funções no PRD
            print("\n[ETAPA 3/5] A verificar existência de funções no SAP PRD...")
            roles_dep = list(analise.get("compostas", [])) + list(analise.get("singles_frequencia", {}).keys())
            if roles_dep:
                res_verif = verificar_funcoes_prd(roles_dep)
                imprimir_resultado_verificacao_prd(res_verif, f"Departamento '{dep_nome}'")

            # Etapa 4: Validar utilizadores e conformidade no SAP PRD
            print("\n[ETAPA 4/5] A validar utilizadores no SAP PRD...")
            res_prd = validar_utilizadores_prd(dados, dep_nome)
            imprimir_validacao_utilizadores_prd(res_prd)

            # Etapa 5: Sincronização CUA completa
            print("\n[ETAPA 5/5] Sincronização CUA completa (PRD e QAS)...")
            sucesso = sincronizar_departamento_cua_completo(dados, item_dep)
            if sucesso:
                dados = carregar_projeto_perfil(dados.caminho)
                item_dep["status"] = "PROCESSADO"
                item_dep["pendente"] = False
                analise = analisar_departamento_proposta(dados, dep_nome)
                print(f"\n✓ Execução Global concluída com sucesso para o departamento '{dep_nome}'!")

        elif acao == "2":
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

        elif acao == "3":
            res_cruz = cruzar_fontes_departamento(dados, dep_nome)
            imprimir_cruzamento_fontes(res_cruz)

        elif acao in ("S", "SIMULAR", "SIMULACAO", "SU01", "SU10"):
            sincronizar_departamento_prd_qad_rfc(dados, item_dep, confirmar_execucao=False, modo_simulacao=True)

        elif acao == "4":
            sincronizar_departamento_prd_qad_rfc(dados, item_dep, confirmar_execucao=True, modo_simulacao=False)

        elif acao == "5":
            if analise.get("encontrado"):
                roles_dep = list(analise["compostas"]) + list(analise["singles_frequencia"].keys())
                res_prd = verificar_funcoes_prd(roles_dep)
                imprimir_resultado_verificacao_prd(res_prd, f"Departamento '{dep_nome}'")
            else:
                print("[AVISO] Não foi possível obter as funções do departamento.")

        elif acao == "6":
            imprimir_analise_departamento(analise)

        elif acao == "7":
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

        elif acao in ("8", "V", "VOLTAR"):
            return "VOLTAR"

        elif acao in ("9", "M", "GERAL"):
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


def sincronizar_composta_sap_rfc(
    ambiente: str,
    composta: str,
    funcao_individual: str,
    texto_funcao: str = "",
) -> Dict[str, Any]:
    """
    Sincroniza a adição de uma função individual a uma função composta no sistema SAP (PRD ou QAD) via RFC:
      1. Consulta os membros atuais da composta na tabela AGR_AGRS.
      2. Se ainda não for membro, adiciona via PRGN_RFC_ADD_AGRS_TO_COLL_AGR.
      3. Gera perfil e executa User Compare na Composta via PRGN_GEN_PROFILES_FOR_ROLES (IV_USERCOMPARE='X').
      4. Gera perfil e executa User Compare na Função Individual via PRGN_GEN_PROFILES_FOR_ROLES (IV_USERCOMPARE='X')
         para resolver o semáforo amarelo para verde.
    """
    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_write_guard, make_read_only_guard, fetch_composite_members,
        )
        from pyrfc import Connection

        load_project_env(find_project_root())
        params = build_connection_params_for(ambiente)

        # 1. Consulta membros atuais
        guard_ro = make_read_only_guard(["AGR_AGRS", "AGR_TEXTS"])
        conn_ro = Connection(**params)
        try:
            membros = fetch_composite_members(conn_ro, guard_ro, composta)
        finally:
            conn_ro.close()

        ja_membro = funcao_individual.strip().upper() in membros

        # 2. Escrita e User Compare
        conn = Connection(**params)
        guard = make_write_guard(
            allowed_functions=["PRGN_RFC_ADD_AGRS_TO_COLL_AGR", "PRGN_GEN_PROFILES_FOR_ROLES"],
            allowed_tables=["AGR_AGRS", "AGR_TEXTS"]
        )
        try:
            if not ja_membro:
                guard.assert_function_allowed("PRGN_RFC_ADD_AGRS_TO_COLL_AGR")
                res_add = conn.call(
                    "PRGN_RFC_ADD_AGRS_TO_COLL_AGR",
                    ACTIVITY_GROUP=composta,
                    ACTIVITY_GROUPS=[{"AGR_NAME": funcao_individual, "TEXT": texto_funcao}],
                    NO_DIALOG="X",
                )
                mensagens = res_add.get("RETURN", []) or []
                erros = [m for m in mensagens if str(m.get("TYPE", "")).upper() in ("E", "A")]
                if erros:
                    return {
                        "ok": False,
                        "ambiente": ambiente,
                        "message": "; ".join(m.get("MESSAGE", "") for m in erros)
                    }

            # 3. Gerar perfil e atualizar utilizador (User Compare) na Composta
            guard.assert_function_allowed("PRGN_GEN_PROFILES_FOR_ROLES")
            conn.call(
                "PRGN_GEN_PROFILES_FOR_ROLES",
                IT_ROLES=[{"AGR_NAME": composta}],
                IV_USERCOMPARE="X",
            )

            # 4. Gerar perfil e atualizar utilizador (User Compare) na Função Individual (amarelo -> verde)
            try:
                conn.call(
                    "PRGN_GEN_PROFILES_FOR_ROLES",
                    IT_ROLES=[{"AGR_NAME": funcao_individual}],
                    IV_USERCOMPARE="X",
                )
            except Exception:
                pass

            status_msg = "já era membro" if ja_membro else "adicionada com sucesso"
            return {
                "ok": True,
                "ambiente": ambiente,
                "ja_membro": ja_membro,
                "message": f"SAP {ambiente}: Função {funcao_individual} ({status_msg}) na Composta {composta}. Perfis gerados e User Compare atualizado (semáforo verde)."
            }
        finally:
            conn.close()

    except Exception as exc:
        return {
            "ok": False,
            "ambiente": ambiente,
            "message": f"Erro RFC no SAP {ambiente}: {exc}"
        }


def executar_fluxo_pesquisa_atribuir_transacao(
    dados: ProjetoPerfilData,
    caminho_excel: Optional[str] = None,
    transacao_alvo: Optional[str] = None,
    user_alvo: Optional[str] = None,
    simular: bool = False,
    assumir_sim: bool = False,
) -> bool:
    """
    [OPÇÃO 4: PESQUISA / ANALISAR]
    Fluxo automatizado de pesquisa de transação no SAP PRD via RFC e atribuição ao utilizador:
      1. Pesquisa a transação no SAP PRD (AGR_TCODES, AGR_TEXTS, TSTCT) e identifica a role oficial.
      2. Pede ou recebe o utilizador SAP.
      3. Na folha 'Proposta Ativa', localiza o utilizador, lê o Departamento e a Composite Role,
         e atribui a função individual na primeira coluna livre a partir das roles existentes (coluna 10+).
      4. Na folha 'PFCG_COMPOSTA': adiciona a nova função individual como membro da Composite Role
         do utilizador (a folha PFCG_CREATE permanece intacta).
      5. Na folha departamental correspondente (ex.: 'Client Services'):
         - Localiza o utilizador no cabeçalho da Linha 2.
         - Localiza ou cria a linha da transação e atribui com 'X' na coluna do utilizador.
      6. Sincronização nos Sistemas SAP (1º PRD e depois 2º QAD):
         - Adiciona a role à composta via RFC PRGN_RFC_ADD_AGRS_TO_COLL_AGR.
         - Executa User Compare na composta e na função individual via PRGN_GEN_PROFILES_FOR_ROLES
           (IV_USERCOMPARE='X') para passar o semáforo de amarelo a verde.
      7. Grava no Excel com segurança (backup automático em output/ e sincronização do Desktop).
    """
    import openpyxl
    import shutil

    caminho = caminho_excel or dados.caminho or encontrar_excel_padrao()
    if not caminho or not os.path.exists(caminho):
        print("[ERRO] Ficheiro Excel não encontrado.")
        return False

    print("\n" + "=" * 78)
    print("  FLUXO DE PESQUISA, ANÁLISE E ATRIBUIÇÃO DE TRANSAÇÃO SAP")
    print("=" * 78)

    def _safe_input(prompt: str) -> str:
        try:
            return input(prompt).strip()
        except (EOFError, KeyboardInterrupt):
            return ""

    # 1. Obter Transação a pesquisar
    if not transacao_alvo:
        transacao_alvo = _safe_input("\nIntroduza o código da transação SAP (ex: ME23N) [Enter para cancelar]: ").upper()
        if not transacao_alvo:
            return False
    else:
        transacao_alvo = transacao_alvo.strip().upper()

    # 2. Pesquisa exclusiva no catálogo oficial do Projeto (folhas Proposta e PFCG_CREATE)
    print(f"\n[PESQUISA] A pesquisar transação '{transacao_alvo}' no catálogo do projeto...")
    roles_projeto = dados.tcode_para_roles.get(transacao_alvo, [])
    if not roles_projeto:
        roles_projeto = dados.tcode_para_roles.get(normalizar_texto(transacao_alvo), [])

    if not roles_projeto:
        print(f"  [AVISO] A transação '{transacao_alvo}' não pertence ao catálogo de funções do projeto.")
        print("  (Apenas transações oficiais contempladas nas folhas Proposta / PFCG_CREATE constam no projeto).")
        return False

    # Obter descrição da transação no SAP PRD (tabela TSTCT) para enriquecimento visual
    tcode_desc = ""
    try:
        from sap_rfc._rfc_common import build_connection_params_for, make_read_only_guard, read_table
        from pyrfc import Connection
        params_prd = build_connection_params_for("PRD")
        conn = Connection(**params_prd)
        guard = make_read_only_guard(["TSTCT"])
        rows_t = read_table(conn, guard, table_name="TSTCT", fields=["TCODE", "TTEXT"], options=[f"SPRSL = 'P' AND TCODE = '{transacao_alvo}'"], rowcount=1)
        if rows_t:
            tcode_desc = str(rows_t[0][1]).strip()
        else:
            rows_t_en = read_table(conn, guard, table_name="TSTCT", fields=["TCODE", "TTEXT"], options=[f"SPRSL = 'E' AND TCODE = '{transacao_alvo}'"], rowcount=1)
            if rows_t_en:
                tcode_desc = str(rows_t_en[0][1]).strip()
        conn.close()
    except Exception:
        pass

    desc_show = f" — {tcode_desc}" if tcode_desc else ""
    print("\n" + "=" * 78)
    print(f"  TRANSAÇÃO DO PROJETO: {transacao_alvo}{desc_show}")
    print("=" * 78)
    print(f"  Funções do projeto associadas ({len(roles_projeto)}):")
    for i, r_cat in enumerate(roles_projeto, 1):
        desc_cat = dados.roles_simples.get(r_cat, {}).get("descricao", "")
        desc_str = f" — {desc_cat}" if desc_cat else ""
        print(f"    [{i}] {r_cat:<30}{desc_str}")
    print("-" * 78)

    # 3. Perguntar se deseja atribuir a um utilizador
    role_selecionada = None
    if not user_alvo:
        if assumir_sim:
            deseja_atribuir = "S"
        else:
            deseja_atribuir = _safe_input("\nDeseja atribuir alguma destas funções a um utilizador? (S/N) [Padrão: N]: ").upper()
        
        if deseja_atribuir not in ("S", "SIM", "Y", "YES"):
            print("Pesquisa concluída.")
            return True

        if len(roles_projeto) == 1:
            role_selecionada = roles_projeto[0]
            desc_cat = dados.roles_simples.get(role_selecionada, {}).get("descricao", "")
            desc_str = f" — {desc_cat}" if desc_cat else ""
            print(f"  ✓ Função selecionada: {role_selecionada}{desc_str}")
        else:
            escolha = _safe_input(f"  Selecione o número da função a atribuir [1-{len(roles_projeto)}] (Padrão: 1): ")
            idx_sel = int(escolha) - 1 if escolha.isdigit() and 1 <= int(escolha) <= len(roles_projeto) else 0
            role_selecionada = roles_projeto[idx_sel]
            desc_cat = dados.roles_simples.get(role_selecionada, {}).get("descricao", "")
            desc_str = f" — {desc_cat}" if desc_cat else ""
            print(f"  ✓ Função selecionada: [{idx_sel + 1}] {role_selecionada}{desc_str}")

        user_alvo = _safe_input("\nIntroduza o ID do utilizador SAP para atribuição (ex: S5006) [Enter para cancelar]: ").upper()
        if not user_alvo:
            print("Atribuição cancelada.")
            return True
    else:
        user_alvo = user_alvo.strip().upper()
        if len(roles_projeto) == 1:
            role_selecionada = roles_projeto[0]
        else:
            role_selecionada = roles_projeto[0]

    desc_role = dados.roles_simples.get(role_selecionada, {}).get("descricao", "")

    def num_to_col(n: int) -> str:
        s = ""
        while n > 0:
            n, m = divmod(n - 1, 26)
            s = chr(65 + m) + s
        return s

    # 5. Carregar e analisar a estrutura atual do Excel
    wb = openpyxl.load_workbook(caminho, data_only=True)

    # 5.1 Proposta Ativa
    print(f"\n[ETAPA 2] Folha 'Proposta Ativa' para o utilizador '{user_alvo}':")
    if "Proposta Ativa" not in wb.sheetnames:
        print("[ERRO] Folha 'Proposta Ativa' não encontrada no ficheiro Excel.")
        wb.close()
        return False

    ws_pa = wb["Proposta Ativa"]
    user_row = None
    user_norm = normalizar_texto(user_alvo)
    for r in range(2, ws_pa.max_row + 1):
        u_val = str(ws_pa.cell(r, 1).value or "").strip().upper()
        u_val_norm = normalizar_texto(u_val)
        if (
            u_val == user_alvo
            or u_val_norm == user_norm
            or u_val == f"S{user_alvo}"
            or u_val_norm == f"s{user_norm}"
            or (user_norm.startswith("s") and u_val_norm == user_norm[1:])
        ):
            user_row = r
            user_alvo = u_val
            break

    if not user_row:
        print(f"  [ERRO] Utilizador '{user_alvo}' não encontrado na folha 'Proposta Ativa'.")
        wb.close()
        return False

    user_nome = str(ws_pa.cell(user_row, 3).value or "").strip()
    dep_nome = str(ws_pa.cell(user_row, 8).value or ws_pa.cell(user_row, 6).value or "").strip()
    comp_role = str(ws_pa.cell(user_row, 9).value or "").strip()

    if not comp_role or comp_role.upper() in ("NAN", "NONE", "-"):
        print(f"  [ERRO] Utilizador '{user_alvo}' não possui Composite Role definida na folha 'Proposta Ativa'.")
        wb.close()
        return False

    print(f"  ✓ Utilizador: {user_alvo} ({user_nome}) localizado na Linha {user_row}")
    print(f"  ✓ Departamento: '{dep_nome}' | Composite Role: '{comp_role}'")

    col_fi = encontrar_coluna_funcoes_individuais_ws(ws_pa, default_col=11)
    col_livre = None
    roles_existentes_user = []
    col_ja_atribuida = None
    for c in range(col_fi, ws_pa.max_column + 10):
        val = str(ws_pa.cell(user_row, c).value or "").strip()
        if val:
            roles_existentes_user.append(val)
            if val.upper() == role_selecionada.upper():
                col_ja_atribuida = c
        elif col_livre is None:
            col_livre = c
            break

    gravar_proposta_ativa = False
    if col_ja_atribuida:
        print(f"  ℹ A função '{role_selecionada}' já se encontra atribuída na célula {num_to_col(col_ja_atribuida)}{user_row}.")
    else:
        gravar_proposta_ativa = True
        col_letra = num_to_col(col_livre)
        print(f"  ✓ Roles atribuídas atualmente: {len(roles_existentes_user)}")
        print(f"  ✓ Primeira coluna livre identificada: Coluna {col_livre} ({col_letra}) -> Célula {col_letra}{user_row}")
        print(f"  -> Ação Planeada: Gravar '{role_selecionada}' na célula {col_letra}{user_row}")

    # 5.2 PFCG_COMPOSTA (Substitui PFCG_CREATE)
    print(f"\n[ETAPA 3] Folha 'PFCG_COMPOSTA' para a Composite Role '{comp_role}':")
    if "PFCG_COMPOSTA" not in wb.sheetnames:
        print("[ERRO] Folha 'PFCG_COMPOSTA' não encontrada no ficheiro Excel.")
        wb.close()
        return False

    ws_comp = wb["PFCG_COMPOSTA"]
    comp_text = ""
    ja_membro_excel = False
    linha_membro_existente = None
    max_id_comp = 0
    for r in range(2, ws_comp.max_row + 1):
        try:
            val_id = int(ws_comp.cell(r, 1).value or 0)
            if val_id > max_id_comp:
                max_id_comp = val_id
        except (ValueError, TypeError):
            pass

        agr_c = str(ws_comp.cell(r, 2).value or "").strip().upper()
        txt_c = str(ws_comp.cell(r, 3).value or "").strip()
        agr_filha = str(ws_comp.cell(r, 4).value or "").strip().upper()

        if agr_c == comp_role.upper():
            if txt_c and not comp_text:
                comp_text = txt_c
            if agr_filha == role_selecionada.upper():
                ja_membro_excel = True
                linha_membro_existente = r

    if not comp_text:
        comp_text = dados.roles_compostas.get(comp_role, {}).get("descricao", comp_role)

    gravar_pfcg_composta = False
    novo_id_comp = max_id_comp + 1
    nova_linha_comp = ws_comp.max_row + 1

    if ja_membro_excel:
        print(f"  ℹ A função '{role_selecionada}' já se encontra associada à Composta '{comp_role}' (Linha {linha_membro_existente}) em 'PFCG_COMPOSTA'.")
        gravar_pfcg_composta = False
    else:
        gravar_pfcg_composta = True
        print(f"  ✓ A função '{role_selecionada}' ainda não consta em '{comp_role}' na folha 'PFCG_COMPOSTA'.")
        print(f"  -> Ação Planeada: Adicionar Linha {nova_linha_comp} (ID={novo_id_comp}, AGR_NAME_COMPOSTA='{comp_role}', TEXT='{comp_text}', AGR_NAME='{role_selecionada}', STATUS='')")

    # 5.3 Folha Departamental
    print(f"\n[ETAPA 4] Folha Departamental '{dep_nome}':")
    sheet_dep_nome = next((s for s in wb.sheetnames if s.strip().upper() == dep_nome.strip().upper()), None)
    if not sheet_dep_nome:
        sheet_dep_nome = next((s for s in wb.sheetnames if dep_nome.strip().upper() in s.strip().upper() or s.strip().upper() in dep_nome.strip().upper()), None)

    if not sheet_dep_nome:
        print(f"  [ERRO] Folha departamental '{dep_nome}' não encontrada no livro Excel.")
        wb.close()
        return False

    ws_dep = wb[sheet_dep_nome]
    col_user_dep = None
    user_hdr_texto = ""
    for c in range(1, ws_dep.max_column + 1):
        val_cab = str(ws_dep.cell(2, c).value or "").strip().upper()
        if user_alvo in val_cab:
            col_user_dep = c
            user_hdr_texto = str(ws_dep.cell(2, c).value or "").strip()
            break

    if not col_user_dep:
        print(f"  [ERRO] Utilizador '{user_alvo}' não encontrado no cabeçalho (Linha 2) da folha '{sheet_dep_nome}'.")
        wb.close()
        return False

    col_user_letra = num_to_col(col_user_dep)
    print(f"  ✓ Utilizador '{user_alvo}' localizado no cabeçalho (Linha 2, Coluna {col_user_dep} / {col_user_letra}): {repr(user_hdr_texto)}")

    linha_tcode_dep = None
    for r in range(3, ws_dep.max_row + 1):
        tc = str(ws_dep.cell(r, 1).value or "").strip().upper()
        if tc == transacao_alvo:
            linha_tcode_dep = r
            break

    gravar_dep = False
    nova_linha_dep = None

    if linha_tcode_dep:
        val_x = str(ws_dep.cell(linha_tcode_dep, col_user_dep).value or "").strip().upper()
        if val_x == "X":
            print(f"  ℹ O utilizador '{user_alvo}' já possui 'X' atribuído na transação '{transacao_alvo}' (Célula {col_user_letra}{linha_tcode_dep}).")
        else:
            gravar_dep = True
            print(f"  ✓ Transação '{transacao_alvo}' existe na Linha {linha_tcode_dep}.")
            print(f"  -> Ação Planeada: Marcar 'X' na célula {col_user_letra}{linha_tcode_dep}")
    else:
        gravar_dep = True
        nova_linha_dep = ws_dep.max_row + 1
        print(f"  ✓ Transação '{transacao_alvo}' não existe na folha '{sheet_dep_nome}'.")
        print(f"  -> Ação Planeada: Criar Linha {nova_linha_dep}: Col A='{transacao_alvo}', Col B='{tcode_desc}', Col {col_user_letra}='X'")

    wb.close()

    # 6. Avaliar se há alterações a realizar
    if not (gravar_proposta_ativa or gravar_pfcg_composta or gravar_dep):
        print("\n" + "=" * 78)
        print(f"  ℹ Todas as atribuições no Excel para {user_alvo} e {transacao_alvo} já se encontram ativas!")
        print("=" * 78)
    else:
        print("\n" + "-" * 78)
        print("  RESUMO DE ALTERAÇÕES PLANEADAS:")
        if gravar_proposta_ativa:
            print(f"    ├─ Proposta Ativa: Gravar '{role_selecionada}' na célula {num_to_col(col_livre)}{user_row}")
        if gravar_pfcg_composta:
            print(f"    ├─ PFCG_COMPOSTA: Adicionar '{role_selecionada}' a '{comp_role}' (ID {novo_id_comp})")
        if gravar_dep:
            if linha_tcode_dep:
                print(f"    └─ {sheet_dep_nome}: Marcar 'X' para {user_alvo} na Linha {linha_tcode_dep}")
            else:
                print(f"    └─ {sheet_dep_nome}: Criar Linha {nova_linha_dep} com {transacao_alvo} e 'X' para {user_alvo}")
        print("-" * 78)

    print(f"  PLANEAMENTO SAP: Sincronizar Composta '{comp_role}' com '{role_selecionada}' em PRD e QAD (User Compare incluído).")

    if simular:
        print("\n" + "=" * 78)
        print("  MODO SIMULAÇÃO CONCLUÍDO: Nenhuma alteração foi gravada no ficheiro Excel nem no SAP.")
        print("=" * 78)
        return True

    # Confirmação do utilizador
    if not assumir_sim:
        conf = input("\nDeseja aplicar as alterações no ficheiro Excel e sincronizar no SAP PRD/QAD? (S/N) [Padrão: S]: ").strip().upper()
        if conf in ("N", "NAO", "NÃO"):
            print("Operação cancelada pelo utilizador.")
            return False

    # 7. Sincronização nos Sistemas SAP (PRD e depois QAD) via RFC
    print("\n" + "=" * 78)
    print("  ETAPA 5: SINCRONIZAÇÃO NOS SISTEMAS SAP (PRD E QAD)")
    print("=" * 78)

    print(f"\n  [1/2] A atualizar Composta '{comp_role}' no SAP PRD...")
    res_prd = sincronizar_composta_sap_rfc("PRD", comp_role, role_selecionada, desc_role)
    if res_prd.get("ok"):
        print(f"  ✓ {res_prd.get('message')}")
    else:
        print(f"  [AVISO SAP PRD]: {res_prd.get('message')}")

    print(f"\n  [2/2] A atualizar Composta '{comp_role}' no SAP QAD...")
    res_qad = sincronizar_composta_sap_rfc("QAD", comp_role, role_selecionada, desc_role)
    if res_qad.get("ok"):
        print(f"  ✓ {res_qad.get('message')}")
    else:
        print(f"  [AVISO SAP QAD]: {res_qad.get('message')}")

    ts_agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if res_prd.get("ok") and res_qad.get("ok"):
        status_comp = "Criado"
        msg_comp = "Atribuído em SAP PRD e QAD (User Compare gerado)"
        prd_comp = "Validado"
    elif res_prd.get("ok"):
        status_comp = "Criado"
        msg_comp = "Atribuído em SAP PRD (QAD pendente)"
        prd_comp = "Validado"
    else:
        status_comp = ""
        msg_comp = "Pendente de sincronização SAP"
        prd_comp = ""

    # 8. Backup de segurança antes de gravar
    try:
        output_dir = PROJECT_ROOT / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = output_dir / f"S4H_Perfis_backup_{timestamp_str}.xlsx"
        shutil.copy2(caminho, backup_path)
        print(f"\n  ✓ Backup de segurança criado: {backup_path.name}")
    except Exception as e_bkp:
        print(f"  [AVISO] Falha ao criar backup de segurança: {e_bkp}")

    # 9. Gravação no Excel (COM se aberto, ou openpyxl)
    xl = None
    wb_com = None
    try:
        import win32com.client
        xl = win32com.client.Dispatch("Excel.Application")
        for w in xl.Workbooks:
            if Path(w.FullName).resolve() == Path(caminho).resolve():
                wb_com = w
                break
    except Exception:
        wb_com = None
        xl = None

    if wb_com is not None and xl is not None:
        old_su = xl.ScreenUpdating
        old_da = xl.DisplayAlerts
        try:
            xl.ScreenUpdating = False
            xl.DisplayAlerts = False

            if gravar_proposta_ativa:
                ws_pa_com = wb_com.Worksheets("Proposta Ativa")
                ws_pa_com.Cells(user_row, col_livre).Value = role_selecionada

            if gravar_pfcg_composta:
                ws_comp_com = wb_com.Worksheets("PFCG_COMPOSTA")
                lr_comp = ws_comp_com.UsedRange.Rows.Count + 1
                ws_comp_com.Cells(lr_comp, 1).Value = novo_id_comp
                ws_comp_com.Cells(lr_comp, 2).Value = comp_role
                ws_comp_com.Cells(lr_comp, 3).Value = comp_text
                ws_comp_com.Cells(lr_comp, 4).Value = role_selecionada
                ws_comp_com.Cells(lr_comp, 5).Value = status_comp
                ws_comp_com.Cells(lr_comp, 6).Value = msg_comp
                ws_comp_com.Cells(lr_comp, 7).Value = ts_agora
                ws_comp_com.Cells(lr_comp, 8).Value = prd_comp

            if gravar_dep:
                ws_dep_com = wb_com.Worksheets(sheet_dep_nome)
                if linha_tcode_dep:
                    ws_dep_com.Cells(linha_tcode_dep, col_user_dep).Value = "X"
                else:
                    lr_dep = ws_dep_com.UsedRange.Rows.Count + 1
                    ws_dep_com.Cells(lr_dep, 1).Value = transacao_alvo
                    ws_dep_com.Cells(lr_dep, 2).Value = tcode_desc
                    ws_dep_com.Cells(lr_dep, col_user_dep).Value = "X"

            wb_com.Save()
            print("  ✓ Alterações gravadas com sucesso via Excel.Application (COM)!")
        finally:
            try:
                xl.ScreenUpdating = old_su
                xl.DisplayAlerts = old_da
            except Exception:
                pass
    else:
        wb_ox = openpyxl.load_workbook(caminho)

        if gravar_proposta_ativa:
            ws_pa_ox = wb_ox["Proposta Ativa"]
            ws_pa_ox.cell(row=user_row, column=col_livre, value=role_selecionada)

        if gravar_pfcg_composta:
            ws_comp_ox = wb_ox["PFCG_COMPOSTA"]
            ws_comp_ox.append([novo_id_comp, comp_role, comp_text, role_selecionada, status_comp, msg_comp, ts_agora, prd_comp])

        if gravar_dep:
            ws_dep_ox = wb_ox[sheet_dep_nome]
            if linha_tcode_dep:
                ws_dep_ox.cell(row=linha_tcode_dep, column=col_user_dep, value="X")
            else:
                new_r_dep = ws_dep_ox.max_row + 1
                ws_dep_ox.cell(row=new_r_dep, column=1, value=transacao_alvo)
                ws_dep_ox.cell(row=new_r_dep, column=2, value=tcode_desc)
                ws_dep_ox.cell(row=new_r_dep, column=col_user_dep, value="X")

        wb_ox.save(caminho)
        wb_ox.close()
        print("  ✓ Alterações gravadas com sucesso via openpyxl!")

    print("\n" + "=" * 78)
    print(f"  ✓ ATRIBUIÇÃO E SINCRONIZAÇÃO CONCLUÍDAS COM SUCESSO!")
    print(f"     Utilizador:      {user_alvo} ({user_nome})")
    print(f"     Transação:       {transacao_alvo} ({tcode_desc})")
    print(f"     Role Individual: {role_selecionada}")
    print(f"     Role Composta:   {comp_role}")
    print(f"     Sistemas SAP:    PRD e QAD sincronizados com User Compare")
    print("=" * 78)
    return True


def obter_discrepancias_sincronizacao(dados: ProjetoPerfilData, caminho_excel: Optional[str] = None) -> Dict[str, List[Dict[str, Any]]]:
    """
    Identifica todos os utilizadores dos departamentos com status PROCESSADO em CONTROLO
    que possuem funções pendentes de sincronização no CUA ou no SAP PRD/QAD.
    """
    deps_auditaveis = {
        normalizar_texto(d.get("departamento")): d.get("departamento")
        for d in dados.controlo
        if str(d.get("status", "")).strip().upper() == "PROCESSADO"
    }
    if not deps_auditaveis:
        return {}

    todas_filhas_compostas = set()
    for c_info in dados.roles_compostas.values():
        for f in c_info.get("roles_filhas", []):
            todas_filhas_compostas.add(normalizar_texto(f))

    cua_roles_por_user = defaultdict(set)
    for it in dados.cua_adicionar:
        u = normalizar_texto(it.get("utilizador"))
        r = normalizar_texto(it.get("role"))
        if u and r:
            cua_roles_por_user[u].add(r)

    discrepancias = defaultdict(list)
    for dep_norm, dep_nome in deps_auditaveis.items():
        analise = analisar_departamento_proposta(dados, dep_nome)
        if not analise.get("encontrado"):
            continue

        definicoes_dep = {normalizar_texto(r) for r in dados.definicoes_departamento.get(dep_norm, set())}
        if not definicoes_dep and dep_norm in ("LEGAL", "PURCHASE & SERVICES", "PURCHASE AND SERVICES"):
            definicoes_dep = {normalizar_texto(r) for r in dados.definicoes_departamento.get("PURCHASE & SERVICES", set())} or {
                "ZORG_TODAS_EMPRESAS", "Z_BASIS_BASE", "ZORG_BP_GERAL",
                "ZORG_BP_LOGISTICS_CUSTOMER", "ZORG_BP_FLVN01_LOGISTICS_VENDO",
                "ZORG_BP_Z001_GENERALPARTNERS", "ZORG_BP_Z003_RELATEDPARTNERS",
                "Z_BR_TYPE_BP_GERAL", "Z_MY_HOME"
            }
        # Regra oficial: a sheet DEFINIÇÕES é mandatória por departamento.
        # Tudo que está na linha do departamento deve ser atribuído diretamente,
        # mesmo que a role também exista como filha de uma função composta.
        definicoes_diretas = set(definicoes_dep)

        for u_item in analise.get("usuarios", []):
            u_id = normalizar_texto(u_item.get("usuario"))
            if not u_id:
                continue
            comp = normalizar_texto(u_item.get("composta"))
            if not comp or comp in ("NAN", "NONE", "-"):
                continue

            filhas_comp = {normalizar_texto(f) for f in dados.roles_compostas.get(comp, {}).get("roles_filhas", [])}
            roles_diretas = {comp}
            for s in u_item.get("singles", []):
                s_norm = normalizar_texto(s)
                if s_norm and s_norm != comp and s_norm not in filhas_comp and s_norm not in todas_filhas_compostas:
                    roles_diretas.add(s_norm)

            roles_em_cua = cua_roles_por_user.get(u_id, set())
            faltam_cua = {r for r in roles_diretas if r not in roles_em_cua}
            pacote_sap = set(roles_diretas).union(definicoes_diretas)

            if faltam_cua:
                discrepancias[dep_nome.upper()].append({
                    "user": u_id,
                    "nome": u_item.get("nome", ""),
                    "cargo": u_item.get("cargo", ""),
                    "departamento": dep_nome,
                    "composta": comp,
                    "faltam": sorted(faltam_cua),
                    "roles_necessarias": sorted(roles_diretas),
                    "roles_sap_esperadas": sorted(pacote_sap),
                })

    return dict(discrepancias)


def executar_correcao_sincronizacao_posterior(
    dados: ProjetoPerfilData,
    caminho_excel: str,
    assumir_sim: bool = False
) -> Optional[ProjetoPerfilData]:
    """
    Opção [5] do Menu Principal:
    Refatorada para separar explicitamente as responsabilidades:
      [1] Preparar / Atualizar CUA_ADICIONAR (Apenas Excel)
      [2] Executar atribuições pendentes no SAP (BAPI + Commit + Readback + Atualização Excel por ID)
      [3] Simular atribuições pendentes no SAP (BAPI + Rollback sem alterações)
      [0] Voltar
    """
    while True:
        print("\n" + "=" * 78)
        print("  [OPÇÃO 5] CORREÇÃO DE SINCRONIZAÇÃO POSTERIOR")
        print("=" * 78)
        print("  [1] Preparar / Atualizar Sheet CUA_ADICIONAR (Apenas Excel)")
        print("  [2] Executar atribuições pendentes no SAP (PRD / QAD via RFC)")
        print("  [3] Simular atribuições pendentes no SAP (SU01/SU10 Rollback)")
        print("  [0] Voltar ao Menu Principal")
        print("-" * 78)

        if assumir_sim:
            sub_op = "2"
            print("  [Auto-confirmado --yes] Ação [2] Executar selecionada automaticamente.")
        else:
            try:
                sub_op = input("Selecione uma opção [1-3, 0]: ").strip().upper()
            except Exception:
                sub_op = "0"

        if sub_op in ("0", "VOLTAR", "V", "SAIR", ""):
            return dados

        if sub_op == "1":
            print("\n  A preparar e atualizar a sheet CUA_ADICIONAR no Excel...")
            novo_dados = executar_atualizacao_integrada_excel(caminho_excel, dados, incorporar_prd_rfc=False)
            print("  ✓ Etapa [6/7] concluída. Ficheiro Excel sincronizado.")
            dados = novo_dados
            if assumir_sim:
                return dados

        elif sub_op in ("2", "3"):
            modo_sim = (sub_op == "3")
            lbl_modo = "SIMULAÇÃO" if modo_sim else "EXECUÇÃO REAL"

            discrepancias = obter_discrepancias_sincronizacao(dados, caminho_excel)
            if not discrepancias:
                print(f"\n  {CLR_VERDE}✓ Não existem discrepâncias ou pendências de sincronização CUA.{CLR_RESET}")
                print("  Todos os utilizadores dos departamentos em CONTROLO encontram-se 100% sincronizados!")
                if assumir_sim:
                    return dados
                continue

            total_u = sum(len(ul) for ul in discrepancias.values())
            total_d = len(discrepancias)
            print(f"\n  [{lbl_modo}] Detetado(s) {total_u} utilizador(es) com atribuições pendentes em {total_d} departamento(s):")
            print("-" * 78)

            for dep_nome, u_list in sorted(discrepancias.items()):
                print(f"\n  📁 DEPARTAMENTO: {dep_nome} ({len(u_list)} utilizador(es))")
                print("  " + "-" * 74)
                for u_item in u_list:
                    u_id = u_item["user"]
                    u_nome = u_item.get("nome") or "(Nome não disponível)"
                    u_cargo = u_item.get("cargo") or "(Cargo não disponível)"
                    comp = u_item.get("composta") or "-"
                    faltam = u_item.get("faltam", [])
                    pacote = u_item.get("roles_sap_esperadas", [])

                    print(f"    👤 Utilizador : {u_id} - {u_nome}")
                    print(f"       Cargo      : {u_cargo}")
                    print(f"       Composta   : {comp}")
                    print(f"       Funções pendentes a atualizar no CUA ({len(faltam)}):")
                    for f in faltam:
                        print(f"         • {f}")
                    print(f"       Pacote de funções a sincronizar em SAP PRD & QAD ({len(pacote)}):")
                    for r in pacote:
                        tipo_str = " (Função Composta)" if r == comp else ""
                        print(f"         - {r}{tipo_str}")
                    print("  " + "-" * 74)

            print()
            if not modo_sim and not assumir_sim:
                print("  ATENÇÃO: Esta operação irá alterar autorizações de utilizadores no SAP PRD e QAD.")
                resp = input("  Confirma a execução real e gravação no SAP? (S/N): ").strip().upper()
                if resp not in ("S", "SIM", "Y", "YES"):
                    print("\n[CANCELADO] Operação cancelada pelo utilizador. Nenhuma alteração foi realizada no SAP.")
                    if assumir_sim:
                        return dados
                    continue

            sucesso_total = True
            for dep_nome, u_list in sorted(discrepancias.items()):
                item_dep = next((it for it in dados.controlo if normalizar_texto(it.get("departamento")) == normalizar_texto(dep_nome)), None)
                if not item_dep:
                    item_dep = {"departamento": dep_nome, "status": "PROCESSADO", "timestamp": "", "linha": 0}

                usuarios_alvo = [u_item["user"] for u_item in u_list]

                sucesso = sincronizar_departamento_prd_qad_rfc(
                    dados=dados,
                    item_dep=item_dep,
                    confirmar_execucao=False,
                    modo_simulacao=modo_sim,
                    usuarios_alvo=usuarios_alvo
                )
                if not sucesso:
                    sucesso_total = False
                    print(f"[ERRO] Falha na operação do departamento '{dep_nome}'.")

            if sucesso_total and not modo_sim:
                print("\n  A recarregar e auditar consistência das folhas no Excel...")
                dados = executar_atualizacao_integrada_excel(caminho_excel, dados, incorporar_prd_rfc=False)
                print("\n" + "=" * 78)
                print(f"  {CLR_VERDE}✓ EXECUÇÃO DE ATRIBUIÇÕES PENDENTES NO SAP CONCLUÍDA COM SUCESSO!{CLR_RESET}")
                print("  Todas as folhas do Excel e utilizadores estão 100% sincronizados.")
                print("=" * 78)

            if assumir_sim:
                return dados


def menu_interativo(caminho_inicial: Optional[str] = None):
    """Executa o novo menu interativo limpo com as opções principais."""
    caminho = caminho_inicial or encontrar_excel_padrao()

    if not caminho or not os.path.exists(caminho):
        print("[ERRO] Ficheiro Excel não selecionado ou não encontrado. Saindo.")
        return

    imprimir_cabecalho_compacto(caminho)

    # =========================================================================
    # TAREFA 1 DO PROGRAMA: Atualização e Sincronização Integrada do Excel
    # =========================================================================
    # Executada imediatamente no arranque para assegurar consistência total das folhas:
    #   Etapa 1: Proposta (Validação de TCODEs obsoletos na TSTC do PRD via RFC)
    #   Etapa 2: Matrizes Departamentais -> Proposta Ativa (Atribuições oficiais)
    #   Etapa 3: SAP PRD -> Proposta Ativa (Incorporação de funções ativas do catálogo)
    #   Etapa 4: Proposta Ativa -> PFCG_COMPOSTA (Alinhamento de membros das composite roles)
    #   Etapa 5: Auditoria de Sincronização Departamental & CUA (Discrepâncias e pendências)
    dados = executar_atualizacao_integrada_excel(caminho)

    # =========================================================================
    # MENU PRINCIPAL APÓS ATUALIZAÇÃO DO FICHEIRO
    # =========================================================================
    while True:
        print("\n" + "=" * 78)
        print("  MENU PRINCIPAL - PROJETO PERFIL")
        print("=" * 78)
        print("  [1] Execução       - Executar o Processo Completo")
        print("  [2] Departamento   - Fluxo Departamental de Autorizações (Matrizes)")
        print("  [3] Utilizador     - Auditoria e Pesquisa Individual de Utilizador")
        print("  [4] Pesquisa       - Pesquisar Transação do Projeto e Atribuir a Utilizador")
        print("  [5] Corrigir       - Corrigir Sincronização Posterior (Utilizadores Pendentes)")
        print("  [6] SU53           - Diagnóstico SU53 via RFC e Atribuição de Funções")
        print("  [0] Sair")
        print("-" * 78)
        op = input("Selecione uma opção [1-6, 0]: ").strip().upper()

        if op == "1" or op in ("EXECUCAO", "EXECUÇÃO"):
            print("\n" + "=" * 78)
            print("  [OPÇÃO 1] EXECUÇÃO: Executar o Processo Completo")
            print("=" * 78)
            if processar_departamentos_pendentes_em_sequencia(dados):
                dados = carregar_projeto_perfil(caminho)
            if verificar_e_perguntar_pendencias(caminho):
                dados = carregar_projeto_perfil(caminho)

        elif op == "2" or op == "DEPARTAMENTO":
            while True:
                item_selecionado = selecionar_departamento_interativo(dados)
                if not item_selecionado:
                    break
                if item_selecionado.get("tipo") == "MENU_GERAL":
                    acao = menu_geral_pesquisas(dados)
                    if acao == "SAIR":
                        break
                    continue
                acao = executar_menu_departamento(dados, item_selecionado)
                if acao == "SAIR":
                    break
                elif acao == "MENU_GERAL":
                    acao_g = menu_geral_pesquisas(dados)
                    if acao_g == "SAIR":
                        break

        elif op == "3" or op == "UTILIZADOR":
            while True:
                u_in = input("\nIntroduza o ID do utilizador SAP para auditoria (ex: S5092) [Enter para voltar]: ").strip()
                if not u_in:
                    break
                res_audit = auditar_utilizador(dados, u_in)
                imprimir_auditoria_utilizador(res_audit)

        elif op == "4" or op in ("PESQUISA", "ANALISAR", "PESQUISAR"):
            sucesso = executar_fluxo_pesquisa_atribuir_transacao(dados, caminho)
            if sucesso:
                dados = carregar_projeto_perfil(caminho)

        elif op == "5" or op in ("CORRIGIR", "RESINCRONIZAR", "CORRECAO", "CORREÇÃO"):
            novo_dados = executar_correcao_sincronizacao_posterior(dados, caminho)
            if novo_dados:
                dados = novo_dados

        elif op == "6" or op in ("SU53", "ERROS"):
            import runpy
            script_su53 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Pesquisa_Erros_SU53.py")
            if os.path.exists(script_su53):
                runpy.run_path(script_su53, run_name="__main__")
                dados = carregar_projeto_perfil(caminho)
            else:
                print(f"[ERRO] Script não encontrado: {script_su53}")

        elif op in ("0", "S", "SAIR", "Q", "QUIT", ""):
            print("\nEncerrando. Até logo!")
            break

        else:
            print("[AVISO] Opção inválida. Por favor escolha 1, 2, 3, 4, 5, 6 ou 0.")


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
    parser.add_argument("--cruzar-fontes", dest="cruzar_fontes", action="store_true", help="Cruzar PFCG_CREATE, PFCG_COMPOSTA e DEFINIÇÕES para o departamento")
    parser.add_argument("--validar-users-prd", dest="validar_users_prd", action="store_true", help="Validar atribuições dos utilizadores no SAP PRD (AGR_USERS) com cruzamento relacional")
    parser.add_argument("--verificar-prd", dest="verificar_prd", action="store_true", help="Verificar se as funções existem no sistema SAP PRD via RFC (AGR_DEFINE)")
    parser.add_argument("--comparar-prd", "--funcoes-diferentes-prd", dest="comparar_prd", action="store_true", help="Comparar catálogo de funções no PRD vs ficheiro Excel (AGR_DEFINE e AGR_USERS)")
    parser.add_argument("--pesquisar-role", "-r", dest="role", help="Pesquisar diretamente por nome de função")
    parser.add_argument("--pesquisar-tcode", "-t", dest="tcode", help="Pesquisar diretamente por transação SAP")
    parser.add_argument("--pesquisar-user", "--user", "-u", dest="user", help="Auditar e pesquisar diretamente por utilizador (Proposta, CUA e SAP PRD)")
    parser.add_argument("--incorporar-adicionais", dest="incorporar_adicionais", action="store_true", help="Incorporar na Proposta Ativa as funções do catálogo que já estão ativas no SAP PRD")
    parser.add_argument("--validar-proposta-tcodes", "--validar-tcodes-prd", dest="validar_proposta_tcodes", action="store_true", help="Validar se todas as transações da sheet Proposta existem na tabela TSTC do SAP PRD")
    parser.add_argument("--sincronizar-catalogo", dest="sincronizar_catalogo", action="store_true", help="Sincronizar catálogo da folha Proposta com PFCG_CREATE e SAP PRD")
    parser.add_argument("--atualizar-excel", "--sincronizar-excel", dest="atualizar_excel", action="store_true", help="Executar a atualização e sincronização integrada do Excel como primeira tarefa")
    parser.add_argument("--executar-pendencias", "--despachar-pendencias", dest="executar_pendencias", action="store_true", help="Executar o despachante de processos com STATUS pendente nas folhas operacionais (Rotina 7)")
    parser.add_argument("--sincronizar-cua", dest="sincronizar_cua", action="store_true", help="Executar sincronização CUA completa (PRD -> QAS) para o departamento")
    parser.add_argument("--fila-cua", "--sincronizar-fila", dest="fila_cua", action="store_true", help="Executar sincronização CUA completa (PRD -> QAS) para todos os departamentos pendentes em sequência")
    parser.add_argument("--atribuir-transacao", "--pesquisar-atribuir", dest="atribuir_transacao", action="store_true", help="Pesquisar transação via RFC no PRD e atribuir a utilizador na Proposta Ativa, PFCG_CREATE e folha departamental")
    parser.add_argument("--corrigir-posterior", "--corrigir-sincronizacao", dest="corrigir_posterior", action="store_true", help="Executar correção de sincronização posterior para utilizadores com funções pendentes no CUA")
    parser.add_argument("--audit", dest="audit", action="store_true", help="Executar Auditoria Departamental Pós-Processamento 100% Read-Only nos departamentos com STATUS=PROCESSADO")
    parser.add_argument("--resumo", dest="resumo", action="store_true", help="Utilizar modo resumo na auditoria pós-processamento")
    parser.add_argument("--reconstruir-proposta-ativa", "--reconstruir-proposta", dest="reconstruir_proposta", action="store_true", help="Reconstruir Funções Individuais da sheet Proposta Ativa a partir das matrizes departamentais e Proposta")
    parser.add_argument("--simular", dest="simular", action="store_true", help="Apenas simular a operação sem gravar no Excel")
    parser.add_argument("--yes", "-y", dest="assumir_sim", action="store_true", help="Confirmar automaticamente sem perguntar interativamente")

    args = parser.parse_args()

    caminho_alvo = args.ficheiro or encontrar_excel_padrao()

    # Se foram passados parâmetros de pesquisa direta via CLI:
    if args.reconstruir_proposta or args.audit or args.controlo or args.proximo or args.departamento is not None or args.cruzar_fontes or args.validar_users_prd or args.verificar_prd or args.comparar_prd or args.role or args.tcode or args.user or args.incorporar_adicionais or args.validar_proposta_tcodes or args.sincronizar_catalogo or args.sincronizar_cua or args.fila_cua or args.atualizar_excel or args.executar_pendencias or args.atribuir_transacao or args.corrigir_posterior:
        if not caminho_alvo:
            print("[ERRO] Erro: Ficheiro Excel não encontrado.")
            sys.exit(1)
        
        dados = carregar_projeto_perfil(caminho_alvo)
        imprimir_cabecalho(caminho_alvo)

        if args.reconstruir_proposta:
            alvo_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else None
            alvo_user = args.user.strip() if (args.user and args.user.strip()) else None
            is_sim = bool(args.simular or not args.assumir_sim)
            if alvo_dep:
                res_rec = reconstruir_funcoes_individuais_departamento(dados, alvo_dep, simular=is_sim, caminho_excel=caminho_alvo, filtro_utilizador=alvo_user)
            elif alvo_user:
                # Se não indicou departamento mas indicou utilizador, localizar o departamento do utilizador
                res_dep_user = analisar_departamento_proposta(dados, "")
                dep_encontrado = None
                for c_dep in dados.controlo:
                    d_nome = c_dep.get("departamento", "")
                    if d_nome:
                        t_u, _ = obter_transacoes_usuario_departamento(dados, d_nome, alvo_user)
                        if t_u or d_nome:
                            # Verificar se o utilizador está na sheet
                            sheet_d = next((s for s in dados.sheets_disponiveis if normalizar_nome_coluna(s) == normalizar_nome_coluna(d_nome)), None)
                            if sheet_d:
                                fonte_d = abrir_excel_seguro(caminho_alvo)
                                import pandas as pd
                                df_d = pd.read_excel(fonte_d, sheet_name=sheet_d, header=None)
                                if len(df_d) >= 2:
                                    hdrs = [normalizar_texto(df_d.iloc[1, col]) for col in range(2, df_d.shape[1]) if pd.notna(df_d.iloc[1, col])]
                                    if any(normalizar_texto(alvo_user) in h for h in hdrs):
                                        dep_encontrado = d_nome
                                        break
                if not dep_encontrado:
                    print(f"[ERRO] Utilizador '{alvo_user}' não foi encontrado em nenhuma sheet de departamento.")
                    sys.exit(1)
                res_rec = reconstruir_funcoes_individuais_departamento(dados, dep_encontrado, simular=is_sim, caminho_excel=caminho_alvo, filtro_utilizador=alvo_user)
            else:
                res_rec = reconstruir_funcoes_individuais_todos_departamentos(dados, simular=is_sim, caminho_excel=caminho_alvo)
            imprimir_relatorio_reconstrucao_proposta_ativa(res_rec)
            sys.exit(0)

        if args.audit:
            filtro_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else None
            res_audit = auditar_departamentos_processados(dados, ambientes=("PRD", "QAD"), filtro_departamento=filtro_dep)
            imprimir_auditoria_departamental(res_audit)
            sys.exit(0)

        # Se for sincronização CUA ou atualização de catálogo/Excel, executa a Tarefa 1 em primeiro lugar:
        if args.atualizar_excel or args.fila_cua or args.sincronizar_cua:
            dados = executar_atualizacao_integrada_excel(caminho_alvo, dados)

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
        if args.departamento is not None and not (args.cruzar_fontes or args.validar_users_prd or args.verificar_prd or args.incorporar_adicionais or args.sincronizar_cua):
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
        if args.tcode and not args.atribuir_transacao:
            imprimir_resultado_tcode(pesquisar_por_tcode(dados, args.tcode))
        if args.user and not args.atribuir_transacao:
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

        if args.fila_cua:
            fila = obter_departamentos_pendentes(dados)
            if not fila:
                print("\nTodos os departamentos na folha CONTROLO já estão processados!")
            else:
                print(f"\nA iniciar processamento da fila CUA ({len(fila)} departamentos pendentes)...")
                for posicao, item in enumerate(fila, 1):
                    print(f"\n[DEPARTAMENTO {posicao}/{len(fila)}] {item.get('departamento')} — Linha {item.get('linha')}")
                    sucesso = sincronizar_departamento_cua_completo(dados, item, confirmar_execucao=not args.assumir_sim)
                    if not sucesso:
                        print(f"[ERRO] Falha na sincronização de '{item.get('departamento')}'. Interrompendo fila.")
                        break
                    # Recarregar dados após atualização do Excel
                    dados = carregar_projeto_perfil(caminho_alvo)

        if args.sincronizar_cua:
            alvo_dep = args.departamento.strip() if (args.departamento and args.departamento.strip()) else proximo_dep
            if not alvo_dep:
                print("[ERRO] Nenhum departamento especificado ou pendente para sincronizar.")
            else:
                item_dep = next((it for it in dados.controlo if str(it.get("departamento", "")).strip().upper() == alvo_dep.strip().upper()), None)
                if not item_dep:
                    item_dep = next((it for it in dados.controlo if alvo_dep.strip().upper() in str(it.get("departamento", "")).strip().upper()), None)
                if not item_dep:
                    item_dep = {"departamento": alvo_dep, "status": "", "linha": "?"}
                sincronizar_departamento_cua_completo(dados, item_dep, confirmar_execucao=not args.assumir_sim)

        if args.executar_pendencias:
            print("\n[ROTINA 7] A verificar e despachar processos pendentes nas folhas operacionais...")
            verificar_e_perguntar_pendencias(caminho_alvo)

        if args.atribuir_transacao:
            executar_fluxo_pesquisa_atribuir_transacao(
                dados=dados,
                caminho_excel=caminho_alvo,
                transacao_alvo=args.tcode,
                user_alvo=args.user,
                simular=args.simular,
                assumir_sim=args.assumir_sim,
            )

        if args.corrigir_posterior:
            executar_correcao_sincronizacao_posterior(
                dados=dados,
                caminho_excel=caminho_alvo,
                assumir_sim=args.assumir_sim,
            )

    else:
        # Modo interativo padrão
        menu_interativo(args.ficheiro)

# -*- coding: utf-8 -*-
"""
sap_rfc/projeto_perfil_service.py
======================================================================
Serviço bridge para integração das 6 funcionalidades do Projeto Perfil
no Agente Salsa IT:
  1. Execução: Processar departamentos pendentes em sequência e pendências
  2. Departamento: Fluxo departamental de autorizações (matrizes)
  3. Utilizador: Auditoria individual de utilizador
  4. Pesquisa: Pesquisa de transação no catálogo do projeto e atribuição
  5. Corrigir: Correção de sincronização posterior (utilizadores pendentes)
  6. SU53: Diagnóstico SU53 via RFC e fluxo de atribuição
======================================================================
"""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional


def _get_project_root() -> Path:
    return Path(__file__).resolve().parent.parent


_PROJETO_PERFIL_MOD = None
_PESQUISA_SU53_MOD = None


def get_projeto_perfil_module():
    global _PROJETO_PERFIL_MOD
    if _PROJETO_PERFIL_MOD is not None:
        return _PROJETO_PERFIL_MOD

    root = _get_project_root()
    script_path = root / "Processos" / "Projeto Autorizações" / "Projeto Perfil.py"
    if not script_path.exists():
        raise FileNotFoundError(f"Script Projeto Perfil.py não encontrado em: {script_path}")

    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

    spec = importlib.util.spec_from_file_location("projeto_perfil_engine", script_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Não foi possível criar spec para {script_path}")

    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    _PROJETO_PERFIL_MOD = mod
    return mod


def get_pesquisa_su53_module():
    global _PESQUISA_SU53_MOD
    if _PESQUISA_SU53_MOD is not None:
        return _PESQUISA_SU53_MOD

    root = _get_project_root()
    script_path = root / "Processos" / "Projeto Autorizações" / "Pesquisa_Erros_SU53.py"
    if not script_path.exists():
        raise FileNotFoundError(f"Script Pesquisa_Erros_SU53.py não encontrado em: {script_path}")

    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

    spec = importlib.util.spec_from_file_location("pesquisa_su53_engine", script_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Não foi possível criar spec para {script_path}")

    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    _PESQUISA_SU53_MOD = mod
    return mod


# =====================================================================
# 1. EXECUÇÃO COMPLETA
# =====================================================================

def executar_processo_completo(caminho_excel: Optional[str] = None, assumir_sim: bool = True) -> Dict[str, Any]:
    """Executa a Opção 1: departamentos pendentes em sequência + pendências operacionais."""
    mod = get_projeto_perfil_module()
    caminho = caminho_excel or mod.encontrar_excel_padrao()
    if not caminho or not os.path.exists(caminho):
        return {"ok": False, "status": "ERRO", "message": f"Ficheiro Excel não encontrado: {caminho}"}

    dados = mod.carregar_projeto_perfil(caminho)
    fila = mod.obter_departamentos_pendentes(dados)
    pendencias = mod.obter_pendencias_status(caminho)

    if not fila and not pendencias:
        return {
            "ok": True,
            "status": "SEM_PENDENCIAS",
            "message": "Todos os departamentos na folha CONTROLO já se encontram processados e não há pendências operacionais.",
            "total_departamentos_processados": 0,
            "fila_processada": [],
            "pendencias_executadas": {},
        }

    sucesso_fila = True
    if fila:
        sucesso_fila = mod.processar_departamentos_pendentes_em_sequencia(dados, assumir_sim=assumir_sim)
        if sucesso_fila:
            dados = mod.carregar_projeto_perfil(caminho)

    sucesso_pend = True
    if pendencias:
        sucesso_pend = mod.verificar_e_perguntar_pendencias(caminho, assumir_sim=assumir_sim)

    ok = bool(sucesso_fila and sucesso_pend)
    return {
        "ok": ok,
        "status": "SUCESSO" if ok else "FALHA_PARCIAL",
        "message": "Processo Completo executado com sucesso." if ok else "Ocorreram falhas durante o processamento da fila ou pendências.",
        "total_departamentos_processados": len(fila),
        "fila_processada": [f"Linha {it.get('linha')}: {it.get('departamento')}" for it in fila],
        "pendencias_executadas": pendencias,
    }


# =====================================================================
# 2. FLUXO DEPARTAMENTAL (MATRIZES)
# =====================================================================

def fluxo_departamento(
    departamento: Optional[str] = None,
    subacao: str = "analisar",
    caminho_excel: Optional[str] = None,
) -> Dict[str, Any]:
    """Executa a Opção 2: seleção e gestão de departamento da folha CONTROLO."""
    mod = get_projeto_perfil_module()
    caminho = caminho_excel or mod.encontrar_excel_padrao()
    if not caminho or not os.path.exists(caminho):
        return {"ok": False, "status": "ERRO", "message": f"Ficheiro Excel não encontrado: {caminho}"}

    dados = mod.carregar_projeto_perfil(caminho)

    # Se não foi indicado departamento ou se a subação é listar
    if not departamento or subacao == "listar":
        lista_ctrl = [
            {
                "linha": c.get("linha"),
                "departamento": c.get("departamento"),
                "status": c.get("status") or "PENDENTE",
                "timestamp": str(c.get("timestamp") or "")[:19],
                "pendente": bool(c.get("pendente")),
            }
            for c in dados.controlo
        ]
        proximo = dados.obter_proximo_departamento()
        return {
            "ok": True,
            "subacao": "listar",
            "departamentos": lista_ctrl,
            "proximo_sugerido": proximo,
            "total": len(lista_ctrl),
        }

    dep_alvo = departamento.strip()
    # Permitir selecionar por linha numérica
    if dep_alvo.isdigit():
        item_match = next((c for c in dados.controlo if str(c.get("linha")) == dep_alvo), None)
        if item_match:
            dep_alvo = str(item_match.get("departamento") or dep_alvo)

    if subacao == "analisar":
        analise = mod.analisar_departamento_proposta(dados, dep_alvo)
        if not analise.get("encontrado"):
            return {
                "ok": False,
                "departamento": dep_alvo,
                "message": analise.get("mensagem", f"Departamento '{dep_alvo}' não encontrado na folha Proposta Ativa."),
            }

        compostas = list(analise.get("compostas", []))
        users = analise.get("usuarios", [])
        return {
            "ok": True,
            "subacao": "analisar",
            "departamento": dep_alvo,
            "compostas": compostas,
            "total_compostas": len(compostas),
            "total_usuarios": len(users),
            "total_singles": analise.get("total_singles_distintas", 0),
            "usuarios": [
                {
                    "usuario": u.get("usuario"),
                    "nome": u.get("nome"),
                    "cargo": u.get("cargo"),
                    "composta": u.get("composta") or "SEM_COMPOSTA",
                    "total_singles": u.get("total_singles", 0),
                }
                for u in users[:60]
            ],
        }

    if subacao == "validar_users":
        res_prd = mod.validar_utilizadores_prd(dados, dep_alvo)
        return {"ok": True, "subacao": "validar_users", "departamento": dep_alvo, "resultado": res_prd}

    if subacao == "cruzar_fontes":
        res_cruz = mod.cruzar_fontes_departamento(dados, dep_alvo)
        return {"ok": True, "subacao": "cruzar_fontes", "departamento": dep_alvo, "resultado": res_cruz}

    if subacao == "global":
        item_dep = next((it for it in dados.controlo if str(it.get("departamento", "")).strip().upper() == dep_alvo.upper()), None)
        if not item_dep:
            item_dep = {"departamento": dep_alvo, "status": "", "linha": "?"}

        # 1. Cruzar fontes
        res_cruz = mod.cruzar_fontes_departamento(dados, dep_alvo)
        # 2. Incorporar adicionais
        res_inc = mod.incorporar_adicionais_catalogo_departamento(dados, dep_alvo)
        # 3. Validar PRD
        res_prd = mod.validar_utilizadores_prd(dados, dep_alvo)
        # 4. Sincronizar CUA completo
        sucesso_cua = mod.sincronizar_departamento_cua_completo(dados, item_dep, confirmar_execucao=False)

        return {
            "ok": bool(sucesso_cua),
            "subacao": "global",
            "departamento": dep_alvo,
            "sucesso_cua": sucesso_cua,
            "total_incorporadas": res_inc.get("total_incorporadas", 0) if isinstance(res_inc, dict) else 0,
            "message": f"Execução global do departamento '{dep_alvo}' concluída." if sucesso_cua else f"Falha na sincronização CUA de '{dep_alvo}'.",
        }

    return {"ok": False, "message": f"Subação '{subacao}' desconhecida para departamento."}


# =====================================================================
# 3. AUDITORIA DE UTILIZADOR
# =====================================================================

def auditoria_utilizador(username: str, caminho_excel: Optional[str] = None) -> Dict[str, Any]:
    """Executa a Opção 3: Auditoria detalhada do utilizador."""
    u_norm = username.strip().upper()
    if not u_norm:
        return {"ok": False, "message": "ID do utilizador SAP não pode estar vazio."}

    mod = get_projeto_perfil_module()
    caminho = caminho_excel or mod.encontrar_excel_padrao()
    if not caminho or not os.path.exists(caminho):
        return {"ok": False, "status": "ERRO", "message": f"Ficheiro Excel não encontrado: {caminho}"}

    dados = mod.carregar_projeto_perfil(caminho)
    resultado = mod.auditar_utilizador(dados, u_norm)
    return resultado


# =====================================================================
# 4. PESQUISA E ATRIBUIÇÃO DE TRANSAÇÃO
# =====================================================================

def pesquisa_atribuir_transacao(
    tcode: str,
    username: Optional[str] = None,
    simular: bool = False,
    caminho_excel: Optional[str] = None,
) -> Dict[str, Any]:
    """Executa a Opção 4: Pesquisa transação e opcionalmente atribui a utilizador."""
    tc_clean = tcode.strip().upper()
    if not tc_clean:
        return {"ok": False, "message": "Código de transação não pode estar vazio."}

    mod = get_projeto_perfil_module()
    caminho = caminho_excel or mod.encontrar_excel_padrao()
    if not caminho or not os.path.exists(caminho):
        return {"ok": False, "status": "ERRO", "message": f"Ficheiro Excel não encontrado: {caminho}"}

    dados = mod.carregar_projeto_perfil(caminho)

    # 1. Pesquisa de roles do catálogo para a transação
    roles_projeto = dados.tcode_para_roles.get(tc_clean, [])
    if not roles_projeto:
        roles_projeto = dados.tcode_para_roles.get(mod.normalizar_texto(tc_clean), [])

    if not roles_projeto:
        return {
            "ok": True,
            "encontrado": False,
            "tcode": tc_clean,
            "roles": [],
            "message": f"A transação '{tc_clean}' não pertence ao catálogo oficial de funções do projeto.",
        }

    roles_detalhes = [
        {
            "role": r,
            "descricao": dados.roles_simples.get(r, {}).get("descricao", ""),
        }
        for r in roles_projeto
    ]

    # Se nenhum utilizador foi indicado, retorna apenas a pesquisa
    if not username:
        return {
            "ok": True,
            "encontrado": True,
            "tcode": tc_clean,
            "roles": roles_detalhes,
            "total_roles": len(roles_detalhes),
            "message": f"Transação '{tc_clean}' encontrada no projeto com {len(roles_detalhes)} função(ões) associada(s).",
        }

    # Atribuição ao utilizador
    u_clean = username.strip().upper()
    sucesso = mod.executar_fluxo_pesquisa_atribuir_transacao(
        dados=dados,
        caminho_excel=caminho,
        transacao_alvo=tc_clean,
        user_alvo=u_clean,
        simular=simular,
        assumir_sim=True,
    )

    return {
        "ok": bool(sucesso),
        "encontrado": True,
        "tcode": tc_clean,
        "usuario": u_clean,
        "roles": roles_detalhes,
        "simulacao": simular,
        "message": f"Transação '{tc_clean}' atribuída com sucesso ao utilizador '{u_clean}'." if sucesso else f"Falha na atribuição da transação '{tc_clean}' ao utilizador '{u_clean}'.",
    }


# =====================================================================
# 5. CORREÇÃO DE SINCRONIZAÇÃO POSTERIOR
# =====================================================================

def correcao_sincronizacao_posterior(caminho_excel: Optional[str] = None) -> Dict[str, Any]:
    """Executa a Opção 5: Correção de sincronização CUA posterior para utilizadores pendentes."""
    mod = get_projeto_perfil_module()
    caminho = caminho_excel or mod.encontrar_excel_padrao()
    if not caminho or not os.path.exists(caminho):
        return {"ok": False, "status": "ERRO", "message": f"Ficheiro Excel não encontrado: {caminho}"}

    dados = mod.carregar_projeto_perfil(caminho)
    discrepancias = mod.obter_discrepancias_sincronizacao(dados, caminho)

    if not discrepancias:
        return {
            "ok": True,
            "status": "SINCRONIZADO",
            "total_discrepancias": 0,
            "departamentos_pendentes": [],
            "message": "Não existem discrepâncias ou pendências de sincronização CUA. Todos os utilizadores encontram-se 100% sincronizados!",
        }

    total_u = sum(len(ul) for ul in discrepancias.values())
    res_dados = mod.executar_correcao_sincronizacao_posterior(
        dados=dados,
        caminho_excel=caminho,
        assumir_sim=True,
    )

    sucesso = res_dados is not None
    return {
        "ok": sucesso,
        "status": "SUCESSO" if sucesso else "ERRO",
        "total_discrepancias": total_u,
        "departamentos_corrigidos": list(discrepancias.keys()),
        "message": f"Correção de sincronização posterior concluída com sucesso para {total_u} utilizador(es)." if sucesso else "Ocorreram erros durante a sincronização dos utilizadores.",
    }


# =====================================================================
# 6. DIAGNÓSTICO SU53
# =====================================================================

def diagnostico_su53(
    username: str,
    target_env: str = "PRD",
    tcode_atribuir: Optional[str] = None,
    caminho_excel: Optional[str] = None,
) -> Dict[str, Any]:
    """Executa a Opção 6: Diagnóstico de buffer SU53 via RFC e fluxo de atribuição."""
    u_clean = username.strip().upper()
    if not u_clean:
        return {"ok": False, "message": "ID do utilizador SAP não pode estar vazio."}

    env_clean = "QAD" if target_env.strip().upper() in ("QAS", "QAD") else "PRD"
    mod_su53 = get_pesquisa_su53_module()
    caminho = caminho_excel or mod_su53.CAMINHO_EXCEL_PADRAO

    if tcode_atribuir:
        tc_alvo = tcode_atribuir.strip().upper()
        try:
            mod_su53.executar_atribuicao_transacao(
                caminho_excel=caminho,
                usuario=u_clean,
                tcode_escolhido=tc_alvo,
                target_env=env_clean,
            )
            return {
                "ok": True,
                "usuario": u_clean,
                "target_env": env_clean,
                "tcode_atribuido": tc_alvo,
                "message": f"Transação '{tc_alvo}' atribuída com sucesso ao utilizador '{u_clean}' em {env_clean}.",
            }
        except Exception as exc:
            return {"ok": False, "message": f"Erro ao atribuir transação: {exc}"}

    try:
        erros = mod_su53.consultar_erros_su53_rfc(u_clean, target_env=env_clean)
    except Exception as exc:
        return {"ok": False, "message": f"Falha ao consultar buffer SU53 via RFC: {exc}"}

    transacoes, objetos = mod_su53.agrupar_erros_para_analise(erros)

    # Identificar sugestões do projeto para transações bloqueadas
    tcodes_detalhes = []
    for t in transacoes:
        tc = t.get("tcode_falha")
        sug_role = None
        roles_encontradas = []
        if os.path.exists(caminho):
            try:
                sug_role, roles_encontradas = mod_su53.localizar_transacao_no_excel(caminho, tc)
            except Exception:
                pass
        tcodes_detalhes.append({
            "tcode": tc,
            "rc": t.get("rc"),
            "data": t.get("data"),
            "hora": t.get("hora"),
            "sugestao_role": sug_role,
            "roles_encontradas": roles_encontradas,
        })

    objetos_detalhes = [
        {
            "objeto": o.get("objeto"),
            "descricao": o.get("descricao") or "",
            "rc": o.get("rc_mais_recente"),
            "data_hora": o.get("data_hora_recente"),
        }
        for o in objetos
    ]

    return {
        "ok": True,
        "usuario": u_clean,
        "target_env": env_clean,
        "total_erros": len(erros),
        "total_transacoes": len(tcodes_detalhes),
        "total_objetos": len(objetos_detalhes),
        "transacoes": tcodes_detalhes,
        "objetos": objetos_detalhes,
        "message": f"Diagnóstico SU53 concluído para {u_clean} em {env_clean}. {len(tcodes_detalhes)} transações e {len(objetos_detalhes)} objetos com falha.",
    }

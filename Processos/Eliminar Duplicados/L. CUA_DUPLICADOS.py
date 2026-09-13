# -*- coding: utf-8 -*-
"""
L. CUA_DUPLICADOS.py
======================================================================
FASE 1 (somente leitura, via RFC): analise de atribuicoes duplicadas de
funcoes SAP (AGR_USERS) para os utilizadores dos departamentos
disponiveis/pendentes na sheet CONTROLO do ficheiro Excel oficial.

Nao altera o SAP. Apenas le (RFC_READ_TABLE) e imprime um relatorio de
duplicados no terminal, separando atribuicoes diretas de herdadas
(COL_FLAG) e sinalizando casos ambiguos para revisao humana.
======================================================================
"""

import os
import sys
import importlib.util
from pathlib import Path
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

# Auto-deteção e re-execução no ambiente virtual .venv-rfc se pandas não estiver presente
try:
    import pandas  # noqa: F401
except ImportError:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    venv_python = os.path.join(base_dir, "..", "..", ".venv-rfc", "Scripts", "python.exe")
    venv_python = os.path.abspath(venv_python)
    if os.path.exists(venv_python) and sys.executable.lower() != venv_python.lower():
        import subprocess
        res = subprocess.run([venv_python, os.path.abspath(__file__)] + sys.argv[1:])
        sys.exit(res.returncode)

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _carregar_modulo_projeto_perfil():
    caminho = PROJECT_ROOT / "Projeto Perfil.py"
    spec = importlib.util.spec_from_file_location("projeto_perfil", caminho)
    modulo = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(modulo)
    return modulo


# =====================================================================
# ESCOPO: departamentos (CONTROLO) -> utilizadores (Proposta Ativa)
# =====================================================================

def obter_utilizadores_escopo(caminho_excel: Optional[str] = None) -> Dict[str, Any]:
    projeto_perfil = _carregar_modulo_projeto_perfil()
    dados = projeto_perfil.carregar_projeto_perfil(caminho_excel)

    departamentos = [c["departamento"] for c in dados.controlo]

    utilizadores_por_departamento: Dict[str, List[str]] = {}
    utilizadores: set = set()

    for dep in departamentos:
        resultado = projeto_perfil.analisar_departamento_proposta(dados, dep)
        if not resultado.get("encontrado"):
            utilizadores_por_departamento[dep] = []
            continue
        lista_dep = []
        for u in resultado["usuarios"]:
            user_norm = str(u.get("usuario", "")).strip().upper()
            if not user_norm or user_norm.lower() in ("nan", "none"):
                continue
            lista_dep.append(user_norm)
            utilizadores.add(user_norm)
        utilizadores_por_departamento[dep] = sorted(set(lista_dep))

    return {
        "ficheiro": dados.caminho,
        "departamentos": departamentos,
        "utilizadores_por_departamento": utilizadores_por_departamento,
        "utilizadores": sorted(utilizadores),
    }


# =====================================================================
# LEITURA DE AGR_USERS VIA RFC (SOMENTE LEITURA)
# =====================================================================

def ler_agr_users_prd(utilizadores: List[str]) -> Dict[str, Any]:
    if not utilizadores:
        return {"ok": True, "registos": []}

    os.environ["SAP_TARGET_ENV"] = "PRD"
    try:
        from sap_rfc._rfc_common import (
            build_connection_params_for, load_project_env, find_project_root,
            make_read_only_guard, read_table, make_option_in
        )
        from pyrfc import Connection
    except Exception as err:
        return {"ok": False, "erro": f"Falha ao carregar infraestrutura RFC: {err}", "registos": []}

    try:
        project_root = find_project_root()
        load_project_env(project_root)
        params = build_connection_params_for("PRD")
        conn = Connection(**params)
        guard = make_read_only_guard(("AGR_USERS",))

        registos = []
        chunk_size = 25
        for i in range(0, len(utilizadores), chunk_size):
            chunk = utilizadores[i:i + chunk_size]
            opts = make_option_in("UNAME", chunk)
            rows = read_table(
                conn, guard,
                table_name="AGR_USERS",
                fields=["UNAME", "AGR_NAME", "FROM_DAT", "TO_DAT", "COL_FLAG"],
                options=opts,
                rowcount=0,
            )
            for uname, agr_name, from_dat, to_dat, col_flag in rows:
                registos.append({
                    "uname": uname.strip().upper(),
                    "agr_name": agr_name.strip().upper(),
                    "from_dat_raw": from_dat.strip(),
                    "to_dat_raw": to_dat.strip(),
                    "col_flag": col_flag.strip(),
                })

        conn.close()
        return {"ok": True, "registos": registos, "sistema": "PRD", "client": params.get("client")}
    except Exception as err:
        return {"ok": False, "erro": str(err), "registos": []}


# =====================================================================
# ANALISE DE DUPLICADOS
# =====================================================================

def _parse_data(raw: str) -> Optional[str]:
    """Devolve a data em formato AAAAMMDD (comparavel lexicograficamente) ou None se vazia/invalida."""
    s = str(raw or "").strip()
    if len(s) == 8 and s.isdigit() and s != "00000000":
        return s
    return None


def analisar_duplicados(registos: List[Dict[str, Any]]) -> Dict[str, Any]:
    for r in registos:
        r["from_dat"] = _parse_data(r["from_dat_raw"])
        r["to_dat"] = _parse_data(r["to_dat_raw"])
        r["direta"] = not bool(r["col_flag"])

    grupos: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
    for r in registos:
        grupos[(r["uname"], r["agr_name"])].append(r)

    duplicados = {chave: lista for chave, lista in grupos.items() if len(lista) > 1}

    resultado_grupos = []

    for (uname, agr_name), lista in sorted(duplicados.items()):
        diretas = [r for r in lista if r["direta"]]
        herdadas = [r for r in lista if not r["direta"]]

        classificacoes: List[Tuple[Dict[str, Any], str, str]] = []
        ambiguo = False
        motivo_ambiguo = ""

        if not diretas:
            # Todas as ocorrencias sao herdadas de compostas: nao ha nada a remover na CUA.
            for r in herdadas:
                classificacoes.append((r, "IGNORAR", "HERDADA"))
            tipo_grupo = "SO_HERDADAS"

        elif not herdadas:
            # Todas diretas: aplicar a regra de desempate TO_DAT desc, FROM_DAT desc.
            invalidas = [r for r in diretas if r["to_dat"] is None or r["from_dat"] is None]
            if invalidas:
                ambiguo = True
                motivo_ambiguo = "FROM_DAT ou TO_DAT vazio/invalido numa ou mais ocorrencias diretas."
                for r in diretas:
                    classificacoes.append((r, "AMBIGUO", "DIRETA"))
                tipo_grupo = "DIRETAS_AMBIGUAS"
            else:
                ordenados = sorted(diretas, key=lambda r: (r["to_dat"], r["from_dat"]), reverse=True)
                topo = ordenados[0]
                segundo = ordenados[1]
                if topo["to_dat"] == segundo["to_dat"] and topo["from_dat"] == segundo["from_dat"]:
                    ambiguo = True
                    motivo_ambiguo = (
                        "Duas ou mais ocorrencias diretas com FROM_DAT e TO_DAT identicos "
                        "(a interface nao permite distinguir qual linha do grid sera removida)."
                    )
                    for r in diretas:
                        classificacoes.append((r, "AMBIGUO", "DIRETA"))
                    tipo_grupo = "DIRETAS_AMBIGUAS"
                else:
                    classificacoes.append((topo, "MANTER", "DIRETA"))
                    for r in ordenados[1:]:
                        classificacoes.append((r, "REMOVER", "DIRETA"))
                    tipo_grupo = "DIRETAS_RESOLVIDAS"

        else:
            # Mistura de diretas e herdadas para a mesma UNAME + AGR_NAME.
            if len(diretas) == 1:
                classificacoes.append((diretas[0], "MANTER", "DIRETA"))
                for r in herdadas:
                    classificacoes.append((r, "IGNORAR", "HERDADA"))
                tipo_grupo = "MISTA_UNICA_DIRETA"
            else:
                ambiguo = True
                motivo_ambiguo = (
                    "Duplicados diretos coexistem com atribuicoes herdadas de composta "
                    "para a mesma funcao; requer revisao humana separada."
                )
                for r in diretas:
                    classificacoes.append((r, "AMBIGUO", "DIRETA"))
                for r in herdadas:
                    classificacoes.append((r, "IGNORAR", "HERDADA"))
                tipo_grupo = "MISTA_MULTIPLAS_DIRETAS"

        resultado_grupos.append({
            "uname": uname,
            "agr_name": agr_name,
            "tipo_grupo": tipo_grupo,
            "ambiguo": ambiguo,
            "motivo_ambiguo": motivo_ambiguo,
            "classificacoes": classificacoes,
        })

    return {
        "total_registos": len(registos),
        "total_grupos_duplicados": len(duplicados),
        "grupos": resultado_grupos,
    }


# =====================================================================
# APRESENTACAO NO TERMINAL
# =====================================================================

def imprimir_relatorio(analise: Dict[str, Any], total_users: int) -> None:
    grupos = analise["grupos"]
    users_com_duplicados = sorted({g["uname"] for g in grupos})

    print(f"\nDUPLICADOS ENCONTRADOS: {len(grupos)} | USERS: {len(users_com_duplicados)}")

    manter = remover = ignorar = ambiguo_linhas = 0
    grupos_ambiguos = []

    for g in grupos:
        print(f"{g['uname']} | {g['agr_name']}")
        for reg, acao, tipo in g["classificacoes"]:
            from_s = reg["from_dat_raw"] or "(vazio)"
            to_s = reg["to_dat_raw"] or "(vazio)"
            print(f"  {acao:<8}| FROM={from_s} | TO={to_s} | {tipo}")
            if acao == "MANTER":
                manter += 1
            elif acao == "REMOVER":
                remover += 1
            elif acao == "IGNORAR":
                ignorar += 1
            elif acao == "AMBIGUO":
                ambiguo_linhas += 1
        if g["ambiguo"]:
            grupos_ambiguos.append(g)

    print("\n" + "-" * 75)
    print("RESUMO")
    print("-" * 75)
    print(f"Users analisados: {total_users}")
    print(f"Atribuições lidas: {analise['total_registos']}")
    print(f"Grupos duplicados: {analise['total_grupos_duplicados']}")
    print(f"Ocorrências a manter: {manter}")
    print(f"Ocorrências diretas candidatas a remover: {remover}")
    print(f"Ocorrências herdadas ignoradas: {ignorar}")
    print(f"Casos ambíguos: {len(grupos_ambiguos)}")

    if grupos_ambiguos:
        print("\n" + "-" * 75)
        print("CASOS AMBÍGUOS (revisão humana necessária, nada será alterado)")
        print("-" * 75)
        for g in grupos_ambiguos:
            print(f"{g['uname']} | {g['agr_name']} -> {g['motivo_ambiguo']}")
            for reg, acao, tipo in g["classificacoes"]:
                from_s = reg["from_dat_raw"] or "(vazio)"
                to_s = reg["to_dat_raw"] or "(vazio)"
                print(f"    FROM={from_s} | TO={to_s} | {tipo}")


def escolher_candidato_teste(analise: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Escolhe o primeiro grupo DIRETAS_RESOLVIDAS (nao ambiguo) como candidato ao primeiro teste real."""
    for g in analise["grupos"]:
        if g["tipo_grupo"] != "DIRETAS_RESOLVIDAS" or g["ambiguo"]:
            continue
        manter = next(r for r, acao, _ in g["classificacoes"] if acao == "MANTER")
        remover = [r for r, acao, _ in g["classificacoes"] if acao == "REMOVER"]
        if len(remover) == 1:
            return {
                "uname": g["uname"],
                "agr_name": g["agr_name"],
                "manter": manter,
                "remover": remover[0],
            }
    return None


# =====================================================================
# EXECUÇÃO PRINCIPAL (FASE 1 - SOMENTE ANÁLISE)
# =====================================================================

def executar_analise(caminho_excel: Optional[str] = None) -> Dict[str, Any]:
    print("=" * 75)
    print("  FASE 1: ANALISE (RFC, SOMENTE LEITURA) DE DUPLICADOS EM AGR_USERS")
    print("=" * 75)

    escopo = obter_utilizadores_escopo(caminho_excel)
    print(f"\nFicheiro: {escopo['ficheiro']}")
    print("Departamentos (sheet CONTROLO):")
    for dep in escopo["departamentos"]:
        qtd = len(escopo["utilizadores_por_departamento"].get(dep, []))
        print(f"  - {dep}: {qtd} utilizador(es)")
    print(f"\nTotal de utilizadores únicos no escopo: {len(escopo['utilizadores'])}")

    leitura = ler_agr_users_prd(escopo["utilizadores"])
    if not leitura.get("ok"):
        print(f"\n[ERRO] Falha na leitura RFC de AGR_USERS: {leitura.get('erro')}")
        return {"ok": False, "erro": leitura.get("erro")}

    analise = analisar_duplicados(leitura["registos"])
    imprimir_relatorio(analise, len(escopo["utilizadores"]))

    candidato = escolher_candidato_teste(analise)
    print("\n" + "=" * 75)
    if candidato:
        print("  CANDIDATO PARA O PRIMEIRO TESTE REAL (nenhuma alteração foi feita)")
        print("=" * 75)
        print(f"Utilizador: {candidato['uname']}")
        print(f"Função:     {candidato['agr_name']}")
        print(f"Manter:     FROM={candidato['manter']['from_dat_raw']} | TO={candidato['manter']['to_dat_raw']}")
        print(f"Remover:    FROM={candidato['remover']['from_dat_raw']} | TO={candidato['remover']['to_dat_raw']}")
        print("Motivo: duplicado direto simples e não ambíguo (TO_DAT mais distante mantido).")
    else:
        print("  Nenhum duplicado direto simples e não ambíguo foi encontrado para propor um teste real.")

    return {"ok": True, "escopo": escopo, "analise": analise, "candidato": candidato}


if __name__ == "__main__":
    executar_analise()

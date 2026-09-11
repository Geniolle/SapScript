from __future__ import annotations

import importlib.util
import fnmatch
import json
import os
import sys
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ["SAP_TARGET_ENV"] = "PRD"

spec = importlib.util.spec_from_file_location("projeto_perfil", ROOT / "Projeto Perfil.py")
if spec is None or spec.loader is None:
    raise RuntimeError("Não foi possível carregar Projeto Perfil.py")
modulo = importlib.util.module_from_spec(spec)
spec.loader.exec_module(modulo)

from pyrfc import Connection  # type: ignore  # noqa: E402
from sap_rfc._rfc_common import (  # noqa: E402
    build_connection_params_for,
    find_project_root,
    load_project_env,
    make_option_in,
    make_read_only_guard,
    read_table,
)


def assignment_status(from_dat: str, to_dat: str) -> str:
    hoje = int(date.today().strftime("%Y%m%d"))
    inicio = int(from_dat) if str(from_dat).strip().isdigit() else 0
    fim = int(to_dat) if str(to_dat).strip().isdigit() else 99991231
    if hoje < inicio:
        return "FUTURO"
    if hoje > fim:
        return "EXPIRADO"
    return "ATIVO"


def main() -> None:
    dados = modulo.carregar_projeto_perfil()
    import pandas as pd

    sheet_exclusao = next(
        (s for s in dados.sheets_disponiveis if modulo.normalizar_nome_coluna(s) == "EXCLUCAO"),
        None,
    )
    padroes_exclusao: list[str] = []
    if sheet_exclusao:
        df_exclusao = pd.read_excel(
            modulo.abrir_excel_seguro(dados.caminho), sheet_name=sheet_exclusao
        )
        if len(df_exclusao.columns):
            padroes_exclusao = [
                str(v).strip().upper()
                for v in df_exclusao.iloc[:, 0].dropna().tolist()
                if str(v).strip()
            ]
    sheet_authority = next(
        (s for s in dados.sheets_disponiveis if modulo.normalizar_nome_coluna(s) == "PFCGAUTHORITY"),
        None,
    )
    authority_por_composta: dict[str, set[str]] = {}
    if sheet_authority:
        df_authority = pd.read_excel(
            modulo.abrir_excel_seguro(dados.caminho), sheet_name=sheet_authority
        )
        colunas = {modulo.normalizar_nome_coluna(c): c for c in df_authority.columns}
        col_role = colunas.get("AGRNAME")
        col_composta = colunas.get("AGRNAMECOMPOSTA")
        if col_role and col_composta:
            for _, row in df_authority.iterrows():
                composta = str(row.get(col_composta, "")).strip().upper()
                role = str(row.get(col_role, "")).strip().upper()
                if composta and role and composta not in ("NAN", "NONE") and role not in ("NAN", "NONE"):
                    authority_por_composta.setdefault(composta, set()).add(role)

    def excluida(role: str) -> bool:
        return any(fnmatch.fnmatchcase(role, padrao) for padrao in padroes_exclusao)

    def expandir_funcoes(roles_iniciais: set[str]) -> set[str]:
        expandidas: set[str] = set()
        pendentes = list(roles_iniciais)
        while pendentes:
            role = pendentes.pop()
            if not role or role in expandidas or excluida(role):
                continue
            expandidas.add(role)
            membros = set(dados.roles_compostas.get(role, {}).get("roles_filhas", []))
            membros.update(authority_por_composta.get(role, set()))
            pendentes.extend(membros - expandidas)
        return expandidas
    departamentos = [c["departamento"] for c in dados.controlo if c.get("pendente")]
    esperado_por_user: dict[str, set[str]] = {}
    detalhe_por_user: dict[str, dict] = {}

    for departamento in departamentos:
        analise = modulo.analisar_departamento_proposta(dados, departamento)
        if not analise.get("encontrado"):
            continue
        for item in analise["usuarios"]:
            user = str(item.get("usuario", "")).strip().upper()
            if not user:
                continue
            roles = {str(r).strip().upper() for r in item.get("singles", []) if str(r).strip()}
            if item.get("composta"):
                roles.add(str(item["composta"]).strip().upper())
            esperado_por_user.setdefault(user, set()).update(expandir_funcoes(roles))
            detalhe_por_user[user] = {
                "nome": item.get("nome"),
                "cargo": item.get("cargo"),
                "departamento": departamento,
            }

    load_project_env(find_project_root())
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(["AGR_USERS"])
    ativas = {u: set() for u in esperado_por_user}
    nao_ativas = {u: {} for u in esperado_por_user}

    conn = Connection(**params)
    try:
        users = sorted(esperado_por_user)
        for inicio in range(0, len(users), 20):
            lote = users[inicio : inicio + 20]
            rows = read_table(
                conn,
                guard,
                table_name="AGR_USERS",
                fields=["AGR_NAME", "UNAME", "FROM_DAT", "TO_DAT"],
                options=make_option_in("UNAME", lote),
                rowcount=0,
            )
            for role, user, from_dat, to_dat in rows:
                role = role.strip().upper()
                user = user.strip().upper()
                if user not in ativas or not role:
                    continue
                status = assignment_status(from_dat, to_dat)
                if status == "ATIVO":
                    ativas[user].add(role)
                else:
                    nao_ativas[user][role] = {
                        "status": status,
                        "inicio": str(from_dat).strip(),
                        "fim": str(to_dat).strip(),
                    }
    finally:
        conn.close()

    utilizadores = []
    for user in sorted(esperado_por_user):
        faltam = sorted(esperado_por_user[user] - ativas[user])
        adicionais_brutos = sorted(ativas[user] - esperado_por_user[user])
        excluidas = sorted(
            role for role in adicionais_brutos
            if excluida(role)
        )
        adicionais = sorted(set(adicionais_brutos) - set(excluidas))
        inativas_relevantes = {
            role: nao_ativas[user][role]
            for role in faltam
            if role in nao_ativas[user]
        }
        utilizadores.append(
            {
                "utilizador": user,
                **detalhe_por_user[user],
                "funcoes_esperadas": len(esperado_por_user[user]),
                "funcoes_esperadas_ativas": len(esperado_por_user[user] & ativas[user]),
                "conforme": not faltam,
                "funcoes_em_falta": faltam,
                "funcoes_adicionais": adicionais,
                "funcoes_desconsideradas_por_exclusao": excluidas,
                "atribuicoes_nao_ativas": inativas_relevantes,
            }
        )

    divergentes = [u for u in utilizadores if not u["conforme"]]
    resultado = {
        "ficheiro": dados.caminho,
        "sistema": "PRD",
        "mandante": params.get("client"),
        "data_validacao": date.today().isoformat(),
        "departamentos_em_aberto": departamentos,
        "utilizadores_verificados": len(utilizadores),
        "utilizadores_conformes": len(utilizadores) - len(divergentes),
        "utilizadores_divergentes": len(divergentes),
        "atribuicoes_esperadas": sum(len(v) for v in esperado_por_user.values()),
        "atribuicoes_esperadas_ativas": sum(len(esperado_por_user[u] & ativas[u]) for u in esperado_por_user),
        "padroes_exclusao": padroes_exclusao,
        "roles_pfcg_authority": sorted(set().union(*authority_por_composta.values()) if authority_por_composta else set()),
        "funcoes_adicionais_ativas": sum(len(u["funcoes_adicionais"]) for u in utilizadores),
        "funcoes_desconsideradas_por_exclusao": sum(
            len(u["funcoes_desconsideradas_por_exclusao"]) for u in utilizadores
        ),
        "utilizadores": utilizadores,
    }
    print(json.dumps(resultado, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

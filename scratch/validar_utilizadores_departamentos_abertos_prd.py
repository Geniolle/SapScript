from __future__ import annotations

import importlib.util
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
            esperado_por_user.setdefault(user, set()).update(roles)
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
        "utilizadores": utilizadores,
    }
    print(json.dumps(resultado, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

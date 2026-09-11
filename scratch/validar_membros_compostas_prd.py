from __future__ import annotations

import importlib.util
import json
import os
import sys
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


def main() -> None:
    dados = modulo.carregar_projeto_perfil()
    esperados = {
        role: {str(r).strip().upper() for r in info.get("roles_filhas", []) if str(r).strip()}
        for role, info in dados.roles_compostas.items()
    }

    load_project_env(find_project_root())
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(["AGR_AGRS"])
    atuais = {role: set() for role in esperados}

    conn = Connection(**params)
    try:
        roles = sorted(esperados)
        for inicio in range(0, len(roles), 20):
            lote = roles[inicio : inicio + 20]
            rows = read_table(
                conn,
                guard,
                table_name="AGR_AGRS",
                fields=["AGR_NAME", "CHILD_AGR"],
                options=make_option_in("AGR_NAME", lote),
                rowcount=0,
            )
            for composta, filha in rows:
                composta = composta.strip().upper()
                filha = filha.strip().upper()
                if composta in atuais and filha:
                    atuais[composta].add(filha)
    finally:
        conn.close()

    divergencias = []
    for role in sorted(esperados):
        faltam = sorted(esperados[role] - atuais[role])
        adicionais = sorted(atuais[role] - esperados[role])
        if faltam or adicionais:
            divergencias.append(
                {
                    "role_composta": role,
                    "esperadas": len(esperados[role]),
                    "prd": len(atuais[role]),
                    "faltam_no_prd": faltam,
                    "adicionais_no_prd": adicionais,
                }
            )

    resultado = {
        "ficheiro": dados.caminho,
        "sistema": "PRD",
        "mandante": params.get("client"),
        "roles_compostas": len(esperados),
        "atribuicoes_esperadas_total": sum(len(v) for v in esperados.values()),
        "atribuicoes_prd_total": sum(len(v) for v in atuais.values()),
        "compostas_conformes": len(esperados) - len(divergencias),
        "compostas_divergentes": len(divergencias),
        "divergencias": divergencias,
    }
    print(json.dumps(resultado, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

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

from sap_rfc._rfc_common import (  # noqa: E402
    build_connection_params_for,
    find_project_root,
    load_project_env,
    make_option_in,
    make_read_only_guard,
    read_table,
)
from pyrfc import Connection  # type: ignore  # noqa: E402


def main() -> None:
    dados = modulo.carregar_projeto_perfil()
    esperadas = {
        role: {str(t).strip().upper() for t in info.get("tcodes", []) if str(t).strip()}
        for role, info in dados.roles_simples.items()
    }

    project_root = find_project_root()
    load_project_env(project_root)
    params = build_connection_params_for("PRD")
    guard = make_read_only_guard(["AGR_TCODES"])
    atuais = {role: set() for role in esperadas}

    conn = Connection(**params)
    try:
        roles = sorted(esperadas)
        for inicio in range(0, len(roles), 20):
            lote = roles[inicio : inicio + 20]
            rows = read_table(
                conn,
                guard,
                table_name="AGR_TCODES",
                fields=["AGR_NAME", "TCODE"],
                options=make_option_in("AGR_NAME", lote),
                rowcount=0,
            )
            for role, tcode in rows:
                role = role.strip().upper()
                tcode = tcode.strip().upper()
                if role in atuais and tcode:
                    atuais[role].add(tcode)
    finally:
        conn.close()

    divergencias = []
    for role in sorted(esperadas):
        faltam = sorted(esperadas[role] - atuais[role])
        adicionais = sorted(atuais[role] - esperadas[role])
        if faltam or adicionais:
            divergencias.append(
                {
                    "role": role,
                    "esperadas": len(esperadas[role]),
                    "prd": len(atuais[role]),
                    "faltam_no_prd": faltam,
                    "adicionais_no_prd": adicionais,
                }
            )

    resultado = {
        "ficheiro": dados.caminho,
        "sistema": "PRD",
        "mandante": params.get("client"),
        "roles_simples": len(esperadas),
        "roles_sem_tcode_no_excel": sum(not tcodes for tcodes in esperadas.values()),
        "tcodes_esperadas_total": sum(len(v) for v in esperadas.values()),
        "tcodes_prd_total": sum(len(v) for v in atuais.values()),
        "roles_conformes": len(esperadas) - len(divergencias),
        "roles_divergentes": len(divergencias),
        "divergencias": divergencias,
    }
    print(json.dumps(resultado, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

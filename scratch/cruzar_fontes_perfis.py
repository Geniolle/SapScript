from __future__ import annotations

import fnmatch
import importlib.util
import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
spec = importlib.util.spec_from_file_location("projeto_perfil", ROOT / "Projeto Perfil.py")
if spec is None or spec.loader is None:
    raise RuntimeError("Não foi possível carregar Projeto Perfil.py")
modulo = importlib.util.module_from_spec(spec)
spec.loader.exec_module(modulo)

import pandas as pd  # noqa: E402


def main() -> None:
    dados = modulo.carregar_projeto_perfil()
    excel = pd.ExcelFile(modulo.abrir_excel_seguro(dados.caminho))

    sheet_auth = next(s for s in excel.sheet_names if modulo.normalizar_nome_coluna(s) == "PFCGAUTHORITY")
    df_auth = pd.read_excel(excel, sheet_name=sheet_auth, dtype=object)
    auth_por_composta: dict[str, set[str]] = {}
    for _, row in df_auth.iterrows():
        composta = modulo.normalizar_texto(row.get("AGR_NAME_COMPOSTA"))
        role = modulo.normalizar_texto(row.get("AGR_NAME"))
        if composta and role:
            auth_por_composta.setdefault(composta, set()).add(role)

    sheet_exc = next(s for s in excel.sheet_names if modulo.normalizar_nome_coluna(s) == "EXCLUCAO")
    df_exc = pd.read_excel(excel, sheet_name=sheet_exc, dtype=object)
    padroes = [modulo.normalizar_texto(v) for v in df_exc.iloc[:, 0].dropna() if modulo.normalizar_texto(v)]

    def excluida(role: str) -> bool:
        return any(fnmatch.fnmatchcase(role, p) for p in padroes)

    departamentos = [c["departamento"] for c in dados.controlo if c.get("pendente")]
    resultado_users = []
    for departamento in departamentos:
        analise = modulo.analisar_departamento_proposta(dados, departamento)
        for user in analise.get("usuarios", []):
            composta = modulo.normalizar_texto(user.get("composta"))
            proposta = {modulo.normalizar_texto(r) for r in user.get("singles", []) if modulo.normalizar_texto(r)}
            membros = set(dados.roles_compostas.get(composta, {}).get("roles_filhas", []))
            authority = auth_por_composta.get(composta, set())
            esperado_cruzado = ({composta} if composta else set()) | membros | authority
            esperado_cruzado = {r for r in esperado_cruzado if not excluida(r)}
            proposta_considerada = {r for r in proposta if not excluida(r)}
            resultado_users.append({
                "utilizador": modulo.normalizar_texto(user.get("usuario")),
                "composta": composta,
                "proposta_singles": len(proposta_considerada),
                "membros_pfcg_composta": len(membros),
                "avulsas_authority_da_composta": sorted(authority),
                "esperado_cruzado_total": len(esperado_cruzado),
                "na_composta_ou_authority_e_fora_da_proposta": sorted((membros | authority) - proposta_considerada),
                "na_proposta_e_fora_da_composta_authority": sorted(proposta_considerada - (membros | authority)),
                "proposta_fora_pfcg_create_authority": sorted(
                    r for r in proposta_considerada
                    if r not in dados.roles_simples and r not in authority
                ),
                "esperado_cruzado": sorted(esperado_cruzado),
            })

    print(json.dumps({
        "ficheiro": dados.caminho,
        "padroes_exclusao": padroes,
        "roles_pfcg_create": len(dados.roles_simples),
        "roles_pfcg_composta": len(dados.roles_compostas),
        "users": resultado_users,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

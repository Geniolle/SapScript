# -*- coding: utf-8 -*-
"""
C. SYNC_LOTE.py
======================================================================
FASE 3: aplica `sincronizar_utilizador` (de B. SYNC_EXECUTAR.py) a todos
os utilizadores dos departamentos da sheet CONTROLO (ficheiro Excel
oficial), um de cada vez, na mesma sessão SAP GUI "SPA".

Para na primeira anomalia (ok=False) e mostra um resumo do que já foi
processado com sucesso e do utilizador que falhou.
======================================================================
"""

import sys
import time
import importlib.util
from pathlib import Path
from typing import Any, Dict, List

FOLDER = Path(__file__).resolve().parent
PROJECT_ROOT = FOLDER.parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _importar_modulo(caminho: Path, nome_modulo: str):
    spec = importlib.util.spec_from_file_location(nome_modulo, caminho)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def obter_utilizadores_controlo() -> Dict[str, Any]:
    caminho = PROJECT_ROOT / "Processos" / "Eliminar Duplicados" / "L. CUA_DUPLICADOS.py"
    lcua = _importar_modulo(caminho, "l_cua_duplicados")
    return lcua.obter_utilizadores_escopo()


def executar_lote(excluir: List[str] = None) -> Dict[str, Any]:
    excluir = {u.strip().upper() for u in (excluir or [])}

    escopo = obter_utilizadores_controlo()
    utilizadores = [u for u in escopo["utilizadores"] if u not in excluir]

    print("=" * 75)
    print("  SINCRONIZAR AMBIENTES EM LOTE: PRD -> QAS (via CUA)")
    print("=" * 75)
    print(f"\nDepartamentos (CONTROLO): {escopo['departamentos']}")
    for dep, users in escopo["utilizadores_por_departamento"].items():
        print(f"  {dep}: {len(users)} utilizador(es) -> {users}")
    print(f"\nTotal a processar: {len(utilizadores)} (excluídos já feitos: {sorted(excluir)})")

    sync_mod = _importar_modulo(FOLDER / "B. SYNC_EXECUTAR.py", "b_sync_executar")

    resultados: List[Dict[str, Any]] = []
    for i, uname in enumerate(utilizadores, start=1):
        print("\n" + "#" * 75)
        print(f"# [{i}/{len(utilizadores)}] {uname}")
        print("#" * 75)
        try:
            r = sync_mod.sincronizar_utilizador(uname)
        except Exception as err:
            r = {"ok": False, "uname": uname, "message": f"Exceção não tratada: {err}"}
            print(f"[ERRO] {uname}: {err}")

        resultados.append(r)
        if not r.get("ok"):
            print(f"\n[LOTE ABORTADO] Falha em '{uname}'. A parar o processamento em lote.")
            break

        time.sleep(0.5)

    print("\n" + "=" * 75)
    print("RESUMO DO LOTE")
    print("=" * 75)
    sucesso = [r for r in resultados if r.get("ok")]
    falha = [r for r in resultados if not r.get("ok")]
    print(f"Processados: {len(resultados)} / {len(utilizadores)}")
    print(f"Sucesso: {len(sucesso)} -> {[r.get('uname') for r in sucesso]}")
    print(f"Falha: {len(falha)} -> {[r.get('uname') for r in falha]}")
    restantes = utilizadores[len(resultados):]
    if restantes:
        print(f"Não processados (lote parado antes): {restantes}")

    return {
        "ok": len(falha) == 0,
        "resultados": resultados,
        "sucesso": [r.get("uname") for r in sucesso],
        "falha": [r.get("uname") for r in falha],
        "nao_processados": restantes,
    }


if __name__ == "__main__":
    executar_lote(excluir=["S4244"])

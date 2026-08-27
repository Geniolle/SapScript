from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path


def find_project_root() -> Path:
    current = Path(__file__).resolve().parent
    for candidate in [current, *current.parents]:
        if (candidate / ".env.example").exists():
            return candidate
    raise RuntimeError("Não foi possível localizar a raiz do projeto a partir de tools/.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Teste direto da task fixa pfcg_role_analysis sem frontend."
    )
    parser.add_argument("--role-name", help="Nome exato da função/perfil PFCG.")
    return parser.parse_args()


def prompt_role_name(role_name: str | None) -> str:
    if role_name:
        return role_name
    return input("Nome da função/perfil PFCG: ")


def main() -> int:
    args = parse_args()
    project_root = find_project_root()
    os.environ.setdefault("SAP_SCRIPT_PROJECT_DIR", str(project_root))

    worker_dir = project_root / "sap_script_web_cockpit_v2" / "worker"
    if str(worker_dir) not in sys.path:
        sys.path.insert(0, str(worker_dir))

    from sap_tasks import run_sap_task  # type: ignore

    job = {
        "id": "manual-pfcg-role-analysis",
        "task": "pfcg_role_analysis",
        "params": {"role_name": prompt_role_name(args.role_name)},
    }
    try:
        status, log = run_sap_task(job)
    except Exception as exc:
        print("=" * 60)
        print("ERRO")
        print("=" * 60)
        print(str(exc) or exc.__class__.__name__)
        return 1

    print("=" * 60)
    print("STATUS JSON")
    print("=" * 60)
    try:
        payload = json.loads(status)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    except Exception:
        print(status)

    print()
    print("=" * 60)
    print("LOG")
    print("=" * 60)
    print(log)
    return 0


if __name__ == "__main__":
    sys.exit(main())

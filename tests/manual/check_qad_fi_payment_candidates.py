from __future__ import annotations

import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params_for_env, find_project_root, load_project_env
from sap_rfc.fi_document_service import _bridge_python_executable

from tests.manual.scan_qad_z_programs_for_bapi_document_post import _collect_hits, asdict

try:
    from pyrfc import Connection  # type: ignore
except Exception as exc:  # pragma: no cover - runtime guard
    Connection = None  # type: ignore[assignment]
    PYRFC_IMPORT_ERROR = exc
else:
    PYRFC_IMPORT_ERROR = None


CANDIDATES = [
    "ZFI_EXTERNAL_PAYMENTS",
    "ZFI_UP_EXT_PAYMENT",
    "ZFI_UP_EXT_PAYMENT_SRC",
    "ZFI_FI_INST_UPD_ACC_POST",
    "ZFI_DOCUMENT_UPDATE",
]


def main() -> int:
    if Connection is None:
        bridge_python = _bridge_python_executable()
        if bridge_python is None:
            raise RuntimeError(f"PyRFC indisponível: {PYRFC_IMPORT_ERROR}")
        import subprocess

        proc = subprocess.run(
            [str(bridge_python), str(Path(__file__).resolve())],
            cwd=str(REPO_ROOT),
            capture_output=True,
            text=True,
        )
        if proc.stdout:
            print(proc.stdout, end="")
        if proc.stderr:
            print(proc.stderr, end="", file=sys.stderr)
        return proc.returncode

    project_root = find_project_root()
    load_project_env(project_root)
    connection = Connection(**build_connection_params_for_env("QAD"))  # type: ignore[misc]
    try:
        hits = []
        for program in CANDIDATES:
            hits.extend(_collect_hits(connection, program))
        print(
            json.dumps(
                {
                    "candidates": CANDIDATES,
                    "hit_count": len(hits),
                    "hits": [asdict(hit) for hit in hits],
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    finally:
        try:
            connection.close()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())

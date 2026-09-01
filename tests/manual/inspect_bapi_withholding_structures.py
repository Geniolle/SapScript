import json
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc.fi_document_service import _bridge_python_executable


def main() -> int:
    bridge_python = _bridge_python_executable()
    if bridge_python is None:
        raise RuntimeError("Bridge Python indisponível.")

    bridge_script = f"""
from pathlib import Path
import json
import sys

repo_root = Path(r'{REPO_ROOT}')
sys.path.insert(0, str(repo_root))

from sap_rfc._rfc_common import build_connection_params_for_env, load_project_env
from pyrfc import Connection

load_project_env(repo_root)
connection = Connection(**build_connection_params_for_env('QAD'))
try:
    def field_names(table_name):
        result = connection.call('DDIF_FIELDINFO_GET', TABNAME=table_name, LANGU='E')
        rows = result.get('DFIES_TAB') or []
        return [str(row.get('FIELDNAME', '')).strip() for row in rows if isinstance(row, dict) and row.get('FIELDNAME')]

    payload = {{
        'BAPIACWT09': field_names('BAPIACWT09'),
        'BAPIACAP09': field_names('BAPIACAP09'),
    }}
    print(json.dumps(payload, ensure_ascii=False, indent=2))
finally:
    connection.close()
"""
    proc = subprocess.run([str(bridge_python), "-c", bridge_script], cwd=str(REPO_ROOT), capture_output=True, text=True)
    if proc.returncode != 0:
        print(proc.stdout, end="")
        print(proc.stderr, end="", file=sys.stderr)
        return proc.returncode
    print(proc.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

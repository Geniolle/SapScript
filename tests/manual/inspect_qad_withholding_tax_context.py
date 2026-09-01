import json
import os
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import load_project_env
from sap_rfc.fi_document_service import _bridge_python_executable


def main() -> int:
    load_project_env(REPO_ROOT)

    vendor_account = str(os.getenv("SAP_FI_VENDOR_ACCOUNT", "0010000040") or "").strip().upper()
    customer_account = str(os.getenv("SAP_FI_CUSTOMER_ACCOUNT", "0010002949") or "").strip().upper()
    company_code = str(os.getenv("SAP_FI_COMPANY_CODE", "2010") or "").strip().upper()

    bridge_python = _bridge_python_executable()
    if bridge_python is None:
        raise RuntimeError("Bridge Python indisponível para consultar SAP.")

    bridge_script = f"""
from pathlib import Path
import json
import sys

repo_root = Path(r'{REPO_ROOT}')
sys.path.insert(0, str(repo_root))

from sap_rfc._rfc_common import build_connection_params_for_env, load_project_env, make_read_only_guard, read_table
from pyrfc import Connection

load_project_env(repo_root)
connection = Connection(**build_connection_params_for_env('QAD'))

def field_names(table_name):
    result = connection.call('DDIF_FIELDINFO_GET', TABNAME=table_name, LANGU='E')
    rows = result.get('DFIES_TAB') or []
    return [str(row.get('FIELDNAME', '')).strip() for row in rows if isinstance(row, dict) and row.get('FIELDNAME')]

def preview(table_name, options):
    fields = field_names(table_name)
    selected = fields[:12]
    guard = make_read_only_guard((table_name,))
    rows = read_table(
        connection,
        guard,
        table_name=table_name,
        fields=selected,
        options=options,
        rowcount=5,
    )
    return {{
        'field_count': len(fields),
        'selected_fields': selected,
        'row_count': len(rows),
        'rows': [dict(zip(selected, row)) for row in rows],
    }}

report = {{
    'environment': 'QAD',
    'company_code': {json.dumps(company_code)},
    'vendor_account': {json.dumps(vendor_account)},
    'customer_account': {json.dumps(customer_account)},
    'tables': {{}},
}}

try:
    queries = [
        ('T059P', [{{'TEXT': "WITHT = 'P5'"}}]),
        ('T059Z', [{{'TEXT': "WITHT = 'P5'"}}, {{'TEXT': "AND WT_WITHCD = '63'"}}]),
        ('T001WT', [{{'TEXT': f"BUKRS = '{company_code}'"}}, {{'TEXT': "AND WITHT = 'P5'"}}]),
        ('LFBW', [{{'TEXT': f"BUKRS = '{company_code}'"}}, {{'TEXT': f"AND LIFNR = '{vendor_account}'"}}]),
        ('KNBW', [{{'TEXT': f"BUKRS = '{company_code}'"}}, {{'TEXT': f"AND KUNNR = '{customer_account}'"}}]),
    ]
    for table_name, options in queries:
        try:
            report['tables'][table_name] = preview(table_name, options)
        except Exception as exc:
            report['tables'][table_name] = {{
                'error': f"{{exc.__class__.__name__}}: {{exc}}",
            }}
    print(json.dumps(report, ensure_ascii=False, indent=2))
finally:
    connection.close()
"""

    proc = subprocess.run(
        [str(bridge_python), "-c", bridge_script],
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        print(proc.stdout, end="")
        print(proc.stderr, end="", file=sys.stderr)
        return proc.returncode

    print(proc.stdout)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

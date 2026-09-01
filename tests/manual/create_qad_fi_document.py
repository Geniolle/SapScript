from __future__ import annotations

import json
import os
import sys
import subprocess
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import find_project_root, load_project_env
from sap_rfc._rfc_common import build_connection_params_for_env, make_option_eq, make_read_only_guard, read_table
from sap_rfc.fi_document_service import _bridge_python_executable, post_fi_document


def main() -> int:
    project_root = find_project_root()
    load_project_env(project_root)

    tax_code = str(os.getenv("SAP_FI_TAX_CODE", "") or "").strip()
    tax_amount = str(os.getenv("SAP_FI_TAX_AMOUNT", "") or "").strip()
    tax_rate = str(os.getenv("SAP_FI_TAX_RATE", "") or "").strip()
    tax_gl_account = str(os.getenv("SAP_FI_TAX_GL_ACCOUNT", "") or "").strip()
    tax_direction = str(os.getenv("SAP_FI_TAX_DIRECTION", "credit") or "credit").strip()
    payment_method = str(
        os.getenv("SAP_FI_FORM_PAGTO_FORNECEDOR")
        or os.getenv("SAP_FI_FORM_PAGTO")
        or os.getenv("SAP_FI_PAYMENT_METHOD_FORNECEDOR")
        or os.getenv("SAP_FI_PAYMENT_METHOD")
        or os.getenv("SAP_F110_PAYMENT_METHOD")
        or ""
    ).strip().upper()

    payload = {
        "data_mode": "default",
        "company_code": "",
        "posting_date": date.today().isoformat(),
        "document_date": date.today().isoformat(),
        "tax_code": tax_code,
        "tax_amount": tax_amount or "0.00",
        "tax_rate": tax_rate,
        "tax_gl_account": tax_gl_account,
        "tax_direction": tax_direction,
        "payment_method": payment_method,
    }

    result = post_fi_document("QAD", "fornecedor", payload)
    print(
        json.dumps(
            {
                "ok": result.ok,
                "status": result.status,
                "message": result.message,
                "company_code": result.company_code,
                "document_number": result.document_number,
                "branch": result.branch,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    if not result.ok:
        return 1

    document_number = str(result.document_number or "").strip()
    company_code = str(result.company_code or payload["company_code"]).strip().upper()
    fiscal_year = str(payload["posting_date"])[:4]

    bridge_python = _bridge_python_executable()
    if bridge_python is None:
        print("Bridge Python indisponível para ler WITH_ITEM.", file=sys.stderr)
        return 2

    bridge_script = f"""
from pathlib import Path
import json
import sys

repo_root = Path(r'{REPO_ROOT}')
sys.path.insert(0, str(repo_root))

from sap_rfc._rfc_common import build_connection_params_for_env, make_option_eq, make_read_only_guard, read_table
from pyrfc import Connection

params = build_connection_params_for_env('QAD')
connection = Connection(**params)
try:
    guard = make_read_only_guard(('WITH_ITEM',))
    rows = read_table(
        connection,
        guard,
        table_name='WITH_ITEM',
        fields=['BUKRS', 'BELNR', 'GJAHR', 'BUZEI', 'WITHT', 'WT_WITHCD', 'WT_QBSHB', 'WT_QSSHH'],
        options=[],
        rowcount=0,
    )
    rows = [
        row
        for row in rows
        if len(row) > 2
        and str(row[0]).strip().upper() == {json.dumps(company_code)}
        and str(row[1]).strip().zfill(10) == {json.dumps(document_number)}
        and str(row[2]).strip() == {json.dumps(fiscal_year)}
    ]
    print(json.dumps({{
        'with_item_found': bool(rows),
        'with_item_row_count': len(rows),
        'with_item_rows': [
            {{
                'BUKRS': row[0] if len(row) > 0 else '',
                'BELNR': row[1] if len(row) > 1 else '',
                'GJAHR': row[2] if len(row) > 2 else '',
                'BUZEI': row[3] if len(row) > 3 else '',
                'WITHT': row[4] if len(row) > 4 else '',
                'WT_WITHCD': row[5] if len(row) > 5 else '',
                'WT_QBSHB': row[6] if len(row) > 6 else '',
                'WT_QSSHH': row[7] if len(row) > 7 else '',
            }}
            for row in rows
        ],
    }}, ensure_ascii=False, indent=2))
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
        return 2
    print(proc.stdout.strip())
    data = json.loads(proc.stdout or "{}")
    return 0 if data.get("with_item_found") else 2


if __name__ == "__main__":
    raise SystemExit(main())

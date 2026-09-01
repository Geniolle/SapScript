import json
import os
import sys
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from sap_rfc._rfc_common import build_connection_params_for_env, load_project_env, make_read_only_guard, read_table
from sap_rfc.fi_document_service import _bridge_python_executable, post_fi_document


def _read_with_item(document_number: str, company_code: str, fiscal_year: str) -> dict[str, object]:
    bridge_python = _bridge_python_executable()
    if bridge_python is None:
        return {"error": "Bridge Python indisponível"}

    bridge_script = f"""
from pathlib import Path
import json
import sys

repo_root = Path(r'{REPO_ROOT}')
sys.path.insert(0, str(repo_root))

from sap_rfc._rfc_common import build_connection_params_for_env, make_read_only_guard, read_table
from pyrfc import Connection

connection = Connection(**build_connection_params_for_env('QAD'))
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
    proc = __import__("subprocess").run(
        [str(bridge_python), "-c", bridge_script],
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        return {"error": proc.stderr.strip() or proc.stdout.strip() or "bridge_failed"}
    return json.loads(proc.stdout or "{}")


def main() -> int:
    load_project_env(REPO_ROOT)

    payload = {
        "data_mode": "default",
        "company_code": "2010",
        "posting_date": date.today().isoformat(),
        "document_date": date.today().isoformat(),
        "vendor_account": str(os.getenv("SAP_FI_VENDOR_ACCOUNT", "0010000040") or "").strip(),
        "expense_gl_account": str(os.getenv("SAP_FI_EXPENSE_GL_ACCOUNT", "") or "").strip(),
        "amount": "100.00",
        "currency": "EUR",
        "withholding_tax_type": "P5",
        "withholding_tax_code": "63",
        "withholding_tax_base_amount": "100.00",
    }

    if not payload["expense_gl_account"]:
        raise RuntimeError("SAP_FI_EXPENSE_GL_ACCOUNT não está disponível na configuração local.")

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

    with_item = _read_with_item(
        str(result.document_number or ""),
        str(result.company_code or payload["company_code"]),
        str(payload["posting_date"])[:4],
    )
    print(json.dumps(with_item, ensure_ascii=False, indent=2))
    return 0 if with_item.get("with_item_found") else 2


if __name__ == "__main__":
    raise SystemExit(main())

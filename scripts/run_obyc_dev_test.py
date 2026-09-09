from __future__ import annotations

import json
import sys
from pathlib import Path

from dotenv import load_dotenv

PROJECT_DIR = Path(__file__).resolve().parents[1]

if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

load_dotenv(PROJECT_DIR / ".env", override=False)

from sap_rfc.obyc_service import _load_excel_rows_for_validation, validate_obyc_excel  # noqa: E402


def _find_default_excel() -> Path:
    desktop = Path.home() / "OneDrive - Salsajeans" / "Desktop"
    if not desktop.exists():
        raise FileNotFoundError(f"Desktop não encontrado: {desktop}")

    preferred_names = [
        "Configurações OBYC.xlsx",
        "Configuracoes OBYC.xlsx",
    ]
    for name in preferred_names:
        candidate = desktop / name
        if candidate.exists():
            return candidate

    matches = sorted(desktop.glob("*OBYC*.xlsx"))
    if matches:
        return matches[0]

    raise FileNotFoundError(f"Não encontrei ficheiro OBYC em {desktop}")


def main() -> None:
    excel_path = _find_default_excel()
    workbook = _load_excel_rows_for_validation(str(excel_path))

    status, log = validate_obyc_excel(
        {
            "system": "DEV",
            "table": "T030",
            "preview_data": {
                "file_name": excel_path.name,
                "sheet_name": workbook["sheet_name"],
                "headers": workbook["headers"],
                "rows": workbook["rows"],
                "row_count": workbook["row_count"],
            },
        }
    )
    print(status)
    print(log)


if __name__ == "__main__":
    main()

from __future__ import annotations

import os


def main() -> int:
    keys = [
        "SAP_FI_TAX_CODE",
        "SAP_FI_TAX_AMOUNT",
        "SAP_FI_TAX_RATE",
        "SAP_FI_TAX_GL_ACCOUNT",
        "SAP_FI_TAX_DIRECTION",
        "SAP_FI_FORM_PAGTO_FORNECEDOR",
        "SAP_QAD_USER",
        "SAP_QAD_PASSWD",
    ]
    for key in keys:
        print(f"{key}={'set' if os.getenv(key) else 'missing'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

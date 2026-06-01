from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .config import SapConnectionConfig
from .safety import SafetyGuard


class SapRfcUnavailable(RuntimeError):
    pass


@dataclass
class SapRfcClient:
    """Read-only SAP RFC wrapper.

    This client intentionally exposes generic read helpers and blocks mutation-like
    RFC names by default through SafetyGuard.
    """

    config: SapConnectionConfig
    safety_guard: SafetyGuard
    _connection: Any | None = None

    def connect(self) -> None:
        try:
            from pyrfc import Connection  # type: ignore
        except Exception as exc:  # pragma: no cover - depends on local SAP SDK
            raise SapRfcUnavailable(
                "PyRFC is not available. Install SAP NetWeaver RFC SDK and pyrfc."
            ) from exc
        self._connection = Connection(**self.config.as_pyrfc_params())

    @property
    def connection(self) -> Any:
        if self._connection is None:
            self.connect()
        return self._connection

    def call(self, function_name: str, **parameters: Any) -> dict[str, Any]:
        self.safety_guard.assert_function_allowed(function_name)
        result = self.connection.call(function_name, **parameters)
        return dict(result or {})

    def ping(self) -> bool:
        self.call("RFC_PING")
        return True

    def read_table(
        self,
        table_name: str,
        fields: list[str] | None = None,
        options: list[str] | None = None,
        rowcount: int = 50,
    ) -> list[dict[str, str]]:
        self.safety_guard.assert_table_allowed(table_name)
        fields_payload = [{"FIELDNAME": field} for field in (fields or [])]
        options_payload = [{"TEXT": option} for option in (options or [])]
        result = self.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table_name,
            DELIMITER="|",
            FIELDS=fields_payload,
            OPTIONS=options_payload,
            ROWCOUNT=rowcount,
        )
        sap_fields = [entry["FIELDNAME"] for entry in result.get("FIELDS", [])]
        rows: list[dict[str, str]] = []
        for row in result.get("DATA", []):
            values = str(row.get("WA", "")).split("|")
            rows.append({field: values[index].strip() if index < len(values) else "" for index, field in enumerate(sap_fields)})
        return rows

    def get_message_text(self, message_id: str, message_number: str, language: str = "E") -> list[dict[str, str]]:
        return self.read_table(
            "T100",
            fields=["SPRSL", "ARBGB", "MSGNR", "TEXT"],
            options=[
                f"SPRSL = '{language}'",
                f"AND ARBGB = '{message_id.upper()}'",
                f"AND MSGNR = '{message_number.zfill(3)}'",
            ],
            rowcount=10,
        )

    def get_transport_request(self, request: str) -> list[dict[str, str]]:
        return self.read_table(
            "E070",
            fields=["TRKORR", "TRFUNCTION", "TRSTATUS", "AS4USER", "AS4DATE", "AS4TEXT"],
            options=[f"TRKORR = '{request.upper()}'"],
            rowcount=10,
        )

    def get_fi_document_header(self, company_code: str, document_number: str, fiscal_year: str) -> list[dict[str, str]]:
        return self.read_table(
            "BKPF",
            fields=["BUKRS", "BELNR", "GJAHR", "BLART", "BUDAT", "CPUDT", "USNAM", "TCODE", "XBLNR", "BKTXT"],
            options=[
                f"BUKRS = '{company_code}'",
                f"AND BELNR = '{document_number.zfill(10)}'",
                f"AND GJAHR = '{fiscal_year}'",
            ],
            rowcount=5,
        )

    def get_payment_request(self, company_code: str | None = None, keyno: str | None = None, iban: str | None = None) -> list[dict[str, str]]:
        options: list[str] = []
        if company_code:
            options.append(f"ZBUKR = '{company_code}'")
        if keyno:
            options.append(("AND " if options else "") + f"KEYNO = '{keyno}'")
        if iban:
            options.append(("AND " if options else "") + f"ZIBAN = '{iban}'")
        return self.read_table(
            "PAYRQ",
            fields=["ZBUKR", "KEYNO", "DMBTR", "DUEDT", "ZIBAN", "HBKID", "HKTID", "ZLSCH"],
            options=options,
            rowcount=20,
        )

    def get_background_jobs(self, job_name: str, username: str | None = None, rowcount: int = 20) -> list[dict[str, str]]:
        options = [f"JOBNAME = '{job_name}'"]
        if username:
            options.append(f"AND SDLUNAME = '{username.upper()}'")
        return self.read_table(
            "TBTCO",
            fields=["JOBNAME", "JOBCOUNT", "SDLUNAME", "STATUS", "STRTDATE", "STRTTIME", "ENDDATE", "ENDTIME"],
            options=options,
            rowcount=rowcount,
        )

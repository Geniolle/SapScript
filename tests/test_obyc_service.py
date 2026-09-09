from __future__ import annotations

import json
import sys
import types

from sap_rfc import obyc_service


def test_validate_obyc_excel_success_with_preview_data(monkeypatch) -> None:
    monkeypatch.setattr(
        obyc_service,
        "read_obyc_table",
        lambda environment, table, filters, fields: [
            {
                "KTOPL": "1000",
                "KTOSL": "VBR",
                "BWMOD": "01",
                "KOMOK": "A1",
                "BKLAS": "3000",
            }
        ],
    )

    status, log = obyc_service.validate_obyc_excel(
        {
            "system": "DEV",
            "table": "T030",
            "preview_data": {
                "file_name": "obyc.xlsx",
                "sheet_name": "Folha1",
                "headers": ["KTOPL", "KTOSL", "BWMOD", "KOMOK", "BKLAS"],
                "rows": [
                    {
                        "_row_number": 2,
                        "KTOPL": "1000",
                        "KTOSL": "VBR",
                        "BWMOD": "01",
                        "KOMOK": "A1",
                        "BKLAS": "3000",
                    }
                ],
                "row_count": 1,
            },
        }
    )

    result = json.loads(status)
    assert result["ok"] is True
    assert result["validated_rows"] == 1
    assert result["matched_rows"] == 1
    assert "Validação do Excel concluída" in result["message"]
    assert "OBYC executada via RFC" in log


def test_worker_obyc_validate_delegates_to_service(monkeypatch) -> None:
    import sap_script_web_cockpit_v2.worker.sap_tasks as sap_tasks

    captured = {}

    def fake_validate(params):
        captured["params"] = params
        return "{\"ok\": true}", "ok"

    monkeypatch.setattr("sap_rfc.obyc_service.validate_obyc_excel", fake_validate)

    status, log = sap_tasks._run_obyc_excel_validate({"system": "DEV", "table": "T030", "preview_data": {}})

    assert status == "{\"ok\": true}"
    assert log == "ok"
    assert captured["params"]["system"] == "DEV"
    assert captured["params"]["table"] == "T030"


def test_validate_obyc_excel_treats_non_key_fields_as_optional(monkeypatch) -> None:
    monkeypatch.setattr(
        obyc_service,
        "read_obyc_table",
        lambda environment, table, filters, fields: [
            {
                "KTOPL": "1000",
                "KTOSL": "GBB",
                "BWMOD": "F100",
                "KOMOK": "INV",
                "BKLAS": "1000",
                "HKONT": "400000",
            }
        ],
    )

    status, log = obyc_service.validate_obyc_excel(
        {
            "system": "DEV",
            "table": "T030",
            "preview_data": {
                "file_name": "Configurações OBYC.xlsx",
                "sheet_name": "OBYC",
                "headers": ["KTOPL", "KTOSL", "BWMOD", "KOMOK", "BKLAS", "HKONT"],
                "rows": [
                    {
                        "_row_number": 2,
                        "KTOPL": "1000",
                        "KTOSL": "GBB",
                        "BWMOD": "F100",
                        "KOMOK": "INV",
                        "BKLAS": "1000",
                        "HKONT": "500000",
                    }
                ],
                "row_count": 1,
            },
        }
    )

    result = json.loads(status)
    assert result["ok"] is True
    assert result["matched_rows"] == 1
    assert result["optional_rows"] == 1
    assert result["missing_rows"] == 0
    assert result["mismatched_rows"] == 0
    optional_issue = next(issue for issue in result["issues"] if issue["reason"] == "comparacao_opcional")
    assert optional_issue["filters"], "issue de comparacao_opcional deve incluir as chaves que bateram, não 'Sem chaves'"
    assert "Comparações opcionais" in log or "Comparações opcionais" in result["message"]


def test_read_obyc_table_treats_table_without_data_as_empty(monkeypatch) -> None:
    class FakeRfcError(Exception):
        def __init__(self) -> None:
            super().__init__("RFC_READ_TABLE returned no rows")
            self.key = "TABLE_WITHOUT_DATA"
            self.message = "TABLE_WITHOUT_DATA"

    class FakeConnection:
        def __init__(self, **kwargs) -> None:
            self.kwargs = kwargs

        def call(self, *args, **kwargs):
            raise FakeRfcError()

        def close(self) -> None:
            pass

    monkeypatch.setattr(obyc_service, "Connection", FakeConnection)
    monkeypatch.setattr(
        obyc_service,
        "_obyc_connection_params",
        lambda environment: {
            "user": "U",
            "passwd": "P",
            "ashost": "localhost",
            "sysnr": "00",
            "client": "100",
            "lang": "EN",
        },
    )
    fake_pyrfc = types.ModuleType("pyrfc")
    fake_pyrfc.Connection = FakeConnection
    monkeypatch.setitem(sys.modules, "pyrfc", fake_pyrfc)

    rows = obyc_service._obyc_read_table_core(
        "DEV",
        "T030",
        [{"field": "KTOPL", "value": "1000"}],
        ["KTOPL", "KTOSL"],
    )

    assert rows == []

"""Conexão RFC falsa para testes (sem SAP)."""

from __future__ import annotations

from typing import Any


class FakeRfcError(Exception):
    def __init__(self, key: str) -> None:
        super().__init__(key)
        self.key = key
        self.message = key


class FakeConnection:
    """Emula `RFC_READ_TABLE` sobre tabelas em memória.

    `tables` = { "PPDIT": {"fields": [(name, type, length, text), ...],
                            "rows": [ {col: value, ...}, ... ] } }
    """

    def __init__(self, tables: dict[str, dict[str, Any]], *,
                 raise_for: dict[str, str] | None = None,
                 functions: dict[str, Any] | None = None):
        self.tables = tables
        self.raise_for = raise_for or {}
        # functions: {"FUNCNAME": dict-to-return  |  Exception  |  callable(**kwargs)->dict}
        self.functions = functions or {}
        self.calls: list[tuple[str, dict[str, Any]]] = []

    def call(self, function_name: str, **kwargs: Any) -> dict[str, Any]:
        self.calls.append((function_name, kwargs))
        if function_name == "RFC_PING":
            return {}
        if function_name in self.functions:
            handler = self.functions[function_name]
            if isinstance(handler, BaseException):
                raise handler
            if callable(handler):
                return dict(handler(**kwargs) or {})
            return dict(handler or {})
        if function_name != "RFC_READ_TABLE":
            raise FakeRfcError(f"UNEXPECTED_FUNCTION:{function_name}")

        table = kwargs["QUERY_TABLE"]
        if table in self.raise_for:
            raise FakeRfcError(self.raise_for[table])
        if table not in self.tables:
            raise FakeRfcError("TABLE_NOT_AVAILABLE")

        spec = self.tables[table]
        all_fields = spec["fields"]
        by_name = {f[0]: f for f in all_fields}
        requested = [f["FIELDNAME"] for f in kwargs.get("FIELDS", [])]
        if requested:
            # RFC_READ_TABLE devolve as colunas pela ordem pedida em FIELDS.
            sel = [by_name[n] for n in requested if n in by_name]
        else:
            sel = list(all_fields)

        fields_meta = [
            {"FIELDNAME": n, "TYPE": t, "LENGTH": str(ln), "OFFSET": "0", "FIELDTEXT": tx}
            for (n, t, ln, tx) in sel
        ]

        rows = spec["rows"]
        skip = int(kwargs.get("ROWSKIPS", 0) or 0)
        count = int(kwargs.get("ROWCOUNT", 0) or 0)
        window = rows[skip:] if count == 0 else rows[skip : skip + count]
        data = [
            {"WA": "|".join(str(r.get(n, "")) for (n, _t, _l, _x) in sel)}
            for r in window
        ]
        return {"FIELDS": fields_meta, "DATA": data}

    def close(self) -> None:  # pragma: no cover
        pass

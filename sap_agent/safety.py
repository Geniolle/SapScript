from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable


_MUTATION_KEYWORDS = (
    "CREATE",
    "CHANGE",
    "UPDATE",
    "DELETE",
    "POST",
    "COMMIT",
    "ROLLBACK",
    "SAVE",
    "START",
    "RUN_START",
    "ENQUEUE",
    "DEQUEUE",
    "BAPI_TRANSACTION_COMMIT",
)


@dataclass(frozen=True)
class SafetyGuard:
    """Central guard to keep the SAP agent read-only.

    All SAP RFC calls must pass through this guard before execution.
    """

    allow_write_operations: bool = False
    allowed_functions: tuple[str, ...] = ()
    allowed_tables: tuple[str, ...] = ()

    @classmethod
    def build(
        cls,
        allow_write_operations: bool,
        allowed_functions: Iterable[str],
        allowed_tables: Iterable[str],
    ) -> "SafetyGuard":
        return cls(
            allow_write_operations=allow_write_operations,
            allowed_functions=tuple(name.upper() for name in allowed_functions),
            allowed_tables=tuple(name.upper() for name in allowed_tables),
        )

    def assert_function_allowed(self, function_name: str) -> None:
        upper_name = function_name.upper()
        if self.allowed_functions and upper_name not in self.allowed_functions:
            raise PermissionError(f"RFC function is not in whitelist: {function_name}")
        if not self.allow_write_operations and any(keyword in upper_name for keyword in _MUTATION_KEYWORDS):
            raise PermissionError(f"RFC function blocked by read-only guard: {function_name}")

    def assert_table_allowed(self, table_name: str) -> None:
        upper_name = table_name.upper()
        if self.allowed_tables and upper_name not in self.allowed_tables:
            raise PermissionError(f"SAP table is not in whitelist: {table_name}")

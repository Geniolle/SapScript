from __future__ import annotations

from dataclasses import dataclass, field
import os
import re
from collections.abc import Iterable
from typing import Any


SOURCE_LINE_KEYS = ("LINE", "TEXT", "SOURCE_LINE", "ABAP")
INCLUDE_KEYS = ("INCLUDE", "NAME", "PROGRAM", "PROGNAME")
CLASS_METADATA_FUNCTIONS = (
    "SEO_CLASS_TYPEINFO_GET_RFC",
    "SEO_CLASS_GET",
    "SEO_CLASS_READ",
)
CLASS_NAME_KEYS = ("CLSNAME", "CLASS_NAME", "CLASS", "NAME")
SOURCE_NAME_KEYS = (
    "MAINPROGRAM",
    "MAIN_PROG",
    "PROGRAM",
    "PROGRAM_NAME",
    "PROGNAME",
    "INCLUDE",
    "INCLUDE_NAME",
)
ABAP_NAME_RE = re.compile(r"\b[A-Z][A-Z0-9_/$#]{2,}\b")
STOPWORDS = {
    "ACTIVE",
    "ATTRIBUTE",
    "ATTRIBUTES",
    "CLASS",
    "CLASSNAME",
    "COMPONENT",
    "COMPONENTS",
    "DEFINITION",
    "DEFINITIONS",
    "FUNCTION",
    "FUNCTIONS",
    "INCLUDE",
    "INCLUDES",
    "INTERFACE",
    "LANGUAGE",
    "METHOD",
    "METHODS",
    "NAME",
    "NAMES",
    "PROGRAM",
    "PROGRAMS",
    "PUBLIC",
    "PRIVATE",
    "PROTECTED",
    "SOURCE",
    "SOURCE_TAB",
    "TYPEINFO",
    "VERSION",
}


@dataclass
class AbapSourceHit:
    root_name: str
    program_name: str
    include_name: str
    marker: str
    line_number: int
    line_text: str


@dataclass
class AbapClassMetadataProbe:
    class_name: str
    function_name: str | None = None
    request_parameters: dict[str, Any] = field(default_factory=dict)
    response: dict[str, Any] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)


def _language() -> str:
    return os.getenv("SAP_QAD_LANG", "PT").strip() or "PT"


def _normalize_lines(source: Any) -> list[str]:
    if isinstance(source, str):
        return source.splitlines()
    if not isinstance(source, list):
        return []

    lines: list[str] = []
    for item in source:
        if isinstance(item, dict):
            for key in SOURCE_LINE_KEYS:
                value = item.get(key)
                if value is not None:
                    lines.append(str(value))
                    break
            else:
                lines.append(str(item))
        else:
            lines.append(str(item))
    return lines


def _normalize_includes(response: dict[str, Any]) -> list[str]:
    includes: list[str] = []
    for key in ("INCLUDE_TAB", "INCLUDETAB", "INCLUDES"):
        rows = response.get(key) or []
        if isinstance(rows, dict):
            rows = [rows]
        for row in rows:
            if isinstance(row, dict):
                for field in INCLUDE_KEYS:
                    value = str(row.get(field) or "").strip().upper()
                    if value:
                        includes.append(value)
                        break
            else:
                value = str(row).strip().upper()
                if value:
                    includes.append(value)
    deduped: list[str] = []
    seen: set[str] = set()
    for include in includes:
        if include not in seen:
            seen.add(include)
            deduped.append(include)
    return deduped


def read_program_source(connection: Any, program_name: str, *, language: str | None = None) -> dict[str, Any]:
    return dict(
        connection.call(  # type: ignore[attr-defined]
            "RPY_PROGRAM_READ",
            PROGRAM_NAME=program_name,
            LANGUAGE=language or _language(),
            WITH_INCLUDELIST="X",
            ONLY_SOURCE="X",
            READ_LATEST_VERSION="X",
            WITH_LOWERCASE="X",
        )
        or {}
    )


def read_program_source_lines(connection: Any, program_name: str, *, language: str | None = None) -> list[str]:
    response = read_program_source(connection, program_name, language=language)
    return _normalize_lines(response.get("SOURCE") or response.get("SOURCE_TAB"))


def read_program_includes(connection: Any, program_name: str, *, language: str | None = None) -> list[str]:
    response = read_program_source(connection, program_name, language=language)
    return _normalize_includes(response)


def read_class_metadata(connection: Any, class_name: str, *, language: str | None = None) -> AbapClassMetadataProbe:
    normalized = class_name.strip().upper()
    probe = AbapClassMetadataProbe(class_name=normalized)
    if not normalized:
        probe.errors.append("Nome de classe vazio.")
        return probe

    for function_name in CLASS_METADATA_FUNCTIONS:
        for param_name in CLASS_NAME_KEYS:
            request = {param_name: normalized}
            if language:
                request["LANGUAGE"] = language
            try:
                response = dict(connection.call(function_name, **request) or {})  # type: ignore[attr-defined]
            except Exception as exc:  # noqa: BLE001 - diagnostic helper
                probe.errors.append(f"{function_name}({param_name}) -> {exc}")
                continue

            probe.function_name = function_name
            probe.request_parameters = request
            probe.response = response
            return probe

    if not probe.errors:
        probe.errors.append("Nenhuma combinação de RFC/parametro devolveu metadata de classe.")
    return probe


def extract_abap_names(value: Any) -> list[str]:
    found: list[str] = []

    def visit(item: Any) -> None:
        if isinstance(item, str):
            for match in ABAP_NAME_RE.findall(item.upper()):
                if match not in STOPWORDS:
                    found.append(match)
            return
        if isinstance(item, dict):
            for key, child in item.items():
                upper_key = str(key).upper()
                if upper_key in STOPWORDS:
                    continue
                visit(child)
            return
        if isinstance(item, Iterable) and not isinstance(item, (bytes, bytearray)):
            for child in item:
                visit(child)

    visit(value)

    deduped: list[str] = []
    seen: set[str] = set()
    for name in found:
        normalized = name.strip().upper()
        if normalized and normalized not in seen:
            seen.add(normalized)
            deduped.append(normalized)
    return deduped


def extract_metadata_candidates(response: dict[str, Any]) -> list[str]:
    candidates: list[str] = []
    for key in SOURCE_NAME_KEYS:
        value = response.get(key)
        if isinstance(value, str):
            candidates.extend(extract_abap_names(value))
        elif value is not None:
            candidates.extend(extract_abap_names(value))
    candidates.extend(extract_abap_names(response))

    deduped: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        if candidate not in seen:
            seen.add(candidate)
            deduped.append(candidate)
    return deduped


def collect_source_hits(
    connection: Any,
    root_name: str,
    markers: tuple[str, ...],
    *,
    language: str | None = None,
    visited: set[str] | None = None,
    origin_name: str | None = None,
) -> list[AbapSourceHit]:
    normalized = root_name.strip().upper()
    if not normalized:
        return []

    visited = visited or set()
    if normalized in visited:
        return []
    visited.add(normalized)

    origin = origin_name or normalized

    try:
        response = read_program_source(connection, normalized, language=language)
    except Exception:
        return []

    lines = _normalize_lines(response.get("SOURCE") or response.get("SOURCE_TAB"))
    hits: list[AbapSourceHit] = []
    for line_number, line_text in enumerate(lines, start=1):
        upper = line_text.upper()
        for marker in markers:
            if marker in upper:
                hits.append(
                    AbapSourceHit(
                        root_name=origin,
                        program_name=normalized,
                        include_name=normalized,
                        marker=marker,
                        line_number=line_number,
                        line_text=line_text.strip(),
                    )
                )

    for include in _normalize_includes(response):
        hits.extend(
            collect_source_hits(
                connection,
                include,
                markers,
                language=language,
                visited=visited,
                origin_name=origin,
            )
        )
    return hits


def resolve_abap_roots(
    connection: Any,
    object_name: str,
    *,
    object_kind: str = "PROG",
    language: str | None = None,
) -> tuple[list[str], AbapClassMetadataProbe | None]:
    kind = str(object_kind or "PROG").strip().upper()
    normalized = object_name.strip().upper()
    if not normalized:
        return [], None

    if kind in {"PROG", "PROGRAM", "REPORT", "INCLUDE"}:
        return [normalized], None

    if kind in {"CLAS", "CLASS", "INTERFACE", "INTF"}:
        probe = read_class_metadata(connection, normalized, language=language)
        roots = extract_metadata_candidates(probe.response)
        if normalized not in roots:
            roots.insert(0, normalized)
        return roots, probe

    return [normalized], None

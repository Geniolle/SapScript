from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sap_session import load_dotenv_manual


@dataclass(frozen=True)
class RfcTarget:
    user: str
    passwd: str
    ashost: str
    sysnr: str
    client: str
    lang: str
    source: str


@dataclass(frozen=True)
class RfcConnectionConfig:
    user: str
    passwd: str
    ashost: str
    sysnr: str
    client: str
    lang: str

    def as_pyrfc_params(self) -> dict[str, str]:
        return {
            "user": self.user,
            "passwd": self.passwd,
            "ashost": self.ashost,
            "sysnr": self.sysnr,
            "client": self.client,
            "lang": self.lang,
        }


class ReadOnlyRfcClient:
    def __init__(self, config: RfcConnectionConfig):
        try:
            from pyrfc import Connection  # type: ignore
        except Exception as exc:  # pragma: no cover - local SAP SDK dependent
            raise RuntimeError(f"PyRFC indisponivel: {exc}") from exc

        self._connection = Connection(**config.as_pyrfc_params())

    def ping(self) -> bool:
        self._connection.call("RFC_PING")
        return True

    def read_table(
        self,
        table_name: str,
        fields: list[str] | None = None,
        options: list[str] | None = None,
        rowcount: int = 100,
        rowskips: int = 0,
    ) -> list[dict[str, str]]:
        result = self._connection.call(
            "RFC_READ_TABLE",
            QUERY_TABLE=table_name,
            DELIMITER="|",
            FIELDS=[{"FIELDNAME": field} for field in (fields or [])],
            OPTIONS=[{"TEXT": option} for option in (options or [])],
            ROWCOUNT=rowcount,
            ROWSKIPS=rowskips,
            GET_SORTED="X",
        )
        sap_fields = [entry["FIELDNAME"] for entry in result.get("FIELDS", [])]
        rows: list[dict[str, str]] = []
        for row in result.get("DATA", []):
            values = str(row.get("WA", "")).split("|")
            rows.append(
                {
                    field: values[index].strip() if index < len(values) else ""
                    for index, field in enumerate(sap_fields)
                }
            )
        return rows

    def call(self, function_name: str, **parameters: Any) -> dict[str, Any]:
        result = self._connection.call(function_name, **parameters)
        return dict(result or {})


def _first_non_empty(*values: str) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def _resolve_rfc_target() -> RfcTarget:
    candidates: list[RfcTarget] = []

    def add_generic() -> None:
        candidates.append(
            RfcTarget(
                user=os.getenv("SAP_USER", "").strip(),
                passwd=os.getenv("SAP_PASSWD", "").strip(),
                ashost=os.getenv("SAP_ASHOST", "").strip(),
                sysnr=os.getenv("SAP_SYSNR", "").strip(),
                client=_first_non_empty(os.getenv("SAP_CLIENT", ""), os.getenv("SAP_CLIENT_S4DCLNT100", "")),
                lang=_first_non_empty(os.getenv("SAP_LANG", ""), os.getenv("SAP_LANGUAGE", "PT")),
                source="SAP",
            )
        )

    def add_prefixed(prefix: str) -> None:
        candidates.append(
            RfcTarget(
                user=os.getenv(f"{prefix}_USER", "").strip(),
                passwd=os.getenv(f"{prefix}_PASSWD", "").strip(),
                ashost=os.getenv(f"{prefix}_ASHOST", "").strip(),
                sysnr=os.getenv(f"{prefix}_SYSNR", "").strip(),
                client=os.getenv(f"{prefix}_CLIENT", "").strip(),
                lang=os.getenv(f"{prefix}_LANG", "PT").strip() or "PT",
                source=prefix,
            )
        )

    add_generic()
    for prefix in ("SAP_DEV", "SAP_PRD", "SAP_R3", "SAP_BW"):
        add_prefixed(prefix)

    for target in candidates:
        if target.user and target.passwd and target.ashost and target.sysnr and target.client:
            return target

    raise RuntimeError(
        "No usable RFC target found. Expected SAP_USER or one of SAP_DEV_/SAP_PRD_/SAP_R3_/SAP_BW_ variables."
    )


def _build_client() -> ReadOnlyRfcClient:
    target = _resolve_rfc_target()
    config = RfcConnectionConfig(
        user=target.user,
        passwd=target.passwd,
        ashost=target.ashost,
        sysnr=target.sysnr,
        client=target.client,
        lang=target.lang,
    )
    return ReadOnlyRfcClient(config=config)


def _read_table_all(
    client: ReadOnlyRfcClient,
    table_name: str,
    fields: list[str],
    options: list[str],
    rowcount: int = 100,
) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    skip = 0
    while True:
        chunk = client.read_table(
            table_name,
            fields=fields,
            options=options,
            rowcount=rowcount,
            rowskips=skip,
        )
        if not chunk:
            break
        rows.extend(chunk)
        if len(chunk) < rowcount:
            break
        skip += len(chunk)
    return rows


def _safe_read_table_all(
    client: ReadOnlyRfcClient,
    table_name: str,
    fields: list[str],
    options: list[str],
    rowcount: int = 100,
) -> list[dict[str, str]]:
    try:
        return _read_table_all(client, table_name, fields=fields, options=options, rowcount=rowcount)
    except Exception as exc:
        print(f"\n{table_name} / erro: {exc}")
        return []


def _safe_call(client: ReadOnlyRfcClient, function_name: str, **parameters: Any) -> dict[str, Any]:
    try:
        return client.call(function_name, **parameters)
    except Exception as exc:
        print(f"\n{function_name} / erro: {exc}")
        return {}


def _print_rows(title: str, rows: list[dict[str, Any]], limit: int = 10) -> None:
    print(f"\n{title}")
    if not rows:
        print("  (sem resultados)")
        return
    for row in rows[:limit]:
        print("  - " + ", ".join(f"{key}={value}" for key, value in row.items()))
    if len(rows) > limit:
        print(f"  ... {len(rows) - limit} resultado(s) adicional(is)")


def _classify_node(row: dict[str, str]) -> str:
    if row.get("MP_EXIT_FUNC", "").strip():
        return "exit"
    if row.get("MP_CONST", "").strip():
        return "constante"
    if row.get("MP_SC_TAB", "").strip() or row.get("MP_SC_FLD", "").strip():
        return "campo_origem"
    if row.get("MP_SC_NODE", "").strip() or row.get("REF_NAME", "").strip():
        return "referencia_arvore"
    return "arvore_estrutura"


def _describe_node(row: dict[str, str], text_by_node: dict[str, dict[str, str]]) -> dict[str, str]:
    node_id = row.get("NODE_ID", "").strip()
    text_row = text_by_node.get(node_id, {})
    return {
        "tree_type": row.get("TREE_TYPE", "").strip(),
        "tree_id": row.get("TREE_ID", "").strip(),
        "version": row.get("VERSION", "").strip(),
        "node_id": node_id,
        "tech_name": row.get("TECH_NAME", "").strip(),
        "node_text": text_row.get("TEXT", "").strip(),
        "node_comment": text_row.get("NODE_COMMENT", "").strip(),
        "node_type": row.get("NODE_TYPE", "").strip(),
        "parent_id": row.get("PARENT_ID", "").strip(),
        "brother_id": row.get("BROTHER_ID", "").strip(),
        "firstchild_id": row.get("FIRSTCHILD_ID", "").strip(),
        "level": row.get("LEV", "").strip(),
        "data_type": row.get("DATA_TYPE", "").strip(),
        "length": row.get("LENGTH", "").strip(),
        "classification": _classify_node(row),
        "mp_const": row.get("MP_CONST", "").strip(),
        "mp_sc_tab": row.get("MP_SC_TAB", "").strip(),
        "mp_sc_fld": row.get("MP_SC_FLD", "").strip(),
        "mp_sc_node": row.get("MP_SC_NODE", "").strip(),
        "mp_sc_ref_name": row.get("MP_SC_REF_NAME", "").strip(),
        "mp_exit_func": row.get("MP_EXIT_FUNC", "").strip(),
        "mp_selection": row.get("MP_SELECTION", "").strip(),
        "cv_rule": row.get("CV_RULE", "").strip(),
        "ref_name": row.get("REF_NAME", "").strip(),
        "tab_keyfld": row.get("TAB_KEYFLD", "").strip(),
        "atom_handl": row.get("ATOM_HANDL", "").strip(),
        "ex_status": row.get("EX_STATUS", "").strip(),
    }


def _print_fieldinfo(client: ReadOnlyRfcClient, table_name: str, langu: str) -> None:
    result = _safe_call(client, "DDIF_FIELDINFO_GET", TABNAME=table_name, LANGU=langu)
    fields = result.get("DFIES_TAB", []) or []
    print(f"\nDDIF_FIELDINFO_GET / {table_name}")
    if not fields:
        print("  (sem campos)")
        return
    for field in fields[:40]:
        print(
            "  - "
            + ", ".join(
                [
                    f"FIELDNAME={field.get('FIELDNAME', '')}",
                    f"ROLLNAME={field.get('ROLLNAME', '')}",
                    f"DATATYPE={field.get('DATATYPE', '')}",
                    f"LENG={field.get('LENG', '')}",
                ]
            )
        )
    if len(fields) > 40:
        print(f"  ... {len(fields) - 40} campo(s) adicional(is)")


def _build_export(client: ReadOnlyRfcClient, tree_id: str) -> dict[str, Any]:
    tree_rows = _safe_read_table_all(
        client,
        "DMEE_TREE",
        fields=[
            "TREE_TYPE",
            "TREE_ID",
            "CREA_USER",
            "CREA_DATE",
            "CREA_TIME",
            "CHNG_USER",
            "CHNG_DATE",
            "CHNG_TIME",
            "DOCU_TXT",
            "RELEASE_FLAG",
            "ORIG_LANGU",
            "PARENT_ID",
            "TREE_LEVEL",
            "DMEEX",
            "EXTENSIBLE",
        ],
        options=[f"TREE_ID = '{tree_id}'"],
        rowcount=20,
    )

    head_rows = _safe_read_table_all(
        client,
        "DMEE_TREE_HEAD",
        fields=[
            "TREE_TYPE",
            "TREE_ID",
            "VERSION",
            "FIRSTNODE_ID",
            "VERSION_DESCRIPTION",
        ],
        options=[f"TREE_ID = '{tree_id}'"],
        rowcount=20,
    )

    node_rows = _safe_read_table_all(
        client,
        "DMEE_TREE_NODE",
        fields=[
            "TREE_TYPE",
            "TREE_ID",
            "VERSION",
            "NODE_ID",
            "TECH_NAME",
            "REF_NAME",
            "PARENT_ID",
            "BROTHER_ID",
            "FIRSTCHILD_ID",
            "NODE_TYPE",
            "LENGTH",
            "DATA_TYPE",
            "EX_STATUS",
            "LEV",
            "TAB_KEYFLD",
            "ATOM_HANDL",
            "MP_SC_TAB",
            "MP_SC_FLD",
            "MP_SC_NODE",
            "MP_SC_REF_NAME",
            "MP_CONST",
            "CV_RULE",
            "MP_EXIT_FUNC",
            "MP_SELECTION",
        ],
        options=[f"TREE_ID = '{tree_id}'"],
        rowcount=100,
    )

    tree_lang = tree_rows[0].get("ORIG_LANGU", "").strip() if tree_rows else ""
    node_text_options = [f"TREE_ID = '{tree_id}'"]
    if tree_lang:
        node_text_options.append(f"AND LANGU = '{tree_lang}'")
    node_text_rows = _safe_read_table_all(
        client,
        "DMEE_TREE_NODE_T",
        fields=["LANGU", "TREE_TYPE", "TREE_ID", "VERSION", "NODE_ID", "TEXT", "NODE_COMMENT"],
        options=node_text_options,
        rowcount=100,
    )
    text_by_node = {row.get("NODE_ID", "").strip(): row for row in node_text_rows}

    rule_rows = _safe_read_table_all(
        client,
        "DMEE_TREE_RULES",
        fields=["TREE_TYPE", "TREE_ID", "VERSION", "RULE_NUMBER", "ID_VALUE", "ID_OFFSET", "SEGM_REF_ID", "SEGM_NODE_ID"],
        options=[f"TREE_ID = '{tree_id}'"],
        rowcount=100,
    )

    described_nodes = [_describe_node(row, text_by_node) for row in node_rows]
    described_nodes.sort(key=lambda row: (row["version"], row["level"], row["node_id"]))

    classifications = {
        "constante": sum(1 for row in described_nodes if row["classification"] == "constante"),
        "campo_origem": sum(1 for row in described_nodes if row["classification"] == "campo_origem"),
        "referencia_arvore": sum(1 for row in described_nodes if row["classification"] == "referencia_arvore"),
        "exit": sum(1 for row in described_nodes if row["classification"] == "exit"),
        "arvore_estrutura": sum(1 for row in described_nodes if row["classification"] == "arvore_estrutura"),
    }

    return {
        "requested_tree_id": tree_id,
        "tree": tree_rows[0] if tree_rows else {},
        "headers": head_rows,
        "nodes": described_nodes,
        "rules": rule_rows,
        "summary": {
            "tree_count": len(tree_rows),
            "header_count": len(head_rows),
            "node_count": len(described_nodes),
            "rule_count": len(rule_rows),
            "classifications": classifications,
        },
    }


def _search_parent_links(client: ReadOnlyRfcClient, parent_id: str) -> list[dict[str, str]]:
    return _safe_read_table_all(
        client,
        "DMEE_TREE",
        fields=["TREE_TYPE", "TREE_ID", "PARENT_ID", "DOCU_TXT", "RELEASE_FLAG", "ORIG_LANGU", "DMEEX", "EXTENSIBLE"],
        options=[f"PARENT_ID = '{parent_id}'"],
        rowcount=100,
    )


def _search_tree_prefix(client: ReadOnlyRfcClient, prefix: str) -> list[dict[str, str]]:
    return _safe_read_table_all(
        client,
        "DMEE_TREE",
        fields=["TREE_TYPE", "TREE_ID", "PARENT_ID", "DOCU_TXT", "RELEASE_FLAG", "ORIG_LANGU", "DMEEX", "EXTENSIBLE"],
        options=[f"TREE_ID LIKE '{prefix}%'"],
        rowcount=100,
    )


def _print_overview(export_data: dict[str, Any]) -> None:
    tree = export_data.get("tree", {})
    summary = export_data.get("summary", {})
    classifications = summary.get("classifications", {})

    print("\nResumo da árvore")
    if tree:
        print(
            "  "
            + ", ".join(
                [
                    f"TREE_TYPE={tree.get('TREE_TYPE', '')}",
                    f"TREE_ID={tree.get('TREE_ID', '')}",
                    f"DOCU_TXT={tree.get('DOCU_TXT', '')}",
                    f"ORIG_LANGU={tree.get('ORIG_LANGU', '')}",
                    f"DMEEX={tree.get('DMEEX', '')}",
                    f"EXTENSIBLE={tree.get('EXTENSIBLE', '')}",
                ]
            )
        )
    print(f"  Nós: {summary.get('node_count', 0)}")
    print(f"  Regras: {summary.get('rule_count', 0)}")
    print(f"  Constantes: {classifications.get('constante', 0)}")
    print(f"  Campos de origem: {classifications.get('campo_origem', 0)}")
    print(f"  Referências à árvore: {classifications.get('referencia_arvore', 0)}")
    print(f"  Exits: {classifications.get('exit', 0)}")
    print(f"  Estrutura pura: {classifications.get('arvore_estrutura', 0)}")


def _print_sample_nodes(nodes: list[dict[str, Any]], limit: int = 30) -> None:
    print("\nPrimeiros nós")
    for row in nodes[:limit]:
        print(
            "  - "
            + ", ".join(
                [
                    f"NODE_ID={row.get('node_id', '')}",
                    f"TECH_NAME={row.get('tech_name', '')}",
                    f"TYPE={row.get('node_type', '')}",
                    f"CLASS={row.get('classification', '')}",
                    f"TEXT={row.get('node_text', '')}",
                    f"CONST={row.get('mp_const', '')}",
                    f"SRC={row.get('mp_sc_tab', '')}.{row.get('mp_sc_fld', '')}",
                    f"REF={row.get('mp_sc_node', '')}",
                    f"EXIT={row.get('mp_exit_func', '')}",
                ]
            )
        )
    if len(nodes) > limit:
        print(f"  ... {len(nodes) - limit} nó(s) adicional(is)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect SAP DMEEX trees via RFC")
    parser.add_argument("--tree-id", default="Z_PT_CGI_XML_CT_V9", help="DMEE tree ID to inspect")
    parser.add_argument(
        "--find-parent",
        default="",
        help="Search for trees that reference the given TREE_ID in DMEE_TREE.PARENT_ID",
    )
    parser.add_argument(
        "--find-prefix",
        default="",
        help="Search for trees whose TREE_ID starts with the given prefix",
    )
    parser.add_argument(
        "--export-json",
        default="",
        help="Write a complete JSON export to this path. If omitted, a default file under output/ is used.",
    )
    args = parser.parse_args()

    load_dotenv_manual()
    client = _build_client()

    try:
        client.ping()
    except Exception as exc:
        raise SystemExit(f"PyRFC indisponivel ou conexao falhou: {exc}") from exc

    target = _resolve_rfc_target()
    print(f"RFC target source: {target.source}")
    print(f"RFC target client: {target.client}")

    export_data = _build_export(client, args.tree_id)
    _print_overview(export_data)

    tree = export_data.get("tree", {})
    if tree:
        _print_fieldinfo(client, "DMEE_TREE", langu=tree.get("ORIG_LANGU", "")[:1].upper() or target.lang[:1].upper() or "E")
        _print_fieldinfo(client, "DMEE_TREE_HEAD", langu=tree.get("ORIG_LANGU", "")[:1].upper() or target.lang[:1].upper() or "E")
        _print_fieldinfo(client, "DMEE_TREE_NODE", langu=tree.get("ORIG_LANGU", "")[:1].upper() or target.lang[:1].upper() or "E")
        _print_fieldinfo(client, "DMEE_TREE_NODE_T", langu=tree.get("ORIG_LANGU", "")[:1].upper() or target.lang[:1].upper() or "E")
        _print_fieldinfo(client, "DMEE_TREE_RULES", langu=tree.get("ORIG_LANGU", "")[:1].upper() or target.lang[:1].upper() or "E")

    _print_sample_nodes(export_data.get("nodes", []), limit=25)

    if args.find_parent.strip():
        parent_hits = _search_parent_links(client, args.find_parent.strip())
        print(f"\nFilhos com PARENT_ID={args.find_parent.strip()}")
        _print_rows("DMEE_TREE / parent links", parent_hits, limit=50)

    if args.find_prefix.strip():
        prefix_hits = _search_tree_prefix(client, args.find_prefix.strip())
        print(f"\nÁrvores com TREE_ID começado por {args.find_prefix.strip()}")
        _print_rows("DMEE_TREE / prefix search", prefix_hits, limit=100)

    export_path = args.export_json.strip()
    if not export_path:
        export_path = f"output/dmee_{args.tree_id}.json"

    target_path = Path(export_path)
    if not target_path.is_absolute():
        target_path = (ROOT / target_path).resolve()
    target_path.parent.mkdir(parents=True, exist_ok=True)
    with target_path.open("w", encoding="utf-8") as handle:
        json.dump(export_data, handle, ensure_ascii=False, indent=2)
    print(f"\nJSON exportado para: {target_path}")


if __name__ == "__main__":
    main()

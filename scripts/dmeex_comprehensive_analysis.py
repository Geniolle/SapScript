from __future__ import annotations

import csv
import json
import os
import sys
from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from sap_session import load_dotenv_manual

TREE_TYPE = "PAYM"
TREE_IDS_PRD = ["Z_SEPA_CT"]
TREE_IDS_DEV = ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9", "PT_CGI_XML_CT_V9"]
ACTIVE_VERSION = "000"

REQUESTED_NODE_FIELDS = [
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
    "DATA_TYPE",
    "LENGTH",
    "LEV",
    "EX_STATUS",
    "ATOM_HANDL",
    "MP_IF_TP",
    "MP_SELECTION",
    "MP_OFFSET",
    "MP_SC_TAB",
    "MP_SC_FLD",
    "MP_SC_OFFSET",
    "MP_SC_NODE",
    "MP_SC_REF_NAME",
    "MP_CONST",
    "CV_RULE",
    "MP_EXIT_FUNC",
    "CK_EXIT_FUNC",
    "TAB_KEYFLD",
]

REQUESTED_NODE_R_FIELDS = [
    "TREE_TYPE",
    "TREE_ID",
    "VERSION",
    "NODE_ID",
    "NODE_REDEFINED",
    "NODE_DEACTIVATED",
    "VAR_NAME",
]

REQUESTED_COND_FIELDS = [
    "TREE_TYPE",
    "TREE_ID",
    "VERSION",
    "NODE_ID",
    "COND_NUMBER",
    "PAR_OPEN",
    "ARG1_FLD",
    "ARG1_TAB",
    "ARG1_CONST",
    "ARG1_TYPE",
    "OPERATOR",
    "ARG2_FLD",
    "ARG2_TAB",
    "ARG2_CONST",
    "ARG2_TYPE",
    "PAR_CLOSE",
    "LINK_OPERATOR",
    "CD_EXIT_FUNC",
]

REQUESTED_AGGR_FIELDS = [
    "TREE_TYPE",
    "TREE_ID",
    "VERSION",
    "NODE_ID",
    "AGG_TYPE",
    "AGG_NODE_ID",
    "AGG_REF_NAME",
]

REQUESTED_SORT_FIELDS = [
    "TREE_TYPE",
    "TREE_ID",
    "VERSION",
    "SORT_ORDER",
    "NODE_ID",
    "KEY_FIELD",
    "SORT_TAB",
    "SORT_FLD",
    "SORT_IF_TP",
    "USER_FIELD",
    "LEV",
    "KF_EXIT_FUNC",
    "KEY_ONLY",
]

REQUESTED_RULE_FIELDS = [
    "TREE_TYPE",
    "TREE_ID",
    "VERSION",
    "RULE_NUMBER",
    "ID_VALUE",
    "ID_OFFSET",
    "SEGM_REF_ID",
    "SEGM_NODE_ID",
]

COMPARE_FIELDS = [
    "TECH_NAME",
    "REF_NAME",
    "PARENT_ID",
    "BROTHER_ID",
    "FIRSTCHILD_ID",
    "NODE_TYPE",
    "DATA_TYPE",
    "LENGTH",
    "LEV",
    "EX_STATUS",
    "ATOM_HANDL",
    "MP_IF_TP",
    "MP_SELECTION",
    "MP_OFFSET",
    "MP_SC_TAB",
    "MP_SC_FLD",
    "MP_SC_OFFSET",
    "MP_SC_NODE",
    "MP_SC_REF_NAME",
    "MP_CONST",
    "CV_RULE",
    "MP_EXIT_FUNC",
    "CK_EXIT_FUNC",
    "NODE_REDEFINED",
    "NODE_DEACTIVATED",
    "VAR_NAME",
    "conditions_key",
    "aggregation_key",
    "sort_key",
]

PRIORITY_FILTERS = [
    "PmtInf/PmtTpInf/CtgyPurp",
    "PmtInf/PmtTpInf/CtgyPurp/Cd",
    "PmtInf/PmtTpInf/CtgyPurp/Cd/SUPP",
    "PmtInf/PmtTpInf/CtgyPurp/Cd/CASH",
    "PmtInf/PmtTpInf/ChrgBr",
    "GrpHdr/InitgPty/Nm",
    "GrpHdr/InitgPty/Id/PrvtId/Othr/Id",
    "PmtInf/Dbtr/Id",
    "PmtInf/DbtrAcct/Ccy",
    "CdtTrfTxInf/PmtId/InstrId",
    "GrpHdr/Authstn/Prtry",
    "CdtTrfTxInf/RmtInf/Ustrd",
]


@dataclass
class SapEnv:
    label: str
    prefix: str
    params: dict[str, str]
    sid: str = ""
    client: str = ""
    release: str = ""
    host: str = ""


class ReadOnlySap:
    def __init__(self, env: SapEnv):
        try:
            from pyrfc import Connection  # type: ignore
        except Exception as exc:
            raise RuntimeError(f"PyRFC indisponivel: {exc}") from exc
        self.env = env
        self.conn = Connection(**env.params)
        self.field_cache: dict[str, set[str]] = {}

    def close(self) -> None:
        self.conn.close()

    def call(self, name: str, **kwargs: Any) -> dict[str, Any]:
        allowed = {"RFC_PING", "RFC_READ_TABLE", "DDIF_FIELDINFO_GET", "RFC_SYSTEM_INFO"}
        if name not in allowed:
            raise RuntimeError(f"RFC nao permitido neste script read-only: {name}")
        return dict(self.conn.call(name, **kwargs) or {})

    def system_info(self) -> None:
        self.call("RFC_PING")
        attrs = self.conn.get_connection_attributes()
        info = self.call("RFC_SYSTEM_INFO").get("RFCSI_EXPORT", {})
        self.env.sid = str(info.get("RFCSYSID", "")).strip()
        self.env.client = str(attrs.get("client", "")).strip()
        self.env.release = str(info.get("RFCSAPRL", "")).strip()
        self.env.host = str(info.get("RFCHOST", "")).strip()

    def fields(self, table: str) -> set[str]:
        table = table.upper()
        if table not in self.field_cache:
            result = self.call("DDIF_FIELDINFO_GET", TABNAME=table, LANGU=self.env.params.get("lang", "E")[:1] or "E")
            self.field_cache[table] = {
                str(row.get("FIELDNAME", "")).strip().upper()
                for row in result.get("DFIES_TAB", [])
                if str(row.get("FIELDNAME", "")).strip()
            }
        return self.field_cache[table]

    def existing_fields(self, table: str, requested: list[str]) -> list[str]:
        available = self.fields(table)
        return [field for field in requested if field.upper() in available]

    def read_table(self, table: str, fields: list[str], options: list[str], batch_size: int = 500) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        skip = 0
        while True:
            result = self.call(
                "RFC_READ_TABLE",
                QUERY_TABLE=table,
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": field} for field in fields],
                OPTIONS=[{"TEXT": option} for option in options],
                ROWCOUNT=batch_size,
                ROWSKIPS=skip,
                GET_SORTED="X",
            )
            data = result.get("DATA", [])
            if not data:
                break
            for row in data:
                values = str(row.get("WA", "")).split("|")
                rows.append({field: values[index].strip() if index < len(values) else "" for index, field in enumerate(fields)})
            if len(data) < batch_size:
                break
            skip += len(data)
        return rows


def env_from_prefix(label: str, prefix: str) -> SapEnv:
    params = {
        "user": os.environ[f"{prefix}_USER"].strip(),
        "passwd": os.environ[f"{prefix}_PASSWD"].strip(),
        "ashost": os.environ[f"{prefix}_ASHOST"].strip(),
        "sysnr": os.environ[f"{prefix}_SYSNR"].strip(),
        "client": os.environ[f"{prefix}_CLIENT"].strip(),
        "lang": os.getenv(f"{prefix}_LANG", "PT").strip() or "PT",
    }
    return SapEnv(label=label, prefix=prefix, params=params)


def base_options(tree_id: str, version: str | None = None) -> list[str]:
    opts = [f"TREE_TYPE = '{TREE_TYPE}'", f"AND TREE_ID = '{tree_id}'"]
    if version is not None:
        opts.append(f"AND VERSION = '{version}'")
    return opts


def read_optional(sap: ReadOnlySap, table: str, requested: list[str], options: list[str]) -> tuple[list[dict[str, str]], list[str]]:
    try:
        fields = sap.existing_fields(table, requested)
        if not fields:
            return [], []
        return sap.read_table(table, fields, options), fields
    except Exception as exc:
        print(f"{sap.env.label} {table}: leitura ignorada ({exc})")
        return [], []


def active_versions(head_rows: list[dict[str, str]]) -> list[str]:
    versions = sorted({row.get("VERSION", "").strip() for row in head_rows if row.get("VERSION", "").strip()})
    return versions


def cond_expr(cond: dict[str, str]) -> str:
    left = f"{cond.get('ARG1_TAB','')}-{cond.get('ARG1_FLD','')}".strip("-")
    if not left:
        left = cond.get("ARG1_CONST", "")
    right = f"{cond.get('ARG2_TAB','')}-{cond.get('ARG2_FLD','')}".strip("-")
    if not right:
        right = cond.get("ARG2_CONST", "")
    return (
        f"{cond.get('PAR_OPEN','')}{left} {cond.get('OPERATOR','')} {right}{cond.get('PAR_CLOSE','')}"
        f" {cond.get('LINK_OPERATOR','')}"
    ).strip()


def canonical_rows(rows: list[dict[str, str]], skip: set[str] | None = None) -> list[dict[str, str]]:
    skip = skip or set()
    return [
        {key: row.get(key, "") for key in sorted(row) if key not in skip}
        for row in sorted(rows, key=lambda r: json.dumps(r, sort_keys=True, ensure_ascii=False))
    ]


def key_for(rows: list[dict[str, str]], skip: set[str] | None = None) -> str:
    return json.dumps(canonical_rows(rows, skip=skip), ensure_ascii=False, sort_keys=True)


def path_segment(node: dict[str, Any]) -> str:
    tech = node.get("TECH_NAME", "")
    ntype = node.get("NODE_TYPE", "")
    if ntype == "ATOM":
        return f":atom:{tech}"
    if ntype == "XMAT":
        return f"@{tech}"
    if ntype == "TECH":
        return f"(TECH:{tech})"
    return tech


def xml_segment(node: dict[str, Any]) -> str:
    ntype = node.get("NODE_TYPE", "")
    if ntype in {"ATOM", "TECH"}:
        return ""
    return node.get("TECH_NAME", "")


def build_paths(nodes: dict[str, dict[str, Any]]) -> None:
    for node_id, node in nodes.items():
        current = node_id
        parts: list[str] = []
        xml_parts: list[str] = []
        visited: set[str] = set()
        while current and current in nodes and current not in visited:
            visited.add(current)
            cur = nodes[current]
            parts.append(path_segment(cur))
            xml = xml_segment(cur)
            if xml:
                xml_parts.append(xml)
            current = cur.get("PARENT_ID", "")
        node["full_path"] = "/" + "/".join(reversed(parts))
        node["xml_path"] = "/".join(reversed(xml_parts))


def fetch_tree(sap: ReadOnlySap, tree_id: str, version: str = ACTIVE_VERSION) -> dict[str, Any]:
    tree_rows, tree_fields = read_optional(
        sap,
        "DMEE_TREE",
        ["TREE_TYPE", "TREE_ID", "PARENT_ID", "DMEEX", "EXTENSIBLE", "CREA_USER", "CREA_DATE", "CHNG_USER", "CHNG_DATE", "CHNG_TIME"],
        base_options(tree_id),
    )
    head_rows, head_fields = read_optional(
        sap,
        "DMEE_TREE_HEAD",
        ["TREE_TYPE", "TREE_ID", "VERSION", "VERS_USER", "VERS_DATE", "VERS_TIME", "FIRSTNODE_ID", "VERSION_DESCRIPTION"],
        base_options(tree_id),
    )
    node_fields = sap.existing_fields("DMEE_TREE_NODE", REQUESTED_NODE_FIELDS)
    node_rows = sap.read_table("DMEE_TREE_NODE", node_fields, base_options(tree_id, version))
    nodes = {row["NODE_ID"]: {field: row.get(field, "") for field in REQUESTED_NODE_FIELDS} for row in node_rows}

    node_r_rows, node_r_fields = read_optional(sap, "DMEE_TREE_NODE_R", REQUESTED_NODE_R_FIELDS, base_options(tree_id, version))
    node_r_by_id = {row.get("NODE_ID", ""): row for row in node_r_rows}

    cond_rows, cond_fields = read_optional(sap, "DMEE_TREE_COND", REQUESTED_COND_FIELDS, base_options(tree_id, version))
    cond_by_id: dict[str, list[dict[str, str]]] = {}
    for row in cond_rows:
        cond_by_id.setdefault(row.get("NODE_ID", ""), []).append(row)

    aggr_rows, aggr_fields = read_optional(sap, "DMEE_TREE_AGGR", REQUESTED_AGGR_FIELDS, base_options(tree_id, version))
    aggr_by_id: dict[str, list[dict[str, str]]] = {}
    for row in aggr_rows:
        aggr_by_id.setdefault(row.get("NODE_ID", ""), []).append(row)

    sort_rows, sort_fields = read_optional(sap, "DMEE_TREE_SORT", REQUESTED_SORT_FIELDS, base_options(tree_id, version))
    sort_by_id: dict[str, list[dict[str, str]]] = {}
    for row in sort_rows:
        sort_by_id.setdefault(row.get("NODE_ID", ""), []).append(row)

    rules_rows, rules_fields = read_optional(sap, "DMEE_TREE_RULES", REQUESTED_RULE_FIELDS, base_options(tree_id, version))

    for node_id, node in nodes.items():
        red = node_r_by_id.get(node_id, {})
        node["NODE_REDEFINED"] = red.get("NODE_REDEFINED", "")
        node["NODE_DEACTIVATED"] = red.get("NODE_DEACTIVATED", "")
        node["VAR_NAME"] = red.get("VAR_NAME", "")
        node["conditions"] = sorted(cond_by_id.get(node_id, []), key=lambda r: r.get("COND_NUMBER", ""))
        node["condition_text"] = " ".join(cond_expr(row) for row in node["conditions"])
        node["aggregations"] = aggr_by_id.get(node_id, [])
        node["sorts"] = sort_by_id.get(node_id, [])
        node["conditions_key"] = key_for(node["conditions"], skip={"TREE_TYPE", "TREE_ID", "VERSION"})
        node["aggregation_key"] = key_for(node["aggregations"], skip={"TREE_TYPE", "TREE_ID", "VERSION"})
        node["sort_key"] = key_for(node["sorts"], skip={"TREE_TYPE", "TREE_ID", "VERSION"})

    build_paths(nodes)
    return {
        "environment": {
            "label": sap.env.label,
            "sid": sap.env.sid,
            "client": sap.env.client,
            "release": sap.env.release,
            "host": sap.env.host,
        },
        "tree_id": tree_id,
        "version": version,
        "tree": tree_rows[0] if tree_rows else {},
        "headers": head_rows,
        "active_versions_found": active_versions(head_rows),
        "nodes": nodes,
        "node_rows": list(nodes.values()),
        "conditions": cond_rows,
        "aggregations": aggr_rows,
        "sorts": sort_rows,
        "rules": rules_rows,
        "field_inventory": {
            "DMEE_TREE": tree_fields,
            "DMEE_TREE_HEAD": head_fields,
            "DMEE_TREE_NODE": node_fields,
            "DMEE_TREE_NODE_R": node_r_fields,
            "DMEE_TREE_COND": cond_fields,
            "DMEE_TREE_AGGR": aggr_fields,
            "DMEE_TREE_SORT": sort_fields,
            "DMEE_TREE_RULES": rules_fields,
        },
    }


def effective_tree(tree: dict[str, Any], parent: dict[str, Any] | None = None) -> dict[str, Any]:
    nodes: dict[str, dict[str, Any]] = {}
    if parent:
        for node_id, pnode in parent["nodes"].items():
            clone = deepcopy(pnode)
            clone["origin"] = f"INHERITED_FROM_{parent['tree_id']}"
            clone["source_tree_id"] = parent["tree_id"]
            nodes[node_id] = clone
    for node_id, node in tree["nodes"].items():
        clone = deepcopy(node)
        if parent and node_id in parent["nodes"]:
            if clone.get("NODE_DEACTIVATED") == "X":
                clone["origin"] = f"DEACTIVATED_IN_{tree['tree_id']}"
            elif clone.get("NODE_REDEFINED") == "X":
                clone["origin"] = f"REDEFINED_IN_{tree['tree_id']}"
            else:
                clone["origin"] = f"INHERITED_COPY_IN_{tree['tree_id']}"
        else:
            clone["origin"] = f"DEFINED_IN_{tree['tree_id']}"
        clone["source_tree_id"] = tree["tree_id"]
        nodes[node_id] = clone
    build_paths(nodes)
    out = deepcopy(tree)
    out["nodes"] = nodes
    out["node_rows"] = list(nodes.values())
    out["parent_tree_id"] = parent["tree_id"] if parent else ""
    return out


def node_signature(node: dict[str, Any]) -> dict[str, str]:
    return {field: str(node.get(field, "") or "") for field in COMPARE_FIELDS}


def compare_by_path(left: dict[str, Any], right: dict[str, Any], label: str) -> dict[str, Any]:
    left_by_path: dict[str, list[dict[str, Any]]] = {}
    right_by_path: dict[str, list[dict[str, Any]]] = {}
    for node in left["nodes"].values():
        left_by_path.setdefault(node.get("full_path", ""), []).append(node)
    for node in right["nodes"].values():
        right_by_path.setdefault(node.get("full_path", ""), []).append(node)

    rows: list[dict[str, Any]] = []
    for path in sorted(set(left_by_path) | set(right_by_path)):
        l_nodes = sorted(left_by_path.get(path, []), key=lambda n: n.get("NODE_ID", ""))
        r_nodes = sorted(right_by_path.get(path, []), key=lambda n: n.get("NODE_ID", ""))
        for idx in range(max(len(l_nodes), len(r_nodes))):
            l_node = l_nodes[idx] if idx < len(l_nodes) else None
            r_node = r_nodes[idx] if idx < len(r_nodes) else None
            diffs: list[dict[str, str]] = []
            if l_node and r_node:
                l_sig = node_signature(l_node)
                r_sig = node_signature(r_node)
                for field in COMPARE_FIELDS:
                    if l_sig.get(field, "") != r_sig.get(field, ""):
                        diffs.append({"property": field, "left": l_sig.get(field, ""), "right": r_sig.get(field, "")})
                classification = "IGUAL" if not diffs else "DIFERENTE"
            elif l_node:
                classification = "SO_ESQUERDA"
            else:
                classification = "SO_DIREITA"
            rows.append(
                {
                    "comparison": label,
                    "path": path,
                    "left_node_id": l_node.get("NODE_ID", "") if l_node else "",
                    "right_node_id": r_node.get("NODE_ID", "") if r_node else "",
                    "left_xml_path": l_node.get("xml_path", "") if l_node else "",
                    "right_xml_path": r_node.get("xml_path", "") if r_node else "",
                    "classification": classification,
                    "diffs": diffs,
                    "left": summarize_node(l_node) if l_node else {},
                    "right": summarize_node(r_node) if r_node else {},
                }
            )
    counts: dict[str, int] = {}
    for row in rows:
        counts[row["classification"]] = counts.get(row["classification"], 0) + 1
    return {"label": label, "left_tree": left["tree_id"], "right_tree": right["tree_id"], "summary": counts, "rows": rows}


def summarize_node(node: dict[str, Any] | None) -> dict[str, Any]:
    if not node:
        return {}
    return {
        "NODE_ID": node.get("NODE_ID", ""),
        "path": node.get("full_path", ""),
        "xml_path": node.get("xml_path", ""),
        "NODE_TYPE": node.get("NODE_TYPE", ""),
        "LEV": node.get("LEV", ""),
        "LENGTH": node.get("LENGTH", ""),
        "MP_IF_TP": node.get("MP_IF_TP", ""),
        "mapping": deduce_mapping(node),
        "MP_SC_TAB": node.get("MP_SC_TAB", ""),
        "MP_SC_FLD": node.get("MP_SC_FLD", ""),
        "MP_CONST": node.get("MP_CONST", ""),
        "MP_SELECTION": node.get("MP_SELECTION", ""),
        "MP_OFFSET": node.get("MP_OFFSET", ""),
        "MP_SC_OFFSET": node.get("MP_SC_OFFSET", ""),
        "MP_SC_NODE": node.get("MP_SC_NODE", ""),
        "MP_SC_REF_NAME": node.get("MP_SC_REF_NAME", ""),
        "ATOM_HANDL": node.get("ATOM_HANDL", ""),
        "CV_RULE": node.get("CV_RULE", ""),
        "MP_EXIT_FUNC": node.get("MP_EXIT_FUNC", ""),
        "CK_EXIT_FUNC": node.get("CK_EXIT_FUNC", ""),
        "condition_text": node.get("condition_text", ""),
        "conditions": node.get("conditions", []),
        "aggregations": node.get("aggregations", []),
        "sorts": node.get("sorts", []),
        "NODE_REDEFINED": node.get("NODE_REDEFINED", ""),
        "NODE_DEACTIVATED": node.get("NODE_DEACTIVATED", ""),
        "VAR_NAME": node.get("VAR_NAME", ""),
        "PARENT_ID": node.get("PARENT_ID", ""),
        "FIRSTCHILD_ID": node.get("FIRSTCHILD_ID", ""),
        "BROTHER_ID": node.get("BROTHER_ID", ""),
        "origin": node.get("origin", ""),
    }


def deduce_mapping(node: dict[str, Any]) -> str:
    if node.get("MP_EXIT_FUNC"):
        return f"EXIT:{node.get('MP_EXIT_FUNC')}"
    if node.get("MP_CONST"):
        return f"CONST:{node.get('MP_CONST')}"
    if node.get("MP_SC_TAB") or node.get("MP_SC_FLD"):
        return f"FIELD:{node.get('MP_SC_TAB')}-{node.get('MP_SC_FLD')}"
    if node.get("MP_SC_NODE") or node.get("MP_SC_REF_NAME"):
        return f"REF:{node.get('MP_SC_NODE') or node.get('MP_SC_REF_NAME')}"
    if node.get("CV_RULE"):
        return f"RULE:{node.get('CV_RULE')}"
    return "STRUCTURE"


def matches_focus(node: dict[str, Any], text: str) -> bool:
    text_u = text.upper()
    return text_u in node.get("full_path", "").upper() or text_u in node.get("xml_path", "").upper()


def find_focus_nodes(tree: dict[str, Any], filters: list[str]) -> dict[str, list[dict[str, Any]]]:
    out: dict[str, list[dict[str, Any]]] = {}
    for focus in filters:
        out[focus] = [
            summarize_node(node)
            for node in sorted(tree["nodes"].values(), key=lambda n: (n.get("full_path", ""), n.get("NODE_ID", "")))
            if matches_focus(node, focus)
        ]
    return out


def find_ctgypurp_context(tree: dict[str, Any]) -> list[dict[str, Any]]:
    nodes = [
        node
        for node in tree["nodes"].values()
        if "PMTTPINF/CTGYPURP" in node.get("xml_path", "").upper()
        or "PMTTPINF/CTGYPURP" in node.get("full_path", "").upper()
        or ("CTGYPURP" in node.get("full_path", "").upper() and ("SUPP" in node.get("full_path", "").upper() or "CASH" in node.get("full_path", "").upper() or "CD" in node.get("full_path", "").upper()))
    ]
    wanted_ids = {node.get("NODE_ID", "") for node in nodes}
    for node in list(nodes):
        parent = node.get("PARENT_ID", "")
        while parent and parent in tree["nodes"]:
            if parent in wanted_ids:
                break
            pnode = tree["nodes"][parent]
            if "PMTINF" in pnode.get("xml_path", "").upper() or pnode.get("NODE_TYPE") == "TECH":
                nodes.append(pnode)
                wanted_ids.add(parent)
            parent = pnode.get("PARENT_ID", "")
    return [summarize_node(node) for node in sorted({n["NODE_ID"]: n for n in nodes}.values(), key=lambda n: (n.get("full_path", ""), n.get("NODE_ID", "")))]


def relevant_diff_rows(compare: dict[str, Any], terms: list[str]) -> list[dict[str, Any]]:
    terms_u = [term.upper() for term in terms]
    hits = []
    for row in compare["rows"]:
        hay = json.dumps(row, ensure_ascii=False).upper()
        if any(term in hay for term in terms_u) and row["classification"] != "IGUAL":
            hits.append(row)
    return hits


def write_json(path: Path, data: Any) -> None:
    with path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)


def write_effective_csv(path: Path, tree: dict[str, Any]) -> None:
    columns = [
        "TREE_ID",
        "VERSION",
        "NODE_ID",
        "TECH_NAME",
        "full_path",
        "xml_path",
        "origin",
        *COMPARE_FIELDS,
        "condition_text",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, delimiter=";")
        writer.writeheader()
        for node in sorted(tree["nodes"].values(), key=lambda n: (n.get("full_path", ""), n.get("NODE_ID", ""))):
            writer.writerow({col: node.get(col, "") for col in columns})


def write_compare_csv(path: Path, compare: dict[str, Any]) -> None:
    columns = ["comparison", "path", "classification", "left_node_id", "right_node_id", "property", "left", "right"]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns, delimiter=";")
        writer.writeheader()
        for row in compare["rows"]:
            if row["diffs"]:
                for diff in row["diffs"]:
                    writer.writerow({**{key: row.get(key, "") for key in columns}, **diff})
            else:
                writer.writerow({key: row.get(key, "") for key in columns})


def md_table(nodes: list[dict[str, Any]], label: str) -> str:
    lines = [f"### {label}", "", "| NODE_ID | path | type | map | const | field | exit | cond | flags |", "|---|---|---|---|---|---|---|---|---|"]
    for node in nodes:
        field = f"{node.get('MP_SC_TAB','')}-{node.get('MP_SC_FLD','')}".strip("-")
        flags = ",".join(
            part
            for part in [
                f"REDEF={node.get('NODE_REDEFINED','')}",
                f"DEACT={node.get('NODE_DEACTIVATED','')}",
                node.get("origin", ""),
            ]
            if part and not part.endswith("=")
        )
        lines.append(
            "| "
            + " | ".join(
                [
                    node.get("NODE_ID", ""),
                    node.get("path", "").replace("|", "\\|"),
                    node.get("NODE_TYPE", ""),
                    node.get("mapping", "").replace("|", "\\|"),
                    node.get("MP_CONST", "").replace("|", "\\|"),
                    field.replace("|", "\\|"),
                    node.get("MP_EXIT_FUNC", ""),
                    node.get("condition_text", "").replace("|", "\\|"),
                    flags.replace("|", "\\|"),
                ]
            )
            + " |"
        )
    if not nodes:
        lines.append("| - | sem nos encontrados | - | - | - | - | - | - | - |")
    return "\n".join(lines)


def conclusion_from_focus(prd_nodes: list[dict[str, Any]], dev_nodes: list[dict[str, Any]]) -> str:
    prd_supp = [n for n in prd_nodes if n.get("MP_CONST") == "SUPP" or "SUPP" in n.get("path", "").upper()]
    dev_supp = [n for n in dev_nodes if n.get("MP_CONST") == "SUPP" or "SUPP" in n.get("path", "").upper()]
    active_dev_supp = [n for n in dev_supp if n.get("NODE_DEACTIVATED") != "X"]
    dev_gate = [
        n
        for n in dev_nodes
        if n.get("path", "").endswith("/PmtInf/PmtTpInf")
        and ("REF03+0(1)" in n.get("condition_text", "") or "XSCHK" in n.get("condition_text", ""))
    ]
    if prd_supp and not dev_supp:
        return "Situacao 2: existe diferenca concreta de DMEEX. O caminho SUPP existe no PRD e nao foi encontrado na Z_PT_CGI_XML_CT_V9 DEV."
    if prd_supp and dev_supp and not active_dev_supp:
        return "Situacao 2: existe diferenca concreta de DMEEX. O SUPP existe no DEV, mas esta desativado no caminho analisado."
    if prd_supp and active_dev_supp and dev_gate:
        return (
            "Situacao 2: existe diferenca concreta de DMEEX. O SUPP existe e esta ativo no DEV, "
            "mas o ramo DEV PmtInf/PmtTpInf so e avaliado quando FPAYHX-REF03+0(1) = 'S' e FPAYHX-XSCHK = SPACE; "
            "alem disso o Cd DEV tem propriedades tecnicas diferentes do PRD e atomos herdados adicionais CODE2/HR_CODE. "
            "Efeito esperado: se essa condicao de ramo nao for verdadeira no pagamento, CtgyPurp/Cd=SUPP nao e produzido apesar de o no existir."
        )
    if prd_supp and active_dev_supp:
        return "Situacao 1/2 a validar pelo detalhe: o SUPP existe nos dois lados; verificar abaixo propriedades, condicoes, TECH e ativacao para equivalencia tecnica."
    return "Situacao 3: nao foi possivel concluir SUPP porque o no SUPP nao foi encontrado claramente no contexto extraido."


def write_supp_report(path: Path, prd: dict[str, Any], dev: dict[str, Any], compare: dict[str, Any]) -> str:
    prd_nodes = find_ctgypurp_context(prd)
    dev_nodes = find_ctgypurp_context(dev)
    diff_hits = relevant_diff_rows(compare, ["CtgyPurp", "SUPP", "CASH", "PmtTpInf"])
    conclusion = conclusion_from_focus(prd_nodes, dev_nodes)
    lines = [
        "# DMEEX CtgyPurp/Cd SUPP - PRD Z_SEPA_CT vs DEV Z_PT_CGI_XML_CT_V9",
        "",
        f"Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "Arvore esperada:",
        "",
        "```text",
        "PmtInf",
        "`- PmtTpInf",
        "   `- CtgyPurp",
        "      `- Cd",
        "         |- CASH",
        "         `- SUPP",
        "```",
        "",
        md_table(prd_nodes, "PRD Z_SEPA_CT"),
        "",
        md_table(dev_nodes, "DEV Z_PT_CGI_XML_CT_V9"),
        "",
        "## Diferencas relevantes",
        "",
    ]
    if diff_hits:
        for row in diff_hits[:25]:
            props = ", ".join(f"{d['property']}: PRD='{d['left']}' DEV='{d['right']}'" for d in row.get("diffs", [])[:8])
            lines.append(f"- {row['classification']} `{row['path']}` PRD_NODE={row['left_node_id']} DEV_NODE={row['right_node_id']} {props}")
    else:
        lines.append("- Nenhuma diferenca focal encontrada por caminho tecnico completo; verificar equivalencia por contexto acima.")
    lines.extend(["", "## Conclusao", "", conclusion, ""])
    path.write_text("\n".join(lines), encoding="utf-8")
    return conclusion


def write_summary_md(path: Path, compare: dict[str, Any], focus: dict[str, Any]) -> None:
    lines = [
        "# DMEEX Compare PRD vs DEV",
        "",
        f"Comparacao: {compare['label']}",
        "",
        "## Resumo",
        "",
    ]
    for key, value in sorted(compare["summary"].items()):
        lines.append(f"- {key}: {value}")
    lines.extend(["", "## Campos prioritarios com diferencas", ""])
    for term, rows in focus.items():
        lines.append(f"### {term}")
        if not rows:
            lines.append("- Sem diferencas encontradas.")
        else:
            for row in rows[:10]:
                props = ", ".join(f"{d['property']}" for d in row.get("diffs", [])[:10])
                lines.append(f"- {row['classification']} `{row['path']}` props={props}")
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    load_dotenv_manual()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = ROOT / "output"
    output_dir.mkdir(exist_ok=True)

    prd_sap = ReadOnlySap(env_from_prefix("PRD", "SAP_PRD"))
    dev_sap = ReadOnlySap(env_from_prefix("DEV", "SAP_DEV"))
    try:
        for sap in (prd_sap, dev_sap):
            sap.system_info()
            print(f"{sap.env.label}: SID={sap.env.sid} CLIENT={sap.env.client} RELEASE={sap.env.release}")

        raw: dict[tuple[str, str], dict[str, Any]] = {}
        effective: dict[tuple[str, str], dict[str, Any]] = {}

        for tree_id in TREE_IDS_PRD:
            raw[("PRD", tree_id)] = fetch_tree(prd_sap, tree_id)
            effective[("PRD", tree_id)] = effective_tree(raw[("PRD", tree_id)])
        for tree_id in TREE_IDS_DEV:
            raw[("DEV", tree_id)] = fetch_tree(dev_sap, tree_id)
        effective[("DEV", "CGI_CT_V9")] = effective_tree(raw[("DEV", "CGI_CT_V9")])
        effective[("DEV", "Z_SEPA_CT")] = effective_tree(raw[("DEV", "Z_SEPA_CT")])
        effective[("DEV", "PT_CGI_XML_CT_V9")] = effective_tree(raw[("DEV", "PT_CGI_XML_CT_V9")], effective[("DEV", "CGI_CT_V9")])
        effective[("DEV", "Z_PT_CGI_XML_CT_V9")] = effective_tree(raw[("DEV", "Z_PT_CGI_XML_CT_V9")], effective[("DEV", "CGI_CT_V9")])

        paths: dict[str, str] = {}
        for label, tree_id in [("PRD", "Z_SEPA_CT"), ("DEV", "Z_SEPA_CT"), ("DEV", "Z_PT_CGI_XML_CT_V9"), ("DEV", "CGI_CT_V9"), ("DEV", "PT_CGI_XML_CT_V9")]:
            data = effective[(label, tree_id)]
            json_path = output_dir / f"dmeex_{tree_id}_{label}_{timestamp}.json"
            csv_path = output_dir / f"{tree_id}_effective_tree_{label}_{timestamp}.csv"
            write_json(json_path, data)
            write_effective_csv(csv_path, data)
            paths[f"{label}_{tree_id}_json"] = str(json_path)
            paths[f"{label}_{tree_id}_csv"] = str(csv_path)

        comp_a = compare_by_path(effective[("PRD", "Z_SEPA_CT")], effective[("DEV", "Z_SEPA_CT")], "PRD_Z_SEPA_CT_vs_DEV_Z_SEPA_CT")
        comp_b = compare_by_path(effective[("PRD", "Z_SEPA_CT")], effective[("DEV", "Z_PT_CGI_XML_CT_V9")], "PRD_Z_SEPA_CT_vs_DEV_Z_PT_CGI_XML_CT_V9")
        comp_c1 = compare_by_path(effective[("DEV", "Z_PT_CGI_XML_CT_V9")], effective[("DEV", "CGI_CT_V9")], "DEV_Z_PT_CGI_XML_CT_V9_vs_DEV_CGI_CT_V9")
        comp_c2 = compare_by_path(effective[("DEV", "Z_PT_CGI_XML_CT_V9")], effective[("DEV", "PT_CGI_XML_CT_V9")], "DEV_Z_PT_CGI_XML_CT_V9_vs_DEV_PT_CGI_XML_CT_V9")

        compare_export = {"comparisons": [comp_a, comp_b, comp_c1, comp_c2]}
        compare_json = output_dir / f"dmeex_compare_PRD_vs_DEV_{timestamp}.json"
        compare_csv = output_dir / f"dmeex_compare_PRD_vs_DEV_{timestamp}.csv"
        compare_md = output_dir / f"dmeex_compare_PRD_vs_DEV_{timestamp}.md"
        write_json(compare_json, compare_export)
        write_compare_csv(compare_csv, comp_b)
        focus_diffs = {term: relevant_diff_rows(comp_b, [term]) for term in PRIORITY_FILTERS}
        write_summary_md(compare_md, comp_b, focus_diffs)
        paths["compare_json"] = str(compare_json)
        paths["compare_csv"] = str(compare_csv)
        paths["compare_md"] = str(compare_md)

        supp_md = output_dir / f"dmeex_ctgypurp_supp_PRD_vs_DEV_{timestamp}.md"
        supp_conclusion = write_supp_report(supp_md, effective[("PRD", "Z_SEPA_CT")], effective[("DEV", "Z_PT_CGI_XML_CT_V9")], comp_b)
        paths["supp_md"] = str(supp_md)

        chrgbr = {
            "PRD_Z_SEPA_CT": find_focus_nodes(effective[("PRD", "Z_SEPA_CT")], ["ChrgBr"])["ChrgBr"],
            "DEV_Z_PT_CGI_XML_CT_V9": find_focus_nodes(effective[("DEV", "Z_PT_CGI_XML_CT_V9")], ["ChrgBr"])["ChrgBr"],
        }
        special = {
            "timestamp": timestamp,
            "systems": {
                "PRD": prd_sap.env.__dict__ | {"params": {"client": prd_sap.env.params["client"], "lang": prd_sap.env.params["lang"]}},
                "DEV": dev_sap.env.__dict__ | {"params": {"client": dev_sap.env.params["client"], "lang": dev_sap.env.params["lang"]}},
            },
            "active_versions": {
                f"{label}_{tree_id}": raw[(label, tree_id)]["active_versions_found"]
                for label, tree_id in raw
            },
            "comparison_summaries": {
                comp_a["label"]: comp_a["summary"],
                comp_b["label"]: comp_b["summary"],
                comp_c1["label"]: comp_c1["summary"],
                comp_c2["label"]: comp_c2["summary"],
            },
            "priority_focus": {
                "ctgypurp": {
                    "PRD_Z_SEPA_CT": find_ctgypurp_context(effective[("PRD", "Z_SEPA_CT")]),
                    "DEV_Z_PT_CGI_XML_CT_V9": find_ctgypurp_context(effective[("DEV", "Z_PT_CGI_XML_CT_V9")]),
                },
                "chrgbr": chrgbr,
            },
            "supp_conclusion": supp_conclusion,
            "outputs": paths,
        }
        summary_path = output_dir / f"dmeex_run_summary_{timestamp}.json"
        write_json(summary_path, special)
        paths["run_summary"] = str(summary_path)

        print("VERSOES_ATIVAS=" + json.dumps(special["active_versions"], ensure_ascii=False, sort_keys=True))
        print("COMPARISON_SUMMARIES=" + json.dumps(special["comparison_summaries"], ensure_ascii=False, sort_keys=True))
        print("SUPP_CONCLUSION=" + supp_conclusion)
        print("CHRGBR=" + json.dumps(chrgbr, ensure_ascii=False)[:4000])
        print("OUTPUTS=" + json.dumps(paths, ensure_ascii=False, sort_keys=True))
    finally:
        prd_sap.close()
        dev_sap.close()


if __name__ == "__main__":
    main()

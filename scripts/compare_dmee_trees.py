from __future__ import annotations

import argparse
import json
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]


CONFIG_FIELDS = (
    "node_type",
    "classification",
    "data_type",
    "length",
    "mp_const",
    "mp_exit_func",
    "mp_selection",
    "cv_rule",
    "ref_name",
    "atom_handl",
    "tab_keyfld",
    "mp_sc_ref_name",
    "mp_sc_node",
    "ex_status",
)


@dataclass(frozen=True)
class TreeSpec:
    label: str
    path: Path


def _load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _source_key(node: dict[str, Any]) -> str | None:
    tab = str(node.get("mp_sc_tab", "") or "").strip().upper()
    fld = str(node.get("mp_sc_fld", "") or "").strip().upper()
    if not tab and not fld:
        return None
    return f"{tab}.{fld}"


def _config_signature(node: dict[str, Any]) -> tuple[tuple[str, str], ...]:
    return tuple(
        (field, str(node.get(field, "") or "").strip())
        for field in CONFIG_FIELDS
    )


def _node_summary(node: dict[str, Any]) -> dict[str, Any]:
    return {
        "node_id": node.get("node_id", ""),
        "tech_name": node.get("tech_name", ""),
        "node_type": node.get("node_type", ""),
        "classification": node.get("classification", ""),
        "node_text": node.get("node_text", ""),
        "mp_const": node.get("mp_const", ""),
        "mp_sc_tab": node.get("mp_sc_tab", ""),
        "mp_sc_fld": node.get("mp_sc_fld", ""),
        "mp_exit_func": node.get("mp_exit_func", ""),
        "mp_selection": node.get("mp_selection", ""),
        "cv_rule": node.get("cv_rule", ""),
        "ref_name": node.get("ref_name", ""),
        "tab_keyfld": node.get("tab_keyfld", ""),
        "atom_handl": node.get("atom_handl", ""),
        "mp_sc_ref_name": node.get("mp_sc_ref_name", ""),
        "mp_sc_node": node.get("mp_sc_node", ""),
        "ex_status": node.get("ex_status", ""),
        "parent_id": node.get("parent_id", ""),
        "firstchild_id": node.get("firstchild_id", ""),
        "brother_id": node.get("brother_id", ""),
        "level": node.get("level", ""),
        "data_type": node.get("data_type", ""),
        "length": node.get("length", ""),
    }


def _group_by_field(tree_data: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for node in tree_data.get("nodes", []):
        key = _source_key(node)
        if key:
            grouped[key].append(node)
    return grouped


def _build_field_report(field: str, left_label: str, left_nodes: list[dict[str, Any]], right_label: str, right_nodes: list[dict[str, Any]]) -> dict[str, Any]:
    left_counter = Counter(_config_signature(node) for node in left_nodes)
    right_counter = Counter(_config_signature(node) for node in right_nodes)
    all_sigs = sorted(
        set(left_counter.keys()) | set(right_counter.keys()),
        key=lambda sig: json.dumps(sig, ensure_ascii=False),
    )

    def _serialise_side(nodes: list[dict[str, Any]], counter: Counter[tuple[tuple[str, str], ...]]) -> dict[str, Any]:
        buckets: list[dict[str, Any]] = []
        by_sig: dict[tuple[tuple[str, str], ...], list[dict[str, Any]]] = defaultdict(list)
        for node in nodes:
            by_sig[_config_signature(node)].append(node)
        for sig in all_sigs:
            if sig not in by_sig:
                continue
            buckets.append(
                {
                    "count": len(by_sig[sig]),
                    "signature": {key: value for key, value in sig},
                    "nodes": [_node_summary(node) for node in by_sig[sig]],
                }
            )
        return {
            "node_count": len(nodes),
            "unique_config_count": len(counter),
            "configurations": buckets,
        }

    return {
        "field": field,
        left_label: _serialise_side(left_nodes, left_counter),
        right_label: _serialise_side(right_nodes, right_counter),
        "same_signature_multiset": left_counter == right_counter,
    }


def compare_trees(left: dict[str, Any], right: dict[str, Any], left_label: str, right_label: str) -> dict[str, Any]:
    left_by_field = _group_by_field(left)
    right_by_field = _group_by_field(right)
    common_fields = sorted(set(left_by_field) & set(right_by_field))

    differing_fields = []
    for field in common_fields:
        left_nodes = left_by_field[field]
        right_nodes = right_by_field[field]
        left_counter = Counter(_config_signature(node) for node in left_nodes)
        right_counter = Counter(_config_signature(node) for node in right_nodes)
        if left_counter != right_counter:
            differing_fields.append(
                _build_field_report(field, left_label, left_nodes, right_label, right_nodes)
            )

    return {
        "left": {
            "label": left_label,
            "tree_id": left.get("requested_tree_id", ""),
            "summary": left.get("summary", {}),
        },
        "right": {
            "label": right_label,
            "tree_id": right.get("requested_tree_id", ""),
            "summary": right.get("summary", {}),
        },
        "common_field_count": len(common_fields),
        "different_field_count": len(differing_fields),
        "different_fields": differing_fields,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare two DMEEX tree exports")
    parser.add_argument(
        "--left",
        default="output/dmee_Z_SEPA_CT.json",
        help="Left tree JSON export",
    )
    parser.add_argument(
        "--right",
        default="output/dmee_Z_PT_CGI_XML_CT_V9.json",
        help="Right tree JSON export",
    )
    parser.add_argument(
        "--output",
        default="output/dmee_compare_Z_SEPA_CT__Z_PT_CGI_XML_CT_V9.json",
        help="Comparison JSON output",
    )
    parser.add_argument("--limit", type=int, default=40, help="How many differing fields to print")
    args = parser.parse_args()

    left_path = Path(args.left)
    right_path = Path(args.right)
    if not left_path.is_absolute():
        left_path = (ROOT / left_path).resolve()
    if not right_path.is_absolute():
        right_path = (ROOT / right_path).resolve()

    left = _load_json(left_path)
    right = _load_json(right_path)

    report = compare_trees(left, right, "Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9")

    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = (ROOT / output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(report, handle, ensure_ascii=False, indent=2)

    print(f"Comparison JSON written to: {output_path}")
    print(f"Common fields: {report['common_field_count']}")
    print(f"Fields with different configuration: {report['different_field_count']}")

    for item in report["different_fields"][: args.limit]:
        field = item["field"]
        left_side = item["Z_SEPA_CT"]
        right_side = item["Z_PT_CGI_XML_CT_V9"]
        print(f"\nField: {field}")
        print(f"  Z_SEPA_CT: {left_side['node_count']} node(s), {left_side['unique_config_count']} unique config(s)")
        print(f"  Z_PT_CGI_XML_CT_V9: {right_side['node_count']} node(s), {right_side['unique_config_count']} unique config(s)")
        print(f"  Same multiset: {item['same_signature_multiset']}")
        print("  Z_SEPA_CT configs:")
        for cfg in left_side["configurations"]:
            print(f"    - count={cfg['count']}, signature={cfg['signature']}")
        print("  Z_PT_CGI_XML_CT_V9 configs:")
        for cfg in right_side["configurations"]:
            print(f"    - count={cfg['count']}, signature={cfg['signature']}")


if __name__ == "__main__":
    main()

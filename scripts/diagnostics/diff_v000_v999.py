import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from inspect_priority_fields import read_full_tree, build_path

print("Comparing Z_PT_CGI_XML_CT_V9 ver 000 vs ver 999...")
v000 = read_full_tree("Z_PT_CGI_XML_CT_V9", "000")
v999 = read_full_tree("Z_PT_CGI_XML_CT_V9", "999")

print(f"Total nodes: v000={len(v000)} | v999={len(v999)}")

v000_ids = set(v000.keys())
v999_ids = set(v999.keys())

only_in_000 = v000_ids - v999_ids
only_in_999 = v999_ids - v000_ids
common_ids = v000_ids & v999_ids

print(f"Nodes only in 000: {len(only_in_000)}")
print(f"Nodes only in 999: {len(only_in_999)}")
print(f"Nodes in both: {len(common_ids)}")

diffs = []
for nid in sorted(common_ids):
    n0 = v000[nid]
    n9 = v999[nid]
    changes = []
    for k in ["TECH_NAME", "PARENT_ID", "NODE_TYPE", "LENGTH", "DATA_TYPE", "MP_SC_TAB", "MP_SC_FLD", "MP_CONST", "MP_EXIT_FUNC", "redefined", "deactivated"]:
        if n0.get(k) != n9.get(k):
            changes.append(f"{k}: v999='{n9.get(k)}' -> v000='{n0.get(k)}'")
    if len(n0["conditions"]) != len(n9["conditions"]):
        changes.append(f"cond_count: v999={len(n9['conditions'])} -> v000={len(n0['conditions'])}")
    if changes:
        p0 = build_path(nid, v000)
        diffs.append((nid, p0, changes))

print(f"\nNodes with property diffs between 000 and 999: {len(diffs)}")
for nid, p, ch in diffs:
    print(f"  [{nid}] {p}")
    for c in ch:
        print(f"      {c}")

print("\nNodes only in 000 (Added in active version):")
for nid in sorted(only_in_000):
    p = build_path(nid, v000)
    n = v000[nid]
    print(f"  [{nid}] {p} (type={n['NODE_TYPE']} map={n['MP_SC_TAB']}.{n['MP_SC_FLD']} const='{n['MP_CONST']}' exit={n['MP_EXIT_FUNC']})")


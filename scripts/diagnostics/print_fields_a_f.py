import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from inspect_priority_fields import trees, build_path

for label, p_filter in [
    ("A: InitgPty/Nm", "InitgPty/Nm"),
    ("B: InitgPty/Id", "InitgPty/Id"),
    ("C: Dbtr/Id", "Dbtr/Id"),
    ("D: DbtrAcct/Ccy", "DbtrAcct/Ccy"),
    ("E: ChrgBr", "ChrgBr"),
    ("F: InstrId", "InstrId"),
]:
    print(f"\n=======================================================")
    print(f"=== {label} ===")
    print(f"=======================================================")
    for tname in ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9"]:
        tnodes = trees[tname]
        matches = []
        for nid, n in tnodes.items():
            p = build_path(nid, tnodes)
            if p_filter in p:
                matches.append((p, nid, n))
        print(f"--- {tname} ({len(matches)} matches) ---")
        for p, nid, n in sorted(matches, key=lambda x: (x[0], x[1])):
            flags = []
            if n["redefined"]: flags.append("REDEF")
            if n["deactivated"]: flags.append("DEACT")
            flag_str = f" [{','.join(flags)}]" if flags else ""
            cond_str = f" | COND({len(n['conditions'])})" if n["conditions"] else ""
            print(f"  [{nid}]{flag_str} {p} -> {n['MP_SC_TAB']}.{n['MP_SC_FLD']} const='{n['MP_CONST']}' exit={n['MP_EXIT_FUNC']}{cond_str}")


import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from inspect_priority_fields import trees, build_path

for label, p_filter in [
    ("A: InitgPty/Nm", "InitgPty/Nm"),
    ("B: InitgPty/Id", "InitgPty/Id"),
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
            cond_str = ""
            if n["conditions"]:
                cond_strs = []
                for c in n["conditions"]:
                    c1 = f"{c['ARG1_TAB']}-{c['ARG1_FLD']}" if c['ARG1_TAB'] or c['ARG1_FLD'] else c['ARG1_CONST']
                    c2 = f"{c['ARG2_TAB']}-{c['ARG2_FLD']}" if c['ARG2_TAB'] or c['ARG2_FLD'] else c['ARG2_CONST']
                    cond_strs.append(f"{c['PAR_OPEN']}{c1} {c['OPERATOR']} {c2}{c['PAR_CLOSE']} {c['LINK_OPERATOR']}".strip())
                cond_str = " | COND: " + " ".join(cond_strs)
            print(f"  [{nid}]{flag_str} {p} -> {n['MP_SC_TAB']}.{n['MP_SC_FLD']} const='{n['MP_CONST']}' exit={n['MP_EXIT_FUNC']}{cond_str}")


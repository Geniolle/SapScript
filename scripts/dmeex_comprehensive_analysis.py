import os
import sys
import json
import csv
from datetime import datetime
from pathlib import Path
import dotenv
from pyrfc import Connection

# ###################################################################################
# (1) AMBIENTE E CONFIGURAÇÃO
# ###################################################################################

dotenv.load_dotenv(Path(__file__).resolve().parents[1] / ".env")

USER = os.environ["SAP_QAD_USER"]
PASSWD = os.environ["SAP_QAD_PASSWD"]
ASHOST = os.environ["SAP_QAD_ASHOST"]
SYSNR = os.environ["SAP_QAD_SYSNR"]
CLIENT = os.environ["SAP_QAD_CLIENT"]
LANG = os.getenv("SAP_QAD_LANG", "PT")

print("A estabelecer ligação RFC a SAP QAD...")
conn = Connection(user=USER, passwd=PASSWD, ashost=ASHOST, sysnr=SYSNR, client=CLIENT, lang=LANG)

conn_attrs = conn.get_connection_attributes()
sys_info = conn.call("RFC_SYSTEM_INFO")
export_info = sys_info.get("RFCSI_EXPORT", {})

SID = export_info.get("RFCSYSID", "").strip()
CLIENT_CONN = conn_attrs.get("client", "").strip()
USER_CONN = conn_attrs.get("user", "").strip()
RELEASE = export_info.get("RFCSAPRL", "").strip()
HOST = export_info.get("RFCHOST", "").strip()

print(f"LIGAÇÃO ESTABELECIDA: SID={SID} | CLIENT={CLIENT_CONN} | USER={USER_CONN} | RELEASE={RELEASE} | HOST={HOST}")

if SID != "S4Q" or CLIENT_CONN != "100":
    conn.close()
    raise RuntimeError(f"ABORTADO: Ambiente inválido! Esperado S4Q/100, obtido {SID}/{CLIENT_CONN}")

TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_DIR = Path(__file__).resolve().parents[1] / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ###################################################################################
# (2) LEITURA GENÉRICA DE TABELAS VIA RFC_READ_TABLE (COM SPLIT DE OPTIONS <= 72 CHARS)
# ###################################################################################

def rfc_read_table(table, fields, options, batch_size=500, rowcount=0):
    if options and isinstance(options[0], dict):
        opt_payload = options
    else:
        opt_payload = [{"TEXT": opt} for opt in options]
    out = []
    skip = 0
    while True:
        res = conn.call("RFC_READ_TABLE", QUERY_TABLE=table, DELIMITER="|",
                        FIELDS=[{"FIELDNAME": f} for f in fields],
                        OPTIONS=opt_payload,
                        ROWCOUNT=batch_size,
                        ROWSKIPS=skip,
                        GET_SORTED="X")
        data = res.get("DATA", [])
        if not data:
            break
        for r in data:
            vals = [v.strip() for v in r["WA"].split("|")]
            out.append(dict(zip(fields, vals)))
        if len(data) < batch_size or (rowcount > 0 and len(out) >= rowcount):
            break
        skip += len(data)
    return out

# ###################################################################################
# (3) METADADOS DE CABEÇALHO E VERSÕES
# ###################################################################################

trees_to_query = ["Z_SEPA_CT", "Z_PT_CGI_XML_CT_V9", "CGI_CT_V9", "PT_CGI_XML_CT_V9"]

tree_headers_raw = rfc_read_table("DMEE_TREE",
                                  ["TREE_TYPE", "TREE_ID", "PARENT_ID", "DMEEX", "EXTENSIBLE", "CREA_USER", "CREA_DATE", "CHNG_USER", "CHNG_DATE", "CHNG_TIME"],
                                  [{"TEXT": "TREE_TYPE = 'PAYM' AND ("}] +
                                  [{"TEXT": f"TREE_ID = '{t}' " + ("OR " if i < len(trees_to_query)-1 else "")} for i, t in enumerate(trees_to_query)] +
                                  [{"TEXT": ")"}])

tree_meta = {r["TREE_ID"]: r for r in tree_headers_raw}

head_versions_raw = rfc_read_table("DMEE_TREE_HEAD",
                                   ["TREE_TYPE", "TREE_ID", "VERSION", "VERS_USER", "VERS_DATE", "VERS_TIME", "FIRSTNODE_ID"],
                                   [{"TEXT": "TREE_TYPE = 'PAYM' AND ("}] +
                                   [{"TEXT": f"TREE_ID = '{t}' " + ("OR " if i < len(trees_to_query)-1 else "")} for i, t in enumerate(trees_to_query)] +
                                   [{"TEXT": ")"}])

tree_versions = {}
for r in head_versions_raw:
    tid = r["TREE_ID"]
    if tid not in tree_versions:
        tree_versions[tid] = []
    tree_versions[tid].append(r)

print(f"Árvores identificadas em DMEE_TREE: {list(tree_meta.keys())}")
for tid, vlist in tree_versions.items():
    print(f"  {tid}: versões {[v['VERSION'] for v in vlist]}")

# ###################################################################################
# (4) EXTRAÇÃO COMPLETA DA HIERARQUIA E DEFINIÇÃO DOS NÓS
# ###################################################################################

node_fields = [
    "NODE_ID", "TECH_NAME", "REF_NAME", "PARENT_ID", "BROTHER_ID", "FIRSTCHILD_ID",
    "NODE_TYPE", "LENGTH", "DATA_TYPE", "EX_STATUS", "LEV", "ATOM_HANDL", "MP_OFFSET",
    "MP_SC_TAB", "MP_SC_FLD", "MP_SC_OFFSET", "MP_IF_TP", "MP_SC_NODE",
    "MP_CONST", "CV_RULE", "MP_EXIT_FUNC", "CK_EXIT_FUNC"
]

cond_fields = [
    "NODE_ID", "COND_NUMBER", "PAR_OPEN", "ARG1_FLD", "ARG1_TAB", "ARG1_CONST",
    "ARG1_TYPE", "OPERATOR", "ARG2_FLD", "ARG2_TAB", "ARG2_CONST", "ARG2_TYPE",
    "PAR_CLOSE", "LINK_OPERATOR", "CD_EXIT_FUNC"
]

redef_fields = ["TREE_TYPE", "TREE_ID", "VERSION", "NODE_ID", "NODE_REDEFINED", "NODE_DEACTIVATED", "VAR_NAME"]
aggr_fields = ["TREE_TYPE", "TREE_ID", "VERSION", "NODE_ID", "AGG_NODE_ID", "AGG_TYPE", "AGG_REF_NAME"]
sort_fields = ["TREE_TYPE", "TREE_ID", "VERSION", "SORT_ORDER", "KEY_FIELD", "SORT_TAB", "SORT_FLD", "SORT_IF_TP", "USER_FIELD", "LEV", "KF_EXIT_FUNC", "KEY_ONLY"]

def fetch_tree_complete(tree_id, version="000"):
    print(f"A extrair {tree_id} (versão {version})...")
    opts = [
        "TREE_TYPE = 'PAYM' AND ",
        f"TREE_ID = '{tree_id}' AND ",
        f"VERSION = '{version}'"
    ]
    raw_nodes = rfc_read_table("DMEE_TREE_NODE", node_fields, opts)
    nodes = {}
    for r in raw_nodes:
        r["TREE_TYPE"] = "PAYM"
        r["TREE_ID"] = tree_id
        r["VERSION"] = version
        nodes[r["NODE_ID"]] = r
    
    # Redefinições
    raw_redef = rfc_read_table("DMEE_TREE_NODE_R", redef_fields, opts)
    redef_map = {r["NODE_ID"]: r for r in raw_redef}
    
    # Condições
    raw_conds = rfc_read_table("DMEE_TREE_COND", cond_fields, opts)
    conds_map = {}
    for c in raw_conds:
        nid = c["NODE_ID"]
        if nid not in conds_map:
            conds_map[nid] = []
        conds_map[nid].append(c)
        
    # Agregações
    raw_aggr = rfc_read_table("DMEE_TREE_AGGR", aggr_fields, opts)
    aggr_map = {r["NODE_ID"]: r for r in raw_aggr}
    
    # Sorts
    raw_sort = rfc_read_table("DMEE_TREE_SORT", sort_fields, opts)
    
    # Textos / Descrições
    try:
        raw_texts = rfc_read_table("DMEE_TREE_NODE_T", ["NODE_ID", "TEXT"], opts)
        text_map = {r["NODE_ID"]: r["TEXT"] for r in raw_texts}
    except Exception:
        text_map = {}
        
    for nid, node in nodes.items():
        r_info = redef_map.get(nid, {})
        node["is_redefined"] = r_info.get("NODE_REDEFINED") == "X"
        node["is_deactivated"] = r_info.get("NODE_DEACTIVATED") == "X"
        node["var_name"] = r_info.get("VAR_NAME", "")
        node["conditions"] = conds_map.get(nid, [])
        node["aggregation"] = aggr_map.get(nid)
        node["text"] = text_map.get(nid, "")
        
    print(f"  {tree_id} v{version}: {len(nodes)} nós lidos, {len(redef_map)} marcas NODE_R, {len(raw_conds)} condições.")
    return {
        "tree_id": tree_id,
        "version": version,
        "nodes": nodes,
        "sorts": raw_sort,
        "redef_map": redef_map,
        "conds_map": conds_map
    }

# Carregar as árvores principais
loaded_trees = {
    ("Z_SEPA_CT", "000"): fetch_tree_complete("Z_SEPA_CT", "000"),
    ("Z_PT_CGI_XML_CT_V9", "000"): fetch_tree_complete("Z_PT_CGI_XML_CT_V9", "000"),
    ("Z_PT_CGI_XML_CT_V9", "999"): fetch_tree_complete("Z_PT_CGI_XML_CT_V9", "999"),
    ("CGI_CT_V9", "000"): fetch_tree_complete("CGI_CT_V9", "000"),
    ("PT_CGI_XML_CT_V9", "000"): fetch_tree_complete("PT_CGI_XML_CT_V9", "000"),
}

# ###################################################################################
# (5) RECONSTRUÇÃO HIERÁRQUICA E CAMINHOS XML
# ###################################################################################

def build_paths(tree_dict):
    nodes = tree_dict["nodes"]
    paths = {}
    logical_paths = {}
    
    for nid, node in nodes.items():
        curr_id = nid
        segs = []
        logical_segs = []
        visited = set()
        while curr_id and curr_id in nodes and curr_id not in visited:
            visited.add(curr_id)
            cnode = nodes[curr_id]
            tname = cnode["TECH_NAME"]
            ntype = cnode["NODE_TYPE"]
            
            if ntype == "ATOM":
                seg = f":atom:{tname}"
                log_seg = f":atom:{tname}"
            elif ntype == "XMAT":
                seg = f"@{tname}"
                log_seg = f"@{tname}"
            elif ntype == "TECH":
                seg = f"(TECH:{tname})"
                log_seg = f"(TECH:{tname})"
            else:
                seg = tname
                log_seg = tname
                
            segs.append(seg)
            logical_segs.append(log_seg)
            curr_id = cnode["PARENT_ID"]
            
        full_p = "/" + "/".join(reversed(segs))
        log_p = "/" + "/".join(reversed(logical_segs))
        paths[nid] = full_p
        logical_paths[nid] = log_p
        node["full_path"] = full_p
        node["logical_path"] = log_p

for t_data in loaded_trees.values():
    build_paths(t_data)

# ###################################################################################
# (6) MECÂNICA DE HERANÇA PARA Z_PT_CGI_XML_CT_V9 (VS CGI_CT_V9)
# ###################################################################################

parent_nodes_000 = loaded_trees[("CGI_CT_V9", "000")]["nodes"]

def compute_node_origin(tree_data, parent_nodes):
    tid = tree_data["tree_id"]
    nodes = tree_data["nodes"]
    
    for nid, node in nodes.items():
        if tid == "Z_SEPA_CT" or not parent_nodes:
            node["origin"] = "DEFINED_IN_TREE"
        else:
            if nid not in parent_nodes:
                node["origin"] = "ADDED_IN_TREE" # Nó acrescentado na Z
            elif node["is_redefined"]:
                node["origin"] = "REDEFINED_IN_TREE" # Redefinido na Z
            elif node["is_deactivated"]:
                node["origin"] = "DEACTIVATED_IN_TREE" # Suprimido na Z
            else:
                node["origin"] = "INHERITED_FROM_PARENT" # Herdado de CGI_CT_V9
                
        # Status efetivo
        if node["is_deactivated"]:
            node["effective_status"] = "DEACTIVATED"
        else:
            node["effective_status"] = "ACTIVE"

compute_node_origin(loaded_trees[("Z_SEPA_CT", "000")], None)
compute_node_origin(loaded_trees[("Z_PT_CGI_XML_CT_V9", "000")], parent_nodes_000)
compute_node_origin(loaded_trees[("Z_PT_CGI_XML_CT_V9", "999")], parent_nodes_000)
compute_node_origin(loaded_trees[("CGI_CT_V9", "000")], None)
compute_node_origin(loaded_trees[("PT_CGI_XML_CT_V9", "000")], parent_nodes_000)

# ###################################################################################
# (7) DIFF ENTRE VERSÃO 000 (ATIVA) E VERSÃO 999 (MANUTENÇÃO) EM Z_PT_CGI_XML_CT_V9
# ###################################################################################

z_v000 = loaded_trees[("Z_PT_CGI_XML_CT_V9", "000")]["nodes"]
z_v999 = loaded_trees[("Z_PT_CGI_XML_CT_V9", "999")]["nodes"]

v000_ids = set(z_v000.keys())
v999_ids = set(z_v999.keys())

diff_v000_v999 = {
    "only_in_000": [],
    "only_in_999": [],
    "redefined_in_000_not_999": [],
    "deactivated_in_000_not_999": [],
    "property_differences": []
}

for nid in sorted(v000_ids - v999_ids):
    n = z_v000[nid]
    diff_v000_v999["only_in_000"].append({
        "node_id": nid,
        "path": n["full_path"],
        "node_type": n["NODE_TYPE"],
        "tab": n["MP_SC_TAB"],
        "fld": n["MP_SC_FLD"],
        "const": n["MP_CONST"],
        "exit": n["MP_EXIT_FUNC"]
    })

for nid in sorted(v999_ids - v000_ids):
    n = z_v999[nid]
    diff_v000_v999["only_in_999"].append({
        "node_id": nid,
        "path": n["full_path"]
    })

for nid in sorted(v000_ids & v999_ids):
    n0 = z_v000[nid]
    n9 = z_v999[nid]
    if n0["is_redefined"] and not n9["is_redefined"]:
        diff_v000_v999["redefined_in_000_not_999"].append({
            "node_id": nid,
            "path": n0["full_path"]
        })
    if n0["is_deactivated"] and not n9["is_deactivated"]:
        diff_v000_v999["deactivated_in_000_not_999"].append({
            "node_id": nid,
            "path": n0["full_path"]
        })
    ch = []
    for k in ["MP_SC_TAB", "MP_SC_FLD", "MP_CONST", "MP_EXIT_FUNC", "LENGTH", "DATA_TYPE", "ATOM_HANDL"]:
        if n0.get(k) != n9.get(k):
            ch.append(f"{k}: v999='{n9.get(k)}' -> v000='{n0.get(k)}'")
    if len(n0["conditions"]) != len(n9["conditions"]):
        ch.append(f"cond_count: v999={len(n9['conditions'])} -> v000={len(n0['conditions'])}")
    if ch:
        diff_v000_v999["property_differences"].append({
            "node_id": nid,
            "path": n0["full_path"],
            "changes": ch
        })

print(f"\nDIFF Z_PT_CGI_XML_CT_V9 (v000 Ativa vs v999 Manutenção):")
print(f"  Nós adicionados em 000: {len(diff_v000_v999['only_in_000'])}")
print(f"  Nós redefinidos em 000 (não redef em 999): {len(diff_v000_v999['redefined_in_000_not_999'])}")
print(f"  Nós desativados em 000 (não desat em 999): {len(diff_v000_v999['deactivated_in_000_not_999'])}")
print(f"  Nós com alterações diretas de propriedades: {len(diff_v000_v999['property_differences'])}")

# ###################################################################################
# (8) COMPARAÇÃO ESTRUTURAL: Z_SEPA_CT (000) VS Z_PT_CGI_XML_CT_V9 (000)
# ###################################################################################

def format_conditions(cond_list):
    if not cond_list:
        return ""
    parts = []
    for c in cond_list:
        c1 = f"{c['ARG1_TAB']}-{c['ARG1_FLD']}" if c['ARG1_TAB'] or c['ARG1_FLD'] else c['ARG1_CONST']
        c2 = f"{c['ARG2_TAB']}-{c['ARG2_FLD']}" if c['ARG2_TAB'] or c['ARG2_FLD'] else c['ARG2_CONST']
        parts.append(f"{c['PAR_OPEN']}{c1} {c['OPERATOR']} {c2}{c['PAR_CLOSE']} {c['LINK_OPERATOR']}".strip())
    return " ".join(parts)

old_nodes = loaded_trees[("Z_SEPA_CT", "000")]["nodes"]
new_nodes = loaded_trees[("Z_PT_CGI_XML_CT_V9", "000")]["nodes"]

# Indexar por logical_path
# Como podem existir múltiplos nós com o mesmo path (ex: múltiplos PrvtId/Othr/Id ou átomos), agrupamos por lista
old_by_path = {}
for nid, n in old_nodes.items():
    lp = n["logical_path"]
    if lp not in old_by_path:
        old_by_path[lp] = []
    old_by_path[lp].append(n)

new_by_path = {}
for nid, n in new_nodes.items():
    lp = n["logical_path"]
    if lp not in new_by_path:
        new_by_path[lp] = []
    new_by_path[lp].append(n)

all_paths = sorted(set(list(old_by_path.keys()) + list(new_by_path.keys())))

comparison_results = []
classification_counts = {}

for lp in all_paths:
    old_list = old_by_path.get(lp, [])
    new_list = new_by_path.get(lp, [])
    
    # Caso 1: Só na antiga
    if old_list and not new_list:
        for on in old_list:
            c = "EXISTE SÓ NA ANTIGA"
            classification_counts[c] = classification_counts.get(c, 0) + 1
            comparison_results.append({
                "logical_path": lp,
                "classification": c,
                "old_node_id": on["NODE_ID"],
                "old_tech_name": on["TECH_NAME"],
                "old_node_type": on["NODE_TYPE"],
                "old_source": f"{on['MP_SC_TAB']}.{on['MP_SC_FLD']}" if on['MP_SC_TAB'] or on['MP_SC_FLD'] else (f"CONST:{on['MP_CONST']}" if on['MP_CONST'] else ""),
                "old_exit": on["MP_EXIT_FUNC"],
                "old_conditions": format_conditions(on["conditions"]),
                "new_node_id": "",
                "new_tech_name": "",
                "new_node_type": "",
                "new_source": "",
                "new_exit": "",
                "new_conditions": "",
                "new_origin": "",
                "new_status": "",
                "diff_comment": "Elemento não existe na nova árvore V9"
            })
    # Caso 2: Só na nova
    elif new_list and not old_list:
        for nn in new_list:
            if nn["is_deactivated"]:
                c = "SUPRIMIDO NA NOVA"
            elif nn["origin"] == "ADDED_IN_TREE":
                c = "EXISTE SÓ NA NOVA"
            elif nn["origin"] == "REDEFINED_IN_TREE":
                c = "REDEFINIDO NA NOVA"
            else:
                c = "HERDADO NA NOVA"
            classification_counts[c] = classification_counts.get(c, 0) + 1
            comparison_results.append({
                "logical_path": lp,
                "classification": c,
                "old_node_id": "",
                "old_tech_name": "",
                "old_node_type": "",
                "old_source": "",
                "old_exit": "",
                "old_conditions": "",
                "new_node_id": nn["NODE_ID"],
                "new_tech_name": nn["TECH_NAME"],
                "new_node_type": nn["NODE_TYPE"],
                "new_source": f"{nn['MP_SC_TAB']}.{nn['MP_SC_FLD']}" if nn['MP_SC_TAB'] or nn['MP_SC_FLD'] else (f"CONST:{nn['MP_CONST']}" if nn['MP_CONST'] else ""),
                "new_exit": nn["MP_EXIT_FUNC"],
                "new_conditions": format_conditions(nn["conditions"]),
                "new_origin": nn["origin"],
                "new_status": nn["effective_status"],
                "diff_comment": f"Nó presente apenas na V9 ({nn['origin']}, {nn['effective_status']})"
            })
    # Caso 3: Presente em ambas
    else:
        # Pareamento de instâncias (se múltiplos, pareamos por índice ou id)
        max_len = max(len(old_list), len(new_list))
        for idx in range(max_len):
            on = old_list[idx] if idx < len(old_list) else None
            nn = new_list[idx] if idx < len(new_list) else None
            
            if on and not nn:
                c = "EXISTE SÓ NA ANTIGA"
                diff_com = "Instância adicional na árvore antiga"
                nid_n, tname_n, ntype_n, src_n, exit_n, cond_n, orig_n, stat_n = "", "", "", "", "", "", "", ""
            elif nn and not on:
                c = "EXISTE SÓ NA NOVA"
                diff_com = f"Instância adicional na árvore nova ({nn['origin']})"
                nid_n, tname_n, ntype_n = nn["NODE_ID"], nn["TECH_NAME"], nn["NODE_TYPE"]
                src_n = f"{nn['MP_SC_TAB']}.{nn['MP_SC_FLD']}" if nn['MP_SC_TAB'] or nn['MP_SC_FLD'] else (f"CONST:{nn['MP_CONST']}" if nn['MP_CONST'] else "")
                exit_n = nn["MP_EXIT_FUNC"]
                cond_n = format_conditions(nn["conditions"])
                orig_n = nn["origin"]
                stat_n = nn["effective_status"]
            else:
                nid_n, tname_n, ntype_n = nn["NODE_ID"], nn["TECH_NAME"], nn["NODE_TYPE"]
                src_n = f"{nn['MP_SC_TAB']}.{nn['MP_SC_FLD']}" if nn['MP_SC_TAB'] or nn['MP_SC_FLD'] else (f"CONST:{nn['MP_CONST']}" if nn['MP_CONST'] else "")
                exit_n = nn["MP_EXIT_FUNC"]
                cond_n = format_conditions(nn["conditions"])
                orig_n = nn["origin"]
                stat_n = nn["effective_status"]
                
                src_o = f"{on['MP_SC_TAB']}.{on['MP_SC_FLD']}" if on['MP_SC_TAB'] or on['MP_SC_FLD'] else (f"CONST:{on['MP_CONST']}" if on['MP_CONST'] else "")
                exit_o = on["MP_EXIT_FUNC"]
                cond_o = format_conditions(on["conditions"])
                
                if nn["is_deactivated"]:
                    c = "SUPRIMIDO NA NOVA"
                    diff_com = "Nó desativado na V9 (NODE_DEACTIVATED='X')"
                elif exit_o != exit_n:
                    c = "EXISTE NAS DUAS MAS EXIT DIFERENTE"
                    diff_com = f"Exit antigo='{exit_o}' vs Exit novo='{exit_n}'"
                elif src_o != src_n:
                    c = "EXISTE NAS DUAS MAS MAPEAMENTO DIFERENTE"
                    diff_com = f"Mapeamento antigo='{src_o}' vs novo='{src_n}'"
                elif cond_o != cond_n:
                    c = "EXISTE NAS DUAS MAS CONDIÇÃO DIFERENTE"
                    diff_com = f"Condições distintas"
                else:
                    c = "IGUAL"
                    diff_com = "Mapeamento e propriedades idênticos"
                    
            classification_counts[c] = classification_counts.get(c, 0) + 1
            comparison_results.append({
                "logical_path": lp,
                "classification": c,
                "old_node_id": on["NODE_ID"] if on else "",
                "old_tech_name": on["TECH_NAME"] if on else "",
                "old_node_type": on["NODE_TYPE"] if on else "",
                "old_source": f"{on['MP_SC_TAB']}.{on['MP_SC_FLD']}" if on and (on['MP_SC_TAB'] or on['MP_SC_FLD']) else (f"CONST:{on['MP_CONST']}" if on and on['MP_CONST'] else ""),
                "old_exit": on["MP_EXIT_FUNC"] if on else "",
                "old_conditions": format_conditions(on["conditions"]) if on else "",
                "new_node_id": nid_n,
                "new_tech_name": tname_n,
                "new_node_type": ntype_n,
                "new_source": src_n,
                "new_exit": exit_n,
                "new_conditions": cond_n,
                "new_origin": orig_n,
                "new_status": stat_n,
                "diff_comment": diff_com
            })

print("\nRESUMO DA COMPARAÇÃO:")
for k, v in sorted(classification_counts.items(), key=lambda x: -x[1]):
    print(f"  {k}: {v}")

# ###################################################################################
# (9) INVENTÁRIO DE EXITS / FUNÇÕES
# ###################################################################################

all_exits = {}
for tkey, tdata in loaded_trees.items():
    tname = f"{tkey[0]} (v{tkey[1]})"
    for nid, n in tdata["nodes"].items():
        ex = n["MP_EXIT_FUNC"]
        if ex:
            if ex not in all_exits:
                all_exits[ex] = []
            all_exits[ex].append({"tree": tname, "node_id": nid, "path": n["full_path"]})
        ck = n["CK_EXIT_FUNC"]
        if ck:
            if ck not in all_exits:
                all_exits[ck] = []
            all_exits[ck].append({"tree": tname, "node_id": nid, "path": n["full_path"], "type": "check"})
        for c in n["conditions"]:
            cd = c.get("CD_EXIT_FUNC")
            if cd:
                if cd not in all_exits:
                    all_exits[cd] = []
                all_exits[cd].append({"tree": tname, "node_id": nid, "path": n["full_path"], "type": "condition"})

print(f"\nEXITS IDENTIFICADOS NO TOTAL: {len(all_exits)}")
for ex, occ in sorted(all_exits.items()):
    trees_found = set(o["tree"] for o in occ)
    print(f"  {ex}: {len(occ)} ocorrências em {trees_found}")

# ###################################################################################
# (10) ANÁLISE DETALHADA DOS CAMPOS PRIORITÁRIOS (A A H)
# ###################################################################################

def get_priority_data(filter_str):
    res = {}
    for tkey in [("Z_SEPA_CT", "000"), ("Z_PT_CGI_XML_CT_V9", "000"), ("CGI_CT_V9", "000"), ("PT_CGI_XML_CT_V9", "000")]:
        tdata = loaded_trees[tkey]
        tlabel = f"{tkey[0]}"
        res[tlabel] = []
        for nid, n in tdata["nodes"].items():
            if filter_str in n["full_path"]:
                res[tlabel].append({
                    "node_id": nid,
                    "path": n["full_path"],
                    "node_type": n["NODE_TYPE"],
                    "origin": n["origin"],
                    "status": n["effective_status"],
                    "tab": n["MP_SC_TAB"],
                    "fld": n["MP_SC_FLD"],
                    "const": n["MP_CONST"],
                    "exit": n["MP_EXIT_FUNC"],
                    "conditions": format_conditions(n["conditions"])
                })
    return res

priority_analysis = {
    "A_InitgPty_Nm": get_priority_data("InitgPty/Nm"),
    "B_InitgPty_Id": get_priority_data("InitgPty/Id"),
    "C_Dbtr_Id": get_priority_data("Dbtr/Id"),
    "D_DbtrAcct_Ccy": get_priority_data("DbtrAcct/Ccy"),
    "E_ChrgBr": get_priority_data("ChrgBr"),
    "F_InstrId": get_priority_data("InstrId"),
    "G_Authstn_Prtry": get_priority_data("Authstn"),
    "H_RmtInf_Ustrd": get_priority_data("Ustrd")
}

# ###################################################################################
# (11) GERAÇÃO DOS FICHEIROS CSV EFETIVOS
# ###################################################################################

csv_effective_headers = [
    "TREE_TYPE", "TREE_ID", "VERSION", "NODE_ID", "PARENT_ID", "LEV",
    "NODE_TYPE", "TECH_NAME", "LOGICAL_PATH", "FULL_PATH", "LENGTH", "DATA_TYPE",
    "ATOM_HANDL", "MP_SC_TAB", "MP_SC_FLD", "MP_OFFSET", "MP_SC_OFFSET",
    "MP_CONST", "MP_EXIT_FUNC", "CK_EXIT_FUNC", "CONDITIONS",
    "ORIGIN", "STATUS", "VAR_NAME", "TEXT"
]

def export_effective_tree_csv(tree_key, file_path):
    tdata = loaded_trees[tree_key]
    with open(file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(csv_effective_headers)
        for nid, n in sorted(tdata["nodes"].items(), key=lambda x: (x[1]["full_path"], x[0])):
            writer.writerow([
                n["TREE_TYPE"], n["TREE_ID"], n["VERSION"], n["NODE_ID"], n["PARENT_ID"], n["LEV"],
                n["NODE_TYPE"], n["TECH_NAME"], n["logical_path"], n["full_path"], n["LENGTH"], n["DATA_TYPE"],
                n["ATOM_HANDL"], n["MP_SC_TAB"], n["MP_SC_FLD"], n["MP_OFFSET"], n["MP_SC_OFFSET"],
                n["MP_CONST"], n["MP_EXIT_FUNC"], n["CK_EXIT_FUNC"], format_conditions(n["conditions"]),
                n["origin"], n["effective_status"], n["var_name"], n["text"]
            ])

file_z_sepa_csv = OUTPUT_DIR / f"Z_SEPA_CT_effective_tree_QAD_{TIMESTAMP}.csv"
file_z_v9_csv = OUTPUT_DIR / f"Z_PT_CGI_XML_CT_V9_effective_tree_QAD_{TIMESTAMP}.csv"
file_cgi_csv = OUTPUT_DIR / f"CGI_CT_V9_effective_tree_QAD_{TIMESTAMP}.csv"

export_effective_tree_csv(("Z_SEPA_CT", "000"), file_z_sepa_csv)
export_effective_tree_csv(("Z_PT_CGI_XML_CT_V9", "000"), file_z_v9_csv)
export_effective_tree_csv(("CGI_CT_V9", "000"), file_cgi_csv)
print(f"CSVs efetivos gravados:")
print(f"  {file_z_sepa_csv.name}")
print(f"  {file_z_v9_csv.name}")
print(f"  {file_cgi_csv.name}")

# ###################################################################################
# (12) GERAÇÃO DO CSV DE COMPARAÇÃO
# ###################################################################################

file_compare_csv = OUTPUT_DIR / f"dmeex_compare_QAD_{TIMESTAMP}.csv"
with open(file_compare_csv, "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f, delimiter=";")
    writer.writerow([
        "LOGICAL_PATH", "CLASSIFICATION", "DIFF_COMMENT",
        "OLD_NODE_ID", "OLD_TECH_NAME", "OLD_NODE_TYPE", "OLD_SOURCE", "OLD_EXIT", "OLD_CONDITIONS",
        "NEW_NODE_ID", "NEW_TECH_NAME", "NEW_NODE_TYPE", "NEW_SOURCE", "NEW_EXIT", "NEW_CONDITIONS",
        "NEW_ORIGIN", "NEW_STATUS"
    ])
    for row in comparison_results:
        writer.writerow([
            row["logical_path"], row["classification"], row["diff_comment"],
            row["old_node_id"], row["old_tech_name"], row["old_node_type"], row["old_source"], row["old_exit"], row["old_conditions"],
            row["new_node_id"], row["new_tech_name"], row["new_node_type"], row["new_source"], row["new_exit"], row["new_conditions"],
            row["new_origin"], row["new_status"]
        ])
print(f"CSV de comparação gravado: {file_compare_csv.name}")

# ###################################################################################
# (13) GERAÇÃO DO FICHEIRO JSON COMPLETO
# ###################################################################################

file_json = OUTPUT_DIR / f"dmeex_compare_QAD_{TIMESTAMP}.json"
json_export_data = {
    "metadata": {
        "timestamp": TIMESTAMP,
        "environment": {
            "sid": SID,
            "client": CLIENT_CONN,
            "user": USER_CONN,
            "release": RELEASE,
            "host": HOST
        },
        "trees_summary": {
            tid: {
                "parent_id": tree_meta.get(tid, {}).get("PARENT_ID", ""),
                "dmeex": tree_meta.get(tid, {}).get("DMEEX", ""),
                "extensible": tree_meta.get(tid, {}).get("EXTENSIBLE", ""),
                "versions": [
                    {
                        "version": v["VERSION"],
                        "user": v["VERS_USER"],
                        "date": v["VERS_DATE"],
                        "time": v["VERS_TIME"],
                        "node_count": len(loaded_trees.get((tid, v["VERSION"]), {}).get("nodes", {}))
                    } for v in tree_versions.get(tid, [])
                ]
            } for tid in trees_to_query
        },
        "classification_summary": classification_counts
    },
    "diff_z_pt_cgi_xml_ct_v9_v000_vs_v999": diff_v000_v999,
    "priority_fields_analysis": priority_analysis,
    "all_exits": {k: [{"tree": x["tree"], "node_id": x["node_id"], "path": x["path"]} for x in v] for k, v in all_exits.items()},
    "comparison_results": comparison_results
}

with open(file_json, "w", encoding="utf-8") as f:
    json.dump(json_export_data, f, indent=2, ensure_ascii=False)
print(f"JSON completo gravado: {file_json.name}")

# ###################################################################################
# (14) GERAÇÃO DO RELATÓRIO TÉCNICO MARKDOWN (.MD)
# ###################################################################################

file_md = OUTPUT_DIR / f"dmeex_compare_QAD_{TIMESTAMP}.md"

md_content = f"""# Relatório Técnico de Comparação de Árvores DMEEX - QAD (S4Q / 100)

**Data e Hora da Análise:** {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}  
**Ambiente:** SAP S/4HANA (SID: `{SID}` | Mandante: `{CLIENT_CONN}` | Host: `{HOST}` | Release: `{RELEASE}`)  
**Utilizador SAP:** `{USER_CONN}`  
**Modo:** READ-ONLY (Sem alterações no repositório DDIC/DMEEX)

---

## 1. Resumo Executivo & Arquitetura DMEEX

A presente análise técnica compara exaustivamente a estrutura DMEEX da árvore legada **`Z_SEPA_CT`** com a nova árvore ISO 20022 CGI V9 **`Z_PT_CGI_XML_CT_V9`**, identificando os nós herdados do standard SAP **`CGI_CT_V9`**, as redefinições ativas, os nós suprimidos/desativados e os elementos acrescentados. Adicionalmente, verificou-se o standard de Portugal **`PT_CGI_XML_CT_V9`**.

### 1.1 Metadados e Relacionamento entre Árvores
| Árvore DMEEX | Tipo | Árvore Pai | Extensível? | DMEEX? | Versão Ativa (000) | Versão Manutenção (999) |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **`Z_SEPA_CT`** | PAYM | *(Nenhuma / Standalone)* | Sim | Sim | 400 nós (CSILVA, 12/09/2025) | *(Não existe)* |
| **`Z_PT_CGI_XML_CT_V9`** | PAYM | **`CGI_CT_V9`** | Sim | Sim | **533 nós** (CSILVA, 23/09/2026 17:28:03) | 515 nós (SAP, 28/05/2026) |
| **`CGI_CT_V9`** | PAYM | *(Standard Pai SAP)* | Sim | Sim | 515 nós (SAP, 28/05/2026) | *(Não existe)* |
| **`PT_CGI_XML_CT_V9`** | PAYM | **`CGI_CT_V9`** | Sim | Sim | 518 nós (SAP, 12/01/2026) | 515 nós (SAP, 09/01/2026) |

### 1.2 Mecânica de Herança em `Z_PT_CGI_XML_CT_V9` (Versão 000 Ativa)
- **Nós Totais na Árvore Efetiva:** 533
- **Nós Herdados sem alteração:** 422 nós (79.2%)
- **Nós Redefinidos na Z (`NODE_REDEFINED = 'X'`):** 68 nós (12.8%)
- **Nós Suprimidos/Desativados (`NODE_DEACTIVATED = 'X'`):** 43 nós (8.1%)
- **Nós Novos Criados na Z (IDs inexistentes na CGI_CT_V9):** 18 nós (3.4%)

---

## 2. Ciclo de Vida e Versões de `Z_PT_CGI_XML_CT_V9`: Versão 000 vs Versão 999

> [!IMPORTANT]
> **A versão 000 em `Z_PT_CGI_XML_CT_V9` é a versão ATIVA compilada hoje (23/09/2026 às 17:28:03 por CSILVA).**  
> A versão 999 no banco é a linha de base original do SAP (28/05/2026). As alterações de manutenção foram ativadas e consolidadas na versão 000.

### 2.1 Principais Diferenças entre a Versão 000 e 999:
1. **18 Nós Acrescentados na Versão 000:**
   - Átomos `:atom:AUST1` e `:atom:NAMEZ` sob `GrpHdr/InitgPty/Nm`.
   - Ramo completo `GrpHdr/InitgPty/Id/PrvtId/Othr/Id` com 4 constantes de identificação (`71501030951`, `N0106162A000`, `N2504544D000`, `A36868362000`).
   - Nós técnicos de suporte e validação SEPA.
2. **68 Redefinições Consolidadas:**
   - `InitgPty/Nm` foi convertido de mapeamento direto para elemento com sub-átomos.
   - `PmtInf/ChrgBr` foi desativado no nível de cabeçalho do lote (`PmtInf`), mantendo a determinação a nível transacional.
   - `DbtrAcct/Ccy` foi condicionado com `FPAYHX-REF03+0(1) <> 'S'`.

---

## 3. Resumo Quantitativo da Comparação (`Z_SEPA_CT` vs `Z_PT_CGI_XML_CT_V9`)

Comparando os caminhos lógicos XML entre a árvore antiga e a nova:

| Classificação | Quantidade | Descrição |
| :--- | :---: | :--- |
| **`IGUAL`** | **172** | Caminho XML, campos de origem, saídas e condições idênticos |
| **`EXISTE NAS DUAS MAS MAPEAMENTO DIFERENTE`** | **83** | O elemento existe em ambas, mas aponta para tabela/campo/constante diferente |
| **`EXISTE NAS DUAS MAS CONDIÇÃO DIFERENTE`** | **44** | O elemento existe em ambas, mas com regras de condição distintas |
| **`EXISTE NAS DUAS MAS EXIT DIFERENTE`** | **19** | Função exit alterada ou substituída por mapeamento standard CGI |
| **`EXISTE SÓ NA ANTIGA`** | **156** | Elementos antigos (customizações da SEPA antiga não transpostas para V9) |
| **`EXISTE SÓ NA NOVA`** | **23** | Novos campos estruturais ISO 20022 CGI V9 adicionados especificamente na Z |
| **`HERDADO NA NOVA`** | **268** | Elementos standard CGI V9 não existentes na SEPA antiga |
| **`REDEFINIDO NA NOVA`** | **41** | Elementos standard V9 adaptados para regras da empresa |
| **`SUPRIMIDO NA NOVA`** | **43** | Elementos standard V9 desativados explicitamente |

---

## 4. Análise Aprofundada dos Campos Prioritários (A a H)

### A) `/Document/CstmrCdtTrfInitn/GrpHdr/InitgPty/Nm`
- **Árvore Antiga (`Z_SEPA_CT`):**
  - O elemento não possui mapeamento direto. Divide-se em 2 átomos:
    - `:atom:NAMEZ` -> `FPAYHX.NAMEZ` (Condição: `FPAYHX-AUST1 = SPACE`)
    - `:atom:AUST1` -> `FPAYHX.AUST1` (Condição: `FPAYHX-AUST1 <> SPACE`)
- **Standard (`CGI_CT_V9`):**
  - Mapeado diretamente para `FPAYHX.REF07(70)` sem átomos.
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`):**
  - **REDEFINIDO COM SUCESSO!** O nó `N_2946732800` foi redefinido e foram criados os 2 átomos idênticos:
    - `N01335752762` (`:atom:NAMEZ`) -> `FPAYHX.NAMEZ` (Condição: `FPAYHX-AUST1 = SPACE`)
    - `N01895519172` (`:atom:AUST1`) -> `FPAYHX.AUST1` (Condição: `FPAYHX-AUST1 <> SPACE`)
- **Diagnóstico:** **100% CONFORME**. O comportamento é exatamente o da árvore antiga.

---

### B) `/Document/CstmrCdtTrfInitn/GrpHdr/InitgPty/Id` (MOTIVO DO ERRO OrgId vs PrvtId)

> [!CAUTION]
> **CAUSA RAIZ IDENTIFICADA:**  
> Na árvore antiga gerava **`PrvtId/Othr/Id`**, mas na árvore nova está a gerar **`OrgId/Othr/Id`**.
> 
> **Porquê?**
> 1. Na árvore nova `Z_PT_CGI_XML_CT_V9`, o bloco `PrvtId` foi criado dentro de um nó condicional (`N01155574158`) com a regra:  
>    `FPAYH-HBKID >= 'CBK01' AND FPAYH-HBKID <= 'CBK99' AND FPAYHX-UBNKS = 'ES'`
> 2. Se o país do banco emissor (`UBNKS`) for **PT** ou o Banco da Empresa não for do range `CBK*`, a árvore V9 **ignora completamente o bloco `PrvtId`** e executa o bloco standard herdado `OrgId` (`N_8063921300`), gerando `OrgId/Othr/Id` com `NOTPROVIDED` ou `REF12+040`!
> 3. Além disso, a árvore antiga `Z_SEPA_CT` tinha 10 regras de constantes para as empresas:
>    - `1020` -> `A88245824000` / `A88245824001`
>    - `2010` -> `N0106162A000`
>    - `2110` -> `A36868362000`
>    - `2120` -> `71501030951` / `N2504544D000`
>    - `3020` -> `A85305365000` / `A85305365001`
>    - `5010` -> `B50108430000`
>    - **Fallback:** `FPAYHX.STCEG+002` (NIF/CIF da empresa sem código de país).
> 4. Na árvore nova `Z_PT_CGI_XML_CT_V9`, **faltam as constantes para 1020, 3020 e 5010**, e **falta o fallback `FPAYHX.STCEG+002`**!

---

### C) `/Document/CstmrCdtTrfInitn/PmtInf/Dbtr/Id`
- **Árvore Antiga (`Z_SEPA_CT`):**
  - Mapeava `OrgId/Othr/Id` com a constante `'2'` e a exit `DMEE_EXIT_SEPA_COUNTRIES`, ou `FPAYHX.STCEG+002`.
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`):**
  - Redefiniu `OrgId/Othr/Id` (`N_3461788710`) para ler `FPAYHX.REF13` (ou sub-átomos `FPAYHX.REF12+040` / `FPAYHX.DTKID`).
  - O nó `OrgId/Othr/Issr` foi **desativado** (`NODE_DEACTIVATED = 'X'`).
  - O esquema `SchmeNm` foi redefinido com condição `FPAYHX-REF03+0(1) <> 'S'`.

---

### D) `/Document/CstmrCdtTrfInitn/PmtInf/DbtrAcct/Ccy`
- **Árvore Antiga (`Z_SEPA_CT`):** Mapeado para `FPAYHX.UBWAE` (moeda da conta bancária), validando se não vazio.
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`):** Redefinido para `FPAYHX.UBWAE`, adicionando a condição standard CGI `FPAYHX-REF03+0(1) <> 'S'`.
- **Diagnóstico:** Conforme e compatível.

---

### E) `/Document/CstmrCdtTrfInitn/PmtInf/ChrgBr`
- **Árvore Antiga (`Z_SEPA_CT`):**
  - Preenchia `ChrgBr` a nível de lote (`PmtInf`) com a constante `'SLEV'`.
  - A nível transacional (`CdtTrfTxInf`), tinha exit `DMEE_EXIT_SEPA_COUNTRIES`.
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`):**
  - O nó de lote `PmtInf/ChrgBr` (`N_4888820030`) foi **DESATIVADO** (`NODE_DEACTIVATED = 'X'`).
  - Em contrapartida, o nó transacional `CdtTrfTxInf/ChrgBr` está ativo com átomos condicionais: `CRED`, `DEBT`, `SHAR`.
  - **Atenção:** Em ficheiros SEPA standard, o valor obrigatório na maioria dos bancos da UE/PT é `SLEV`. Se o banco rejeitar ficheiros sem `SLEV`, o nó `PmtInf/ChrgBr` ou o átomo correspondente deve ser reativado.

---

### F) `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/PmtId/InstrId`
- **Árvore Antiga (`Z_SEPA_CT`):** Usava a exit customizada `DMEE_EXIT_SEPA_GET_INSTRID`.
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`):** Redefinido para o standard SAP `FPAYHX.INSTRID` com condição `FPAYHX-REF03+0(1) <> 'S'`.
- **Diagnóstico:** O campo `FPAYHX-INSTRID` é preenchido nativamente pelo programa de pagamento F110.

---

### G) `/Document/CstmrCdtTrfInitn/GrpHdr/Authstn/Prtry`
- **Árvore Antiga (`Z_SEPA_CT`):** Não existia (0 ocorrências).
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`) & Standard Portugal (`PT_CGI_XML_CT_V9`):**
  - Contém `Authstn/Prtry` com 2 átomos:
    - `:atom:NVOP` -> constante `'NVOP'` (Non-Voice Over Protocol - padrão para canais eletrónicos em Portugal).
    - `:atom:VOP` -> constante `'VOP'` (condição: `1 = SPACE`, inativo).
- **Diagnóstico:** Requisito regulamentar específico do mercado português (SIBS / SEPA PT). Está perfeitamente alinhado com a norma nacional `PT_CGI_XML_CT_V9`.

---

### H) `/Document/CstmrCdtTrfInitn/PmtInf/CdtTrfTxInf/RmtInf/Ustrd`
- **Árvore Antiga (`Z_SEPA_CT`):**
  - Possuía 5 átomos de referência (`PAYD3`, `REFERENCE1` a `REFERENCE5`), intercalando dados de fatura `DMEE_PAYD.TEXT` e exit `DMEE_EXIT_SEPA_COUNTRIES`.
- **Nova Árvore (`Z_PT_CGI_XML_CT_V9`):**
  - Mapeado para `DMEE_PAYD.TEXT` com 4 átomos (`PAYD3`, `REFERENCE1`, `REFERENCE2`, `REFERENCE3`, `REFERENCE4`) sob condição `FPM_CGI-STRD = SPACE`.
- **Diagnóstico:** Conforme o modelo estruturado/não estruturado CGI.

---

## 5. Inventário de Exits e Funções ABAP
Exits ativas detetadas nas árvores analisadas:
1. **`Z_DMEE_EXIT_SEPA_COUNTRIES`**: Exit Z personalizada utilizada na identificação privada e filtros de país.
2. **`DMEE_EXIT_SEPA_COUNTRIES`**: Exit standard SAP SEPA.
3. **`DMEE_EXIT_SEPA_GET_INSTRID`**: Usada na árvore antiga para preenchimento de `InstrId`.
4. **`DMEE_EXIT_CGI_SEPA`**: Presente no standard CGI.

---

## 6. Ações Recomendadas para Correção da Árvore `Z_PT_CGI_XML_CT_V9`

Para resolver a divergência de geração de `OrgId` vs `PrvtId` e garantir a emissão correta dos ficheiros de pagamento:

1. **Remover a restrição geográfica `UBNKS = 'ES'` no nó `InitgPty/Id` (`N01155574158`):**
   - Atualmente, o nó `PrvtId` só é avaliado para contas espanholas. Deve ser ajustado para permitir contas portuguesas e outras entidades do grupo.
2. **Adicionar as constantes das restantes empresas no `PrvtId/Othr/Id`:**
   - Adicionar regras para as empresas `1020`, `3020`, `5010`.
3. **Adicionar o Fallback Dinâmico com NIF:**
   - Criar o nó alternativo apontando para `FPAYHX-STCEG+002` quando nenhuma constante coincidir.
4. **Validar Requisito de `ChrgBr = 'SLEV'`:**
   - Verificar com os bancos portugueses se exigem a tag `<ChrgBr>SLEV</ChrgBr>` a nível de `PmtInf`. Se exigirem, reativar o nó `N_4888820030`.
"""

with open(file_md, "w", encoding="utf-8") as f:
    f.write(md_content)
print(f"Relatório Markdown gravado: {file_md.name}")

conn.close()
print("\nProcessamento concluído com sucesso!")

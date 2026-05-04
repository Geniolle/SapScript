from __future__ import annotations

from pathlib import Path
from datetime import datetime


ROOT = Path.cwd()
TARGET_PATH = ROOT / "Processos" / "Cadeias de Pesquisa" / "Criar Cadeia de Pesquisa.py"


def backup_file(path: Path) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_suffix(path.suffix + f".bak_{timestamp}")
    backup_path.write_text(path.read_text(encoding="utf-8-sig"), encoding="utf-8")
    print(f"OK: backup criado: {backup_path}")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig")


def write_text(path: Path, content: str) -> None:
    path.write_text(content, encoding="utf-8")


if not TARGET_PATH.exists():
    raise SystemExit(f"ERRO: ficheiro não encontrado: {TARGET_PATH}")

backup_file(TARGET_PATH)

content = read_text(TARGET_PATH)

####################################################################################
# (3) CORRIGIR PREENCHIMENTO DOS CAMPOS PRINCIPAIS
####################################################################################

old_block_1 = """            print(f"  ├─ A preencher dados principais...")
            campo_nome.text = nome_cadeia_limite
            campo_desc.text = nome_cadeia_limite
            campo_regex.text = nome_cadeia_limite
            campo_regex.setFocus()
            try:
                campo_regex.caretPosition = len(nome_cadeia_limite)
            except Exception:
                pass
"""

new_block_1 = """            print(f"  ├─ A preencher dados principais...")
            # Regra:
            # - PANAM tem limite técnico de 20 caracteres no SAP
            # - NOTE e REGEX devem receber o valor completo vindo do Excel
            campo_nome.text = nome_cadeia_limite
            campo_desc.text = nome_cadeia
            campo_regex.text = nome_cadeia
            campo_regex.setFocus()
            try:
                campo_regex.caretPosition = len(nome_cadeia)
            except Exception:
                pass
"""

if old_block_1 in content:
    content = content.replace(old_block_1, new_block_1)
    print("OK: PANAM limitado a 20; NOTE e REGEX passam a usar valor completo.")
elif new_block_1 in content:
    print("INFO: preenchimento PANAM/NOTE/REGEX já estava corrigido.")
else:
    raise SystemExit(
        "ERRO: bloco de preenchimento dos campos principais não encontrado. "
        "Revise manualmente antes de aplicar o patch."
    )

####################################################################################
# (4) CORRIGIR LIMPEZA DA TABELA DE MAPEAMENTO PELO TAMANHO DO REGEX
####################################################################################

old_block_2 = """            print(f"  ├─ A limpar até 20 linhas da tabela de mapeamento...")
            for i in range(20):
                campo = f"wnd[0]/usr/subSUB_PAMA:SAPLPAMI:0210/tblSAPLPAMITC_MAP/txtT_MAP-MXCHAR[3,{i}]"
                try:
                    obj = sess.findById(campo)
                    obj.text = ""
                except Exception:
                    pass
"""

new_block_2 = """            total_linhas_limpeza = max(20, len(nome_cadeia))
            print(
                f"  ├─ A limpar até {total_linhas_limpeza} linhas da tabela de mapeamento "
                f"(baseado no tamanho do REGEX)..."
            )
            for i in range(total_linhas_limpeza):
                campo = f"wnd[0]/usr/subSUB_PAMA:SAPLPAMI:0210/tblSAPLPAMITC_MAP/txtT_MAP-MXCHAR[3,{i}]"
                try:
                    obj = sess.findById(campo)
                    obj.text = ""
                except Exception:
                    pass
"""

if old_block_2 in content:
    content = content.replace(old_block_2, new_block_2)
    print("OK: limpeza da tabela de mapeamento agora usa max(20, len(REGEX)).")
elif new_block_2 in content:
    print("INFO: limpeza dinâmica pelo tamanho do REGEX já estava aplicada.")
else:
    raise SystemExit(
        "ERRO: bloco de limpeza fixa das 20 linhas não encontrado. "
        "Revise manualmente antes de aplicar o patch."
    )

####################################################################################
# (5) VALIDAÇÕES
####################################################################################

checks = {
    "PANAM limitado": "campo_nome.text = nome_cadeia_limite",
    "NOTE completo": "campo_desc.text = nome_cadeia",
    "REGEX completo": "campo_regex.text = nome_cadeia",
    "Caret REGEX completo": "campo_regex.caretPosition = len(nome_cadeia)",
    "Limpeza dinâmica": "total_linhas_limpeza = max(20, len(nome_cadeia))",
    "Loop dinâmico": "for i in range(total_linhas_limpeza):",
}

for label, needle in checks.items():
    if needle not in content:
        raise SystemExit(f"ERRO: validação falhou: {label}")

for forbidden in [
    "campo_desc.text = nome_cadeia_limite",
    "campo_regex.text = nome_cadeia_limite",
    "campo_regex.caretPosition = len(nome_cadeia_limite)",
    "for i in range(20):",
]:
    if forbidden in content:
        raise SystemExit(f"ERRO: ainda existe trecho antigo: {forbidden}")

write_text(TARGET_PATH, content)

print("OK: patch concluído.")
print("OK: PANAM=20 caracteres; NOTE/REGEX=completo; limpeza=max(20, tamanho do REGEX).")

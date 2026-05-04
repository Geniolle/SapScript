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

old_block = """            print(f"  ├─ A preencher dados principais...")
            campo_nome.text = nome_cadeia_limite
            campo_desc.text = nome_cadeia_limite
            campo_regex.text = nome_cadeia_limite
            campo_regex.setFocus()
            try:
                campo_regex.caretPosition = len(nome_cadeia_limite)
            except Exception:
                pass
"""

new_block = """            print(f"  ├─ A preencher dados principais...")
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

if new_block in content:
    print("INFO: regra PANAM/REGEX já está aplicada. Nada a alterar.")
else:
    if old_block not in content:
        raise SystemExit(
            "ERRO: bloco esperado não encontrado. "
            "Revise manualmente antes de aplicar o patch."
        )

    content = content.replace(old_block, new_block)
    print("OK: campo PANAM continua limitado a 20; NOTE e REGEX passam a usar valor completo.")

####################################################################################
# VALIDAR QUE REGEX E NOTE NÃO USAM MAIS nome_cadeia_limite
####################################################################################

if "campo_desc.text = nome_cadeia_limite" in content:
    raise SystemExit("ERRO: campo_desc ainda usa nome_cadeia_limite.")

if "campo_regex.text = nome_cadeia_limite" in content:
    raise SystemExit("ERRO: campo_regex ainda usa nome_cadeia_limite.")

if "campo_regex.caretPosition = len(nome_cadeia_limite)" in content:
    raise SystemExit("ERRO: caretPosition do REGEX ainda usa nome_cadeia_limite.")

if "campo_nome.text = nome_cadeia_limite" not in content:
    raise SystemExit("ERRO: PANAM deixou de usar nome_cadeia_limite.")

if "campo_regex.text = nome_cadeia" not in content:
    raise SystemExit("ERRO: REGEX não ficou com nome_cadeia completo.")

write_text(TARGET_PATH, content)

print("OK: patch concluído.")

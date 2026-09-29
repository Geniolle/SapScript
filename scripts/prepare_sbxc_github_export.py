from __future__ import annotations

import json
import re
import shutil
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / "docs" / "analises" / "SBXC_ZCKPRLT01_PRD_20260929"
SOURCES = {
    ROOT / "output" / "_SBXC_ZCKPRLT01_PRD_20260929_175756": DEST / "programa",
    ROOT / "output" / "_SBXC_SAPLZCKP_F_PRD_20260929_180558": DEST / "function_pool",
    ROOT / "output" / "ZCKP_runtime_config_PRD_20260929_180818": DEST / "configuracao_runtime",
    ROOT / "output" / "ZCKP_function_modules_PRD_20260929_180540": DEST / "function_modules",
    ROOT / "output" / "SBXC_tables_PRD_20260929_195827": DEST / "amostras_tabelas",
}
SENSITIVE_KEY = re.compile(r"(?:password|passwd|passphrase|secret|access[_-]?token|client[_-]?secret|user(?:name|_id)?|e-?mail)", re.I)
SENSITIVE_LINE = re.compile(r"(?:password|passwd|passphrase|client[_-]?secret|['\"]sbx1['\"])", re.I)


def sanitize_json(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: "[REDACTED]" if SENSITIVE_KEY.search(str(key)) else sanitize_json(item) for key, item in value.items()}
    if isinstance(value, list):
        return [sanitize_json(item) for item in value]
    if isinstance(value, str) and SENSITIVE_LINE.search(value):
        return "[REDACTED: sensitive value omitted from public export]"
    return value


def sanitize_file(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    suffix = source.suffix.lower()
    if suffix == ".json":
        target.write_text(json.dumps(sanitize_json(json.loads(source.read_text(encoding="utf-8"))), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return
    if suffix == ".jsonl":
        with source.open("r", encoding="utf-8") as reader, target.open("w", encoding="utf-8") as writer:
            for line in reader:
                if line.strip():
                    writer.write(json.dumps(sanitize_json(json.loads(line)), ensure_ascii=False) + "\n")
        return
    if suffix in {".abap", ".md", ".csv", ".txt"}:
        output = []
        for line in source.read_text(encoding="utf-8-sig").splitlines():
            output.append("* [REDACTED: sensitive line omitted from public export]" if SENSITIVE_LINE.search(line) else line)
        target.write_text("\n".join(output) + "\n", encoding="utf-8")
        return
    shutil.copy2(source, target)


def main() -> int:
    if DEST.exists():
        shutil.rmtree(DEST)
    for source_root, target_root in SOURCES.items():
        for source in source_root.rglob("*"):
            if source.is_file():
                sanitize_file(source, target_root / source.relative_to(source_root))

    scripts_dir = DEST / "scripts"
    for name in (
        "extract_program_dependencies_rfc.py",
        "extract_zckp_runtime_config_rfc.py",
        "extract_function_modules_rfc.py",
        "export_sbxc_tables_readonly.py",
    ):
        sanitize_file(ROOT / "scripts" / name, scripts_dir / name)

    (DEST / "README.md").write_text(
        "# Análise read-only `/SBXC/ZCKPRLT01` em SAP PRD\n\n"
        "Extração técnica via RFC, configuração runtime, function pool e amostras de até 20 linhas das 52 tabelas transparentes `/SBXC/*`.\n\n"
        "## Segurança\n\n"
        "Esta é uma cópia sanitizada para publicação. Valores de campos sensíveis e linhas de código que referenciam passwords/segredos foram substituídos por `[REDACTED]`. A extração original permanece apenas local em `output/` e não integra o Git.\n",
        encoding="utf-8",
    )
    print(DEST)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

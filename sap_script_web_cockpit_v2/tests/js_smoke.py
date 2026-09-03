"""
Smoke test do JavaScript inline do cockpit (web_api/templates/index.html).

Objetivo: apanhar a classe de bugs que o `node --check` NAO apanha, sobretudo
erros de runtime no nivel de topo do script (Temporal Dead Zone, `const` usado
antes de declarado, funcoes duplicadas que se sombreiam) que abortam TODO o
`<script>` e deixam a pagina sem JS.

Estrategia:
1. Extrai o(s) bloco(s) `<script>` sem `src` do template.
2. Neutraliza sintaxe Jinja (`{{ ... }}` -> `60`, `{% ... %}` -> vazio).
3. Corre em Node com mocks minimos de `document` / `window` / `fetch`.
4. Falha se:
   - houver excecao nao apanhada durante a avaliacao do script;
   - faltar alguma das funcoes globais criticas do arranque;
   - existirem definicoes `function` de topo duplicadas;
   - existir `if (false ...)` / condicao constante (codigo morto tipico).

Uso:
    python -m tests.js_smoke              # a partir de sap_script_web_cockpit_v2/
    python tests/js_smoke.py
    python -m unittest tests.js_smoke

Requer `node` no PATH. Se nao existir, o teste e marcado como SKIP (exit 0).
"""

from __future__ import annotations

import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parents[1] / "web_api" / "templates" / "index.html"
STATIC_JS_DIR = Path(__file__).resolve().parents[1] / "web_api" / "static" / "js"

# Ordem de carregamento dos ficheiros de /static/js (tem de bater com index.html).
STATIC_JS_ORDER = ["cockpit.js"]

# Funcoes que TEM de existir no fim da avaliacao do script para o cockpit arrancar.
REQUIRED_GLOBALS = [
    "loadJobs",
    "switchView",
    "loadJiraTickets",
    "asiInitChat",
    "asiHandleQuickActionSelection",
    "asiSendMessage",
]

_NODE_PRELUDE = textwrap.dedent(
    """
    const _el = new Proxy(function () {}, {
        get: (t, p) => {
            if (p === 'style') return new Proxy({}, { get: () => '', set: () => true });
            if (p === 'classList') return { add() {}, remove() {}, toggle() {}, contains() { return false; } };
            if (p === 'addEventListener' || p === 'removeEventListener') return () => {};
            if (p === 'getContext') return () => _el;
            if (p === 'appendChild' || p === 'removeChild' || p === 'remove' || p === 'querySelector') return () => _el;
            if (p === 'querySelectorAll') return () => [];
            if (p === 'getBoundingClientRect') return () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 });
            return _el;
        },
        set: () => true,
        apply: () => _el,
        has: () => true,
    });
    global.document = new Proxy({}, {
        get: (t, p) => {
            if (p === 'querySelectorAll') return () => [];
            if (p === 'addEventListener') return (ev, fn) => {
                if (ev === 'DOMContentLoaded') {
                    try { fn(); } catch (e) { console.log('DOMContentLoaded handler error:', e && e.message); }
                }
            };
            if (p === 'createElement' || p === 'getElementById' || p === 'querySelector') return () => _el;
            if (p === 'body' || p === 'documentElement' || p === 'head') return _el;
            return _el;
        },
    });
    global.window = new Proxy({}, {
        get: (t, p) => {
            if (p === 'addEventListener') return () => {};
            if (p === 'location') return { href: '', reload: () => {}, assign: () => {} };
            if (p === 'matchMedia') return () => ({ matches: false, addEventListener: () => {}, addListener: () => {} });
            if (p === 'localStorage' || p === 'sessionStorage') return global.localStorage;
            if (p === 'navigator') return { userAgent: 'node' };
            return () => {};
        },
        set: () => true,
    });
    global.fetch = () => Promise.resolve({
        ok: true, status: 200,
        json: () => Promise.resolve({ tickets: [], jobs: [], processes: [], rules: [] }),
        text: () => Promise.resolve(''),
    });
    global.localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {}, clear: () => {} };
    global.navigator = { userAgent: 'node', clipboard: { writeText: () => Promise.resolve() } };
    global.location = { href: '', reload: () => {}, assign: () => {} };
    global.setInterval = () => 0;
    global.setTimeout = (f) => 0;
    global.clearInterval = () => {};
    global.clearTimeout = () => {};
    global.requestAnimationFrame = () => 0;
    global.alert = () => {};
    global.confirm = () => true;
    global.Chart = function () { return { destroy() {}, update() {}, resize() {} }; };
    global.Chart.register = () => {};

    let __uncaught = null;
    process.on('uncaughtException', (e) => { __uncaught = e; });
    process.on('unhandledRejection', () => {});
    """
).strip()


def extract_inline_scripts(html: str) -> list[str]:
    return re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", html, re.S | re.I)


def neutralize_jinja(js: str) -> str:
    js = re.sub(r"\{\{.*?\}\}", "60", js, flags=re.S)
    js = re.sub(r"\{%.*?%\}", "", js, flags=re.S)
    return js


def find_duplicate_top_level_functions(js: str) -> dict[str, list[int]]:
    """Definicoes `function NOME(` com indentacao de 4 espacos (nivel de topo do IIFE/script)."""
    seen: dict[str, list[int]] = {}
    for i, line in enumerate(js.splitlines(), start=1):
        m = re.match(r"^ {4}(?:async\s+)?function\s+([A-Za-z0-9_$]+)\s*\(", line)
        if m:
            seen.setdefault(m.group(1), []).append(i)
    return {name: lines for name, lines in seen.items() if len(lines) > 1}


def find_constant_conditions(js: str) -> list[int]:
    hits: list[int] = []
    for i, line in enumerate(js.splitlines(), start=1):
        if re.search(r"\bif\s*\(\s*(false|0)\b", line):
            hits.append(i)
    return hits


def run_node(js: str) -> subprocess.CompletedProcess[str]:
    checks = "\n".join(
        f"results.{g} = (typeof {g} === 'function');" for g in REQUIRED_GLOBALS
    )
    epilogue = textwrap.dedent(
        f"""
        const results = {{}};
        try {{ {checks} }} catch (e) {{ results.__checkError = e && e.message; }}
        if (__uncaught) {{
            console.log('UNCAUGHT:', __uncaught.constructor.name, '-', __uncaught.message);
        }}
        console.log('RESULTS:' + JSON.stringify(results));
        """
    ).strip()

    payload = f"{_NODE_PRELUDE}\n\n{js}\n\n{epilogue}\n"
    with tempfile.NamedTemporaryFile("w", suffix="_js_smoke.js", delete=False, encoding="utf-8") as fh:
        fh.write(payload)
        tmp = fh.name
    try:
        return subprocess.run(
            ["node", tmp], capture_output=True, text=True, timeout=60
        )
    finally:
        Path(tmp).unlink(missing_ok=True)


def collect_js() -> str:
    """Junta o(s) <script> inline do template + os ficheiros /static/js na ordem real."""
    html = TEMPLATE.read_text(encoding="utf-8", errors="replace")
    parts = [neutralize_jinja(b) for b in extract_inline_scripts(html)]

    ordered = list(STATIC_JS_ORDER)
    if STATIC_JS_DIR.is_dir():
        for extra in sorted(p.name for p in STATIC_JS_DIR.glob("*.js")):
            if extra not in ordered:
                ordered.append(extra)
    for name in ordered:
        fpath = STATIC_JS_DIR / name
        if fpath.is_file():
            parts.append(fpath.read_text(encoding="utf-8", errors="replace"))

    return "\n;\n".join(parts)


def check() -> list[str]:
    """Devolve lista de erros. Vazia = OK."""
    errors: list[str] = []
    js = collect_js()
    if not js.strip():
        return ["Nao foi encontrado JavaScript (nem inline em index.html nem em static/js)."]

    dups = find_duplicate_top_level_functions(js)
    for name, lines in sorted(dups.items()):
        errors.append(f"Funcao de topo duplicada: {name} (linhas {lines}) - a 2a sombreia a 1a.")

    for ln in find_constant_conditions(js):
        errors.append(f"Condicao constante (codigo morto) na linha {ln} do <script>.")

    proc = run_node(js)
    out = (proc.stdout or "") + "\n" + (proc.stderr or "")

    for line in out.splitlines():
        if line.startswith("UNCAUGHT:"):
            errors.append(f"Excecao nao apanhada ao avaliar o <script>: {line[9:].strip()}")

    m = re.search(r"RESULTS:(\{.*\})", out)
    if not m:
        errors.append(
            "O script nao chegou ao fim da avaliacao (sem linha RESULTS). "
            f"Saida do node:\n{out.strip()[:1500]}"
        )
    else:
        import json

        results = json.loads(m.group(1))
        for g in REQUIRED_GLOBALS:
            if not results.get(g):
                errors.append(f"Funcao global critica em falta apos avaliacao: {g}")

    return errors


class JsSmokeTest(unittest.TestCase):
    def test_inline_script_loads_clean(self) -> None:
        if shutil.which("node") is None:
            self.skipTest("node nao esta no PATH")
        errors = check()
        self.assertEqual(errors, [], "\n- " + "\n- ".join(errors))


def main() -> int:
    if shutil.which("node") is None:
        print("SKIP: node nao esta no PATH")
        return 0
    errors = check()
    if errors:
        print("FALHOU js_smoke:")
        for e in errors:
            print("  -", e)
        return 1
    print("OK js_smoke: <script> avalia limpo, sem duplicados, globais presentes.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

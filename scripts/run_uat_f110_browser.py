# -*- coding: utf-8 -*-
"""
Automacao local em browser para submeter o UAT F110 no SAP Script Web Cockpit.
"""

from __future__ import annotations

import argparse
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests


# ###################################################################################
# (1) CONSTANTES E MODELOS
# ###################################################################################

DEFAULT_BASE_URL = os.getenv("SAP_COCKPIT_URL", "http://127.0.0.1:8010").rstrip("/")
DEFAULT_PROCESS = "UAT Simulação"
DEFAULT_SUBPROCESS = "Executar F110"
DEFAULT_ENVIRONMENT = os.getenv("SAP_COCKPIT_AMBIENTE", "QAD").strip().upper() or "QAD"
DEFAULT_REQUEST_OPTION = os.getenv("SAP_COCKPIT_REQUEST_OPTION", "4").strip() or "4"
DEFAULT_REQUEST_TYPE = os.getenv("SAP_COCKPIT_REQUEST_TYPE", "1").strip() or "1"
DEFAULT_BROWSER_CHANNEL = os.getenv("SAP_COCKPIT_BROWSER_CHANNEL", "chrome").strip() or "chrome"
DEFAULT_TIMEOUT_MS = int(os.getenv("SAP_COCKPIT_TIMEOUT_MS", "30000"))
DEFAULT_JOB_WAIT_SECONDS = int(os.getenv("SAP_COCKPIT_JOB_WAIT_SECONDS", "30"))


@dataclass(frozen=True)
class BrowserAutomationResult:
    job_id: str
    job_state: str
    job_url: str
    workflow_label: str
    workflow_subprocess: str


# ###################################################################################
# (2) ARGUMENTOS E UTILITÁRIOS
# ###################################################################################

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Submete o UAT F110 no SAP Script Web Cockpit via browser local."
    )
    parser.add_argument("--base-url", default=DEFAULT_BASE_URL, help="URL base do cockpit web.")
    parser.add_argument("--processo", default=DEFAULT_PROCESS, help="Nome do processo a selecionar no modal.")
    parser.add_argument(
        "--subprocesso",
        default=DEFAULT_SUBPROCESS,
        help="Nome do subprocesso a selecionar no modal.",
    )
    parser.add_argument(
        "--modo",
        choices=("massivo", "individual"),
        default="massivo",
        help="Modo do fluxo UAT no assistente.",
    )
    parser.add_argument(
        "--ambiente",
        default=DEFAULT_ENVIRONMENT,
        help="Ambiente do job no modal do cockpit.",
    )
    parser.add_argument(
        "--request-option",
        default=DEFAULT_REQUEST_OPTION,
        choices=("1", "2", "4"),
        help="Opção de request do modal (1, 2 ou 4).",
    )
    parser.add_argument("--request-number", default="", help="Número da request quando a opção 1 for usada.")
    parser.add_argument("--request-desc", default="", help="Descrição da request quando a opção 2 for usada.")
    parser.add_argument(
        "--request-type",
        default=DEFAULT_REQUEST_TYPE,
        choices=("1", "2"),
        help="Tipo da request quando a opção 2 for usada.",
    )
    parser.add_argument(
        "--caminho-ficheiro",
        default="",
        help="Caminho do ficheiro Excel a anexar ao job, se aplicável.",
    )
    parser.add_argument("--nome-pasta", default="", help="Nome opcional da pasta do job.")
    parser.add_argument(
        "--param",
        action="append",
        default=[],
        metavar="NOME=VALOR",
        help="Parâmetro Web extra do modal. Pode repetir.",
    )
    parser.add_argument(
        "--browser-channel",
        default=DEFAULT_BROWSER_CHANNEL,
        help="Canal do browser Playwright. Ex.: chrome, chromium.",
    )
    parser.add_argument("--headless", action="store_true", help="Executar sem janela visível.")
    parser.add_argument(
        "--timeout-ms",
        type=int,
        default=DEFAULT_TIMEOUT_MS,
        help="Timeout base do Playwright em milissegundos.",
    )
    parser.add_argument(
        "--job-wait-seconds",
        type=int,
        default=DEFAULT_JOB_WAIT_SECONDS,
        help="Tempo máximo para acompanhar o job depois da submissão.",
    )
    return parser


def normalize_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def parse_param_pairs(values: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for raw_item in values:
        item = str(raw_item or "").strip()
        if not item or "=" not in item:
            raise ValueError(f"Parâmetro inválido: {raw_item!r}. Use o formato NOME=VALOR.")
        key, raw_value = item.split("=", 1)
        key = key.strip()
        if not key:
            raise ValueError(f"Parâmetro inválido: {raw_item!r}. O nome não pode ficar vazio.")
        result[key] = raw_value.strip()
    return result


def _button_locator(page, label: str, scope: str = ""):
    pattern = re.compile(re.escape(normalize_text(label)), re.I)
    locator = page.locator(scope or "body").locator("button").filter(has_text=pattern)
    if locator.count() == 0:
        raise RuntimeError(f"Não encontrei botão com o texto '{label}'.")
    return locator.first


def _click_button(page, label: str, scope: str = "") -> None:
    button = _button_locator(page, label, scope=scope)
    button.click()


def _click_initial_authorization_flow(page) -> None:
    page.locator("#nav-item-autorizacoes").click()
    page.locator("#authorization-chat-messages").wait_for(state="visible")
    _click_button(page, "Executar processo", scope="#authorization-chat-messages")
    _click_button(page, "UAT Simulação", scope="#authorization-chat-messages")
    _click_button(page, "Executar F110", scope="#authorization-chat-messages")


def _click_workflow_mode(page, mode: str) -> None:
    if mode == "individual":
        _click_button(page, "Alteração Individual", scope="#authorization-chat-messages")
    else:
        _click_button(page, "Alteração Massiva", scope="#authorization-chat-messages")


def _set_modal_select(page, selector: str, *, label: str | None = None, value: str | None = None) -> None:
    if label is not None:
        page.locator(selector).select_option(label=label)
        return
    if value is not None:
        page.locator(selector).select_option(value=value)
        return
    raise ValueError("É necessário indicar label ou value para selecionar um elemento.")


def _fill_modal_params(page, params: dict[str, str]) -> None:
    for key, value in params.items():
        locator = page.locator(f'#web-params-container [data-web-param="{key}"]')
        if locator.count() == 0:
            continue
        element = locator.first
        tag_name = element.evaluate("(node) => node.tagName.toLowerCase()")
        if tag_name == "select":
            element.select_option(label=value)
        else:
            element.fill(value)


def _wait_for_job_record(base_url: str, job_id: str, timeout_seconds: int) -> dict[str, Any]:
    deadline = time.monotonic() + max(1, timeout_seconds)
    last_payload: dict[str, Any] | None = None

    while time.monotonic() <= deadline:
        response = requests.get(f"{base_url}/api/jobs/{job_id}", timeout=15)
        response.raise_for_status()
        last_payload = response.json()
        state = normalize_text(last_payload.get("state")).lower()
        if state in {"running", "succeeded", "failed", "succeeded_with_warnings", "cancelled"}:
            return last_payload
        time.sleep(2)

    return last_payload or {"id": job_id, "state": "pending"}


# ###################################################################################
# (3) EXECUÇÃO DO BROWSER
# ###################################################################################

def run_browser_automation(args: argparse.Namespace) -> BrowserAutomationResult:
    try:
        extra_params = parse_param_pairs(list(args.param or []))
    except ValueError as exc:
        raise RuntimeError(str(exc)) from exc

    caminho_ficheiro = str(args.caminho_ficheiro or "").strip()
    if caminho_ficheiro:
        file_path = Path(caminho_ficheiro)
        if not file_path.exists():
            raise FileNotFoundError(f"Ficheiro não encontrado: {file_path}")
        caminho_ficheiro = str(file_path.resolve())

    try:
        from playwright.sync_api import sync_playwright
    except Exception as exc:  # pragma: no cover - depende da instalação local
        raise RuntimeError(
            "A biblioteca Playwright não está disponível. Instale as dependências do projeto "
            "e execute `python -m playwright install chromium` se o browser local não estiver presente."
        ) from exc

    with sync_playwright() as playwright:
        browser = None
        launch_errors: list[str] = []

        for channel in (args.browser_channel, "chrome", "chromium"):
            if not channel:
                continue
            try:
                browser = playwright.chromium.launch(
                    headless=bool(args.headless),
                    channel=channel,
                )
                break
            except Exception as exc:
                launch_errors.append(f"{channel}: {exc}")

        if browser is None:
            raise RuntimeError(
                "Não foi possível abrir um browser local via Playwright.\n"
                + "\n".join(f"- {item}" for item in launch_errors)
            )

        try:
            context = browser.new_context(viewport={"width": 1680, "height": 1100})
            page = context.new_page()
            page.set_default_timeout(int(args.timeout_ms))

            page.goto(str(args.base_url).rstrip("/"), wait_until="domcontentloaded")
            page.locator("#nav-item-autorizacoes").click()
            page.locator("#authorization-chat-messages").wait_for(state="visible")

            _click_initial_authorization_flow(page)
            _click_workflow_mode(page, args.modo)

            modal = page.locator("#modal-novo-job")
            modal.wait_for(state="visible")
            page.locator("#modal-novo-job.active").wait_for(state="visible")

            _set_modal_select(page, "#ambiente-select", label=str(args.ambiente))
            _set_modal_select(page, "#processo-select", label=str(args.processo))

            # O select de subprocesso é alimentado dinamicamente após a escolha do processo.
            page.wait_for_function(
                """() => {
                    const select = document.querySelector('#subprocesso-select');
                    return select && select.options && select.options.length > 1;
                }"""
            )
            _set_modal_select(page, "#subprocesso-select", label=str(args.subprocesso))

            if caminho_ficheiro:
                page.evaluate("(path) => window.setExcelPath(path)", caminho_ficheiro)

            if str(args.request_option) == "1":
                _set_modal_select(page, "#request-option-select", value="1")
                page.locator("#request-number-input").fill(str(args.request_number or "").strip())
            elif str(args.request_option) == "2":
                _set_modal_select(page, "#request-option-select", value="2")
                page.locator("#request-desc-input").fill(str(args.request_desc or "").strip())
                _set_modal_select(page, "#request-type-select", value=str(args.request_type))
            else:
                _set_modal_select(page, "#request-option-select", value="4")

            if str(args.nome_pasta or "").strip():
                page.locator("#nome-pasta-input").fill(str(args.nome_pasta).strip())

            if extra_params:
                _fill_modal_params(page, extra_params)

            submit_button = page.locator("#submit-job-btn")
            with page.expect_response(
                lambda response: response.request.method == "POST" and response.url.endswith("/jobs")
            ) as response_info:
                submit_button.click()

            response = response_info.value
            if not response.ok:
                raise RuntimeError(f"Falha a submeter o job no cockpit. HTTP {response.status}")

            job_payload = response.json()
            job_id = str(job_payload.get("id") or "").strip()
            if not job_id:
                raise RuntimeError("O cockpit não devolveu job_id na resposta de criação.")

            job_snapshot = _wait_for_job_record(str(args.base_url).rstrip("/"), job_id, int(args.job_wait_seconds))
            current_state = normalize_text(job_snapshot.get("state") or job_payload.get("state") or "pending")

            return BrowserAutomationResult(
                job_id=job_id,
                job_state=current_state,
                job_url=f"{str(args.base_url).rstrip('/')}/api/jobs/{job_id}",
                workflow_label=str(args.processo),
                workflow_subprocess=str(args.subprocesso),
            )
        finally:
            browser.close()


# ###################################################################################
# (4) CLI
# ###################################################################################

def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    result = run_browser_automation(args)
    print(f"Job criado: {result.job_id}")
    print(f"Estado atual: {result.job_state}")
    print(f"Processo: {result.workflow_label}")
    print(f"Subprocesso: {result.workflow_subprocess}")
    print(f"URL de referencia: {result.job_url}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

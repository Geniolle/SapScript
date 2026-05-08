from rich.console import Console
from rich.panel import Panel
from rich.table import Table
import os

console = Console()

def _send_log(msg: str):
    job_id = os.getenv("SAP_JOB_ID")
    if not job_id:
        return
    
    api_url = os.getenv("SAP_API_BASE_URL")
    token = os.getenv("SAP_WORKER_TOKEN")
    if not api_url or not token:
        return
        
    try:
        import requests
        requests.post(
            f"{api_url}/api/jobs/{job_id}/log",
            headers={"X-Worker-Token": token},
            json={"log_line": msg},
            timeout=1
        )
    except Exception:
        pass


def mostrar_titulo(ambiente=None, sistema=None, cliente=None, utilizador=None):
    linhas = []
    if ambiente:
        linhas.append(f"[bold cyan]Ambiente:[/] {ambiente}")
    if sistema:
        linhas.append(f"[bold cyan]Sistema:[/] {sistema}")
    if cliente:
        linhas.append(f"[bold cyan]Cliente:[/] {cliente}")
    if utilizador:
        linhas.append(f"[bold cyan]Utilizador:[/] {utilizador}")

    corpo = "\n".join(linhas) if linhas else "[dim]Inicialização...[/dim]"

    console.print(
        Panel(
            corpo,
            title="[bold white]SAP COCKPIT[/bold white]",
            border_style="bright_blue",
            expand=True,
        )
    )


def mostrar_ambientes(ambientes: dict):
    tabela = Table(
        title="Ambientes disponíveis",
        header_style="bold bright_white",
        row_styles=["none", "on rgb(30,30,30)"],
    )
    tabela.add_column("Opção", style="bold cyan", width=8)
    tabela.add_column("Sigla", style="bold green", width=10)
    tabela.add_column("Descrição", style="white")

    for k, (sigla, nome) in ambientes.items():
        tabela.add_row(str(k), sigla, nome)

    console.print(tabela)


def mostrar_processos(pastas: list[str]):
    tabela = Table(
        title="Processos disponíveis",
        header_style="bold bright_white",
        row_styles=["none", "on rgb(30,30,30)"],
    )
    tabela.add_column("Opção", style="bold cyan", width=8)
    tabela.add_column("Processo", style="white")

    for i, pasta in enumerate(pastas, 1):
        tabela.add_row(str(i), pasta)

    console.print(tabela)


def mostrar_subprocessos(scripts_py: list[str]):
    tabela = Table(
        title="Sub-Processos disponíveis",
        header_style="bold bright_white",
        row_styles=["none", "on rgb(30,30,30)"],
    )
    tabela.add_column("Opção", style="bold cyan", width=8)
    tabela.add_column("Sub-Processo", style="white")

    for i, script in enumerate(scripts_py, 1):
        tabela.add_row(str(i), script)

    tabela.add_row(str(len(scripts_py) + 1), "Voltar ao menu de Processos")
    console.print(tabela)


def info(msg: str):
    console.print(f"[bold cyan][INFO][/bold cyan] {msg}")
    _send_log(f"[INFO] {msg}")


def ok(msg: str):
    console.print(f"[bold green][OK][/bold green] {msg}")
    _send_log(f"[OK] {msg}")


def warn(msg: str):
    console.print(f"[bold yellow][WARN][/bold yellow] {msg}")
    _send_log(f"[WARN] {msg}")


def erro(msg: str):
    console.print(f"[bold red][ERRO][/bold red] {msg}")
    _send_log(f"[ERRO] {msg}")


def destaque(msg: str):
    console.print(f"[bold white on blue] {msg} [/bold white on blue]")
    _send_log(f"--- {msg} ---")


def linha():
    console.rule(style="dim")
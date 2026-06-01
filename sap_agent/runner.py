from __future__ import annotations

import argparse
from pathlib import Path

from rich.console import Console

from .config import AgentConfig, JiraConfig, SapConnectionConfig, load_yaml
from .diagnosis import SapDiagnosisEngine
from .jira_client import JiraClient
from .safety import SafetyGuard
from .sap_rfc_client import SapRfcClient
from .validators import SapReadOnlyValidator

console = Console()


def build_engine(config_path: str | Path) -> tuple[SapDiagnosisEngine, AgentConfig, JiraConfig]:
    yaml_data = load_yaml(config_path)
    agent_config = AgentConfig.from_yaml(config_path)
    jira_config = JiraConfig.from_env_and_yaml(yaml_data)
    sap_config = SapConnectionConfig.from_env()
    safety = SafetyGuard.build(
        allow_write_operations=agent_config.allow_write_operations,
        allowed_functions=agent_config.sap_allowed_functions,
        allowed_tables=agent_config.sap_allowed_tables,
    )
    sap_client = SapRfcClient(config=sap_config, safety_guard=safety)
    validator = SapReadOnlyValidator(sap_client)
    return SapDiagnosisEngine(validator), agent_config, jira_config


def run(config_path: str | Path) -> None:
    engine, agent_config, jira_config = build_engine(config_path)
    jira = JiraClient(jira_config)
    tickets = jira.search_tickets()
    console.print(f"[bold]Tickets encontrados:[/bold] {len(tickets)}")

    for ticket in tickets:
        console.rule(f"{ticket.key} - {ticket.summary}")
        diagnosis = engine.diagnose(ticket)
        comment = diagnosis.to_jira_comment(agent_config.jira_comment_prefix)
        console.print(comment)
        if jira_config.update_jira:
            jira.add_comment(ticket.key, comment)
            console.print(f"[green]Comentário gravado no JIRA:[/green] {ticket.key}")
        else:
            console.print("[yellow]Modo dry-run: comentário não gravado no JIRA.[/yellow]")


def main() -> None:
    parser = argparse.ArgumentParser(description="SAP read-only JIRA diagnosis agent")
    parser.add_argument("--config", default="config/sap_agent.yaml", help="Path to yaml configuration file")
    args = parser.parse_args()
    run(args.config)


if __name__ == "__main__":
    main()

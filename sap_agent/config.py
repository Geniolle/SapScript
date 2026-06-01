from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv


@dataclass(frozen=True)
class SapConnectionConfig:
    """SAP RFC connection settings.

    The password is intentionally loaded only from environment variables.
    Do not store SAP credentials in yaml files or in Git.
    """

    user: str
    passwd: str
    ashost: str
    sysnr: str
    client: str
    lang: str = "EN"

    @classmethod
    def from_env(cls, prefix: str = "SAP_") -> "SapConnectionConfig":
        load_dotenv()
        required = ["USER", "PASSWD", "ASHOST", "SYSNR", "CLIENT"]
        missing = [f"{prefix}{key}" for key in required if not os.getenv(f"{prefix}{key}")]
        if missing:
            raise RuntimeError(f"Missing SAP environment variables: {', '.join(missing)}")

        return cls(
            user=os.environ[f"{prefix}USER"],
            passwd=os.environ[f"{prefix}PASSWD"],
            ashost=os.environ[f"{prefix}ASHOST"],
            sysnr=os.environ[f"{prefix}SYSNR"],
            client=os.environ[f"{prefix}CLIENT"],
            lang=os.getenv(f"{prefix}LANG", "EN"),
        )

    def as_pyrfc_params(self) -> dict[str, str]:
        return {
            "user": self.user,
            "passwd": self.passwd,
            "ashost": self.ashost,
            "sysnr": self.sysnr,
            "client": self.client,
            "lang": self.lang,
        }


@dataclass(frozen=True)
class JiraConfig:
    base_url: str
    email: str
    api_token: str
    jql: str
    max_results: int = 10
    update_jira: bool = False

    @classmethod
    def from_env_and_yaml(cls, yaml_data: dict[str, Any]) -> "JiraConfig":
        load_dotenv()
        jira_data = yaml_data.get("jira", {})
        required_env = ["JIRA_BASE_URL", "JIRA_EMAIL", "JIRA_API_TOKEN"]
        missing = [key for key in required_env if not os.getenv(key)]
        if missing:
            raise RuntimeError(f"Missing JIRA environment variables: {', '.join(missing)}")
        return cls(
            base_url=os.environ["JIRA_BASE_URL"].rstrip("/"),
            email=os.environ["JIRA_EMAIL"],
            api_token=os.environ["JIRA_API_TOKEN"],
            jql=jira_data.get("jql", ""),
            max_results=int(jira_data.get("max_results", 10)),
            update_jira=bool(jira_data.get("update_jira", False)),
        )


@dataclass(frozen=True)
class AgentConfig:
    modules_enabled: list[str] = field(default_factory=list)
    safe_mode: bool = True
    allow_write_operations: bool = False
    sap_allowed_functions: list[str] = field(default_factory=list)
    sap_allowed_tables: list[str] = field(default_factory=list)
    web_research_enabled: bool = True
    jira_comment_prefix: str = "Pré-análise automática SAP"

    @classmethod
    def from_yaml(cls, path: str | Path) -> "AgentConfig":
        data = load_yaml(path)
        safety = data.get("safety", {})
        sap = data.get("sap", {})
        research = data.get("research", {})
        jira = data.get("jira", {})
        return cls(
            modules_enabled=list(data.get("modules_enabled", [])),
            safe_mode=bool(safety.get("safe_mode", True)),
            allow_write_operations=bool(safety.get("allow_write_operations", False)),
            sap_allowed_functions=list(sap.get("allowed_functions", [])),
            sap_allowed_tables=list(sap.get("allowed_tables", [])),
            web_research_enabled=bool(research.get("enabled", True)),
            jira_comment_prefix=str(jira.get("comment_prefix", "Pré-análise automática SAP")),
        )


def load_yaml(path: str | Path) -> dict[str, Any]:
    with Path(path).open("r", encoding="utf-8") as stream:
        data = yaml.safe_load(stream) or {}
    if not isinstance(data, dict):
        raise ValueError(f"Configuration file must contain a YAML object: {path}")
    return data

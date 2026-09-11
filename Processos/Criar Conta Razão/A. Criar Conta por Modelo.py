"""Processo executável pelo VS Code para criação de Conta Razão por modelo."""

from sap_rfc.gl_account_cli import main

WEB_PARAMS = [
    {"name": "environment", "label": "Ambiente", "type": "select", "options": ["DEV", "QAD", "PRD"], "required": True},
    {"name": "account", "label": "Conta", "type": "text", "required": True},
    {"name": "company", "label": "Empresa de destino", "type": "text", "required": True},
    {"name": "model_company", "label": "Empresa-modelo", "type": "text", "required": True},
    {"name": "alternative_account", "label": "Conta alternativa (fora de PT)", "type": "text", "required": False},
]
WEB_CONFIG = {"title": "Criar Conta Razão por modelo", "requires_confirmation": True}

if __name__ == "__main__":
    raise SystemExit(main())

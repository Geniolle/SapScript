import os
import pytest
from pathlib import Path
from sap_session import (
    resolve_sap_target_from_env,
    _validate_target,
    read_dotenv_values,
    derive_client_from_key,
    SapTarget,
)

def test_derive_client():
    assert derive_client_from_key("SPACLNT001") == "001"
    assert derive_client_from_key("S4DCLNT100") == "100"
    assert derive_client_from_key("INVALID") == ""

def test_resolve_sap_target_explicit_key_ignores_workflow_vars(tmp_path, monkeypatch):
    # Setup temporary .env file
    env_file = tmp_path / ".env"
    env_file.write_text(
        "SAP_CONNECTION_SPACLNT001=SAP-CUA-CONN\n"
        "SAP_CLIENT_SPACLNT001=001\n"
        "SAP_SYSTEM_SPACLNT001=SPA\n"
        "SAP_USER_SPACLNT001=cua_user\n"
        "SAP_PASSWORD_SPACLNT001=cua_pwd\n"
        "SAP_LANGUAGE_SPACLNT001=PT\n",
        encoding="utf-8"
    )
    
    # Force read_dotenv_values to use our temp file
    monkeypatch.setenv("SAP_AUTH_ENV_FILE", str(env_file))
    
    # Set conflicting workflow vars in os.environ
    monkeypatch.setenv("WORKFLOW_SAP_SYSTEM", "S4D")
    monkeypatch.setenv("WORKFLOW_SAP_CLIENT", "100")
    monkeypatch.setenv("WORKFLOW_SAP_KEY", "S4DCLNT100")
    monkeypatch.setenv("SAP_USER", "global_user")
    monkeypatch.setenv("SAP_LANGUAGE", "EN")
    
    # Read dotenv
    env_values = read_dotenv_values()
    
    # Resolve SPACLNT001 target
    target = resolve_sap_target_from_env("SPACLNT001", env_values=env_values)
    
    assert target.key == "SPACLNT001"
    assert target.system_name == "SPA"      # workflow var ignored
    assert target.client == "001"            # workflow var ignored
    assert target.connection_name == "SAP-CUA-CONN"
    assert target.user == "cua_user"         # specific prioritized over global
    assert target.password == "cua_pwd"
    assert target.language == "PT"           # specific prioritized over global

def test_target_validation_and_errors():
    # 1. Missing password
    target_no_pwd = SapTarget(
        key="SPACLNT001",
        system_name="SPA",
        connection_name="CUA_CONN",
        client="001",
        user="user",
        password="",
        language="PT",
        saplogon_path="path"
    )
    with pytest.raises(RuntimeError) as exc_info:
        _validate_target(target_no_pwd)
    assert "Configuração SAP CUA incompleta" in str(exc_info.value)
    assert "SAP_PASSWORD_SPACLNT001" in str(exc_info.value)
    
    # Check that password value is not leaked in the exception message
    assert "cua_pwd" not in str(exc_info.value)

def test_dotenv_physical_reloads(tmp_path, monkeypatch):
    # Setup temporary .env file
    env_file = tmp_path / ".env"
    env_file.write_text(
        "SAP_CONNECTION_SPACLNT001=CONN-V1\n"
        "SAP_USER_SPACLNT001=user-v1\n",
        encoding="utf-8"
    )
    
    monkeypatch.setenv("SAP_AUTH_ENV_FILE", str(env_file))
    
    env_values = read_dotenv_values()
    assert env_values.get("SAP_USER_SPACLNT001") == "user-v1"
    
    # Modify .env file physically
    env_file.write_text(
        "SAP_CONNECTION_SPACLNT001=CONN-V2\n"
        "SAP_USER_SPACLNT001=user-v2\n",
        encoding="utf-8"
    )
    
    # Reread should reflect changes immediately
    env_values_v2 = read_dotenv_values()
    assert env_values_v2.get("SAP_USER_SPACLNT001") == "user-v2"

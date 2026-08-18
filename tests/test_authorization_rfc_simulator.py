import os
import sys

import pytest

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2")))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2", "worker")))

from authorization_rfc_simulator import resolve_simulated_system, simulate_authorization_rfc_analysis


def test_resolve_simulated_system_defaults_to_dev():
    system = resolve_simulated_system("")
    assert system["choice"] == "DEV"
    assert system["key"] == "S4DCLNT100"
    assert system["system"] == "S4D"
    assert system["client"] == "100"


def test_simulate_authorization_rfc_analysis_returns_expected_schema():
    result = simulate_authorization_rfc_analysis("csilva")
    assert result["success"] is True
    assert result["code"] == "analysis_complete"
    assert result["execution_mode"] == "RFC"
    assert result["target_user"] == "CSILVA"
    assert result["target_system"]["key"] == "S4DCLNT100"
    assert result["execution_system"]["system"] == "S4D"
    assert result["data_source_verified"] is True
    assert result["worker_feature_version"] == "authorization-tables-v1"
    assert len(result["queries"]) == 3
    assert result["summary"]["total_roles"] == 2
    assert result["summary"]["total_profiles"] == 1


def test_simulate_authorization_rfc_analysis_requires_user():
    with pytest.raises(ValueError):
        simulate_authorization_rfc_analysis("")

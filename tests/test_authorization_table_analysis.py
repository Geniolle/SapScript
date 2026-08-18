import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2")))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2", "worker")))

import pytest
from unittest.mock import MagicMock
from authorization_table_analysis import (
    normalize_sap_user,
    validate_target_system_key,
    normalize_sap_date,
    classify_validity,
    classify_assignment_origin,
    deduplicate_roles,
    build_authorization_summary,
    analyze_user_authorizations,
    query_cua_table,
)
from sap_agent.sap_gui_actions import SapGuiResult

# 1. Testar normalização do utilizador
def test_normalize_sap_user():
    assert normalize_sap_user("  csilva  ") == "CSILVA"
    assert normalize_sap_user("user.1_name-") == "USER.1_NAME-"
    
    with pytest.raises(ValueError):
        normalize_sap_user("")
    with pytest.raises(ValueError):
        normalize_sap_user("A" * 41)
    with pytest.raises(ValueError):
        normalize_sap_user("user@domain")

# 2. Testar validação de target_system_key
def test_validate_target_system_key():
    assert validate_target_system_key("  S4DCLNT100  ") == "S4DCLNT100"
    assert validate_target_system_key("SPACLNT001") == "SPACLNT001"
    
    with pytest.raises(ValueError):
        validate_target_system_key("S4D")
    with pytest.raises(ValueError):
        validate_target_system_key("S4DCLNT")
    with pytest.raises(ValueError):
        validate_target_system_key("CLNT100")

# 3. Testar parser de data
def test_normalize_sap_date():
    assert normalize_sap_date("20260720") == "2026-07-20"
    assert normalize_sap_date("2026-07-20") == "2026-07-20"
    assert normalize_sap_date("20.07.2026") == "2026-07-20"
    assert normalize_sap_date("20/07/2026") == "2026-07-20"
    assert normalize_sap_date("00000000") == ""
    assert normalize_sap_date("") == ""

# 4. Testar classificações de validades
def test_classify_validity():
    today = "2026-07-20"
    
    # Ativa
    assert classify_validity("2026-07-15", "2026-07-25", today) == "active"
    assert classify_validity("", "2026-07-25", today) == "active"
    assert classify_validity("2026-07-15", "", today) == "active"
    assert classify_validity("2026-07-15", "9999-12-31", today) == "active"
    
    # Expirada
    assert classify_validity("2026-07-10", "2026-07-18", today) == "expired"
    
    # Futura
    assert classify_validity("2026-07-22", "2026-07-25", today) == "future"
    
    # Indeterminada
    assert classify_validity("invalid_date", "2026-07-25", today) == "undetermined"

# 5. Testar classificações de origens
def test_classify_assignment_origin():
    assert classify_assignment_origin("") == {"origin": "direct", "origin_label": "Direta"}
    assert classify_assignment_origin("X") == {"origin": "organizational_management", "origin_label": "Organização RH"}
    assert classify_assignment_origin("C") == {"origin": "composite_role", "origin_label": "Role composta"}
    assert classify_assignment_origin("E") == {"origin": "enterprise_portal", "origin_label": "Enterprise Portal"}
    assert classify_assignment_origin("Y") == {"origin": "other", "origin_label": "Outra origem"}

# 6. Testar remoção de duplicados
def test_deduplicate_roles():
    roles = [
        {"role": "R1", "subsystem": "S4DCLNT100", "valid_from": "2026-01-01", "valid_to": "2026-12-31", "assignment_origin_code": ""},
        {"role": "R1", "subsystem": "S4DCLNT100", "valid_from": "2026-01-01", "valid_to": "2026-12-31", "assignment_origin_code": ""}, # Duplicado
        {"role": "R2", "subsystem": "S4DCLNT100", "valid_from": "2026-01-01", "valid_to": "2026-12-31", "assignment_origin_code": ""}
    ]
    deduped = deduplicate_roles(roles)
    assert len(deduped) == 2
    assert deduped[0]["role"] == "R1"
    assert deduped[1]["role"] == "R2"

# 7. Testar resumo
def test_build_authorization_summary():
    roles = [
        {"role": "R1", "validity_status": "active", "assignment_origin": "direct"},
        {"role": "R2", "validity_status": "active", "assignment_origin": "direct"},
        {"role": "R3", "validity_status": "expired", "assignment_origin": "composite_role"},
        {"role": "R4", "validity_status": "future", "assignment_origin": "organizational_management"}
    ]
    profiles = [{"profile": "P1"}, {"profile": "P2"}]
    summary = build_authorization_summary(roles, profiles)
    
    assert summary["total_roles"] == 4
    assert summary["active_roles"] == 2
    assert summary["expired_roles"] == 1
    assert summary["future_roles"] == 1
    assert summary["direct_roles"] == 2
    assert summary["indirect_roles"] == 2
    assert summary["total_profiles"] == 2

# 8. Mocks de se16n_query_with_session para testar analyze_user_authorizations
def test_analyze_user_not_assigned_to_system(monkeypatch):
    mock_session = MagicMock()
    mock_session.Info.SystemName = "SPA"
    mock_session.Info.Client = "001"
    
    # Mock query_cua_table to return empty for USZBVSYS (not assigned)
    def mock_query(session, table, filters, max_rows):
        assert table == "USZBVSYS"
        assert filters[0]["value"] == "TESTUSER"
        assert filters[1]["value"] == "S4DCLNT100"
        return []
        
    monkeypatch.setattr("authorization_table_analysis.query_cua_table", mock_query)
    
    res = analyze_user_authorizations(mock_session, "TESTUSER", "S4DCLNT100")
    
    assert res["success"] is True
    assert res["code"] == "user_not_assigned_to_system"
    assert res["user_assigned_to_system"] is False
    assert len(res["roles"]) == 0

def test_analyze_user_authorizations_success(monkeypatch):
    mock_session = MagicMock()
    mock_session.Info.SystemName = "SPA"
    mock_session.Info.Client = "001"
    
    def mock_query(session, table, filters, max_rows):
        if table == "USZBVSYS":
            return [{"BNAME": "TESTUSER", "SUBSYSTEM": "S4DCLNT100"}]
        elif table == "USLA04":
            return [
                {"BNAME": "TESTUSER", "SUBSYSTEM": "S4DCLNT100", "AGR_NAME": "ZROLE1", "FROM_DAT": "20260101", "TO_DAT": "20261231", "ORG_FLAG": ""},
                {"BNAME": "TESTUSER", "SUBSYSTEM": "S4DCLNT100", "AGR_NAME": "ZROLE2", "FROM_DAT": "20260101", "TO_DAT": "20261231", "ORG_FLAG": "X"}
            ]
        elif table == "USL04":
            return [{"BNAME": "TESTUSER", "SUBSYSTEM": "S4DCLNT100", "PROFILE": "ZPROF1"}]
        return []
        
    monkeypatch.setattr("authorization_table_analysis.query_cua_table", mock_query)
    
    res = analyze_user_authorizations(mock_session, "TESTUSER", "S4DCLNT100")
    
    assert res["success"] is True
    assert res["code"] == "analysis_complete"
    assert res["user_assigned_to_system"] is True
    assert len(res["roles"]) == 2
    assert len(res["profiles"]) == 1
    assert res["summary"]["total_roles"] == 2
    assert res["summary"]["active_roles"] == 2
    
    # Check that password or tokens are not present
    assert "password" not in res
    assert "token" not in res

def test_analyze_invalid_cua_session():
    mock_session = MagicMock()
    mock_session.Info.SystemName = "S4D"
    mock_session.Info.Client = "100"
    
    res = analyze_user_authorizations(mock_session, "TESTUSER", "S4DCLNT100")
    assert res["success"] is False
    assert res["code"] == "invalid_cua_session"

def test_validate_completed_analysis():
    from sap_tasks import validate_completed_analysis, SapExecutionError
    
    valid_res = {
        "success": True,
        "code": "analysis_complete",
        "data_source_verified": True,
        "worker_feature_version": "authorization-tables-v1",
        "queries": [
            {"table": "USZBVSYS", "executed": True, "filters_applied": True, "row_count": 1},
            {"table": "USLA04", "executed": True, "filters_applied": True, "row_count": 5},
            {"table": "USL04", "executed": True, "filters_applied": True, "row_count": 2}
        ]
    }
    validate_completed_analysis(valid_res)
    
    legacy_res = valid_res.copy()
    legacy_res["worker_feature_version"] = "old-version"
    with pytest.raises(SapExecutionError):
        validate_completed_analysis(legacy_res)
        
    missing_queries_res = valid_res.copy()
    missing_queries_res["queries"] = []
    with pytest.raises(SapExecutionError):
        validate_completed_analysis(missing_queries_res)
        
    incomplete_tables_res = valid_res.copy()
    incomplete_tables_res["queries"] = [
        {"table": "USZBVSYS", "executed": True, "filters_applied": True, "row_count": 1}
    ]
    with pytest.raises(SapExecutionError):
        validate_completed_analysis(incomplete_tables_res)

def test_se16_navigation_and_query_classic_list():
    from sap_agent.sap_gui_actions import se16_query_with_session
    mock_session = MagicMock()
    
    okcd_mock = MagicMock()
    table_elem_mock = MagicMock()
    
    navigation_called = []
    def dynamic_find_by_id(path):
        if "okcd" in path:
            navigation_called.append("SE16")
            return okcd_mock
        if "DATABROWSE-TABLENAME" in path:
            return table_elem_mock
        
        user_area = MagicMock()
        if "SE16" in navigation_called:
            # Output screen elements
            h1 = MagicMock(); h1.Id = "wnd[0]/usr/lbl[1,2]"; h1.Text = "BNAME"; h1.Type = "GuiLabel"; h1.col = 2; h1.row = 1
            h2 = MagicMock(); h2.Id = "wnd[0]/usr/lbl[1,15]"; h2.Text = "SUBSYSTEM"; h2.Type = "GuiLabel"; h2.col = 15; h2.row = 1
            h3 = MagicMock(); h3.Id = "wnd[0]/usr/lbl[1,30]"; h3.Text = "AGR_NAME"; h3.Type = "GuiLabel"; h3.col = 30; h3.row = 1
            
            d1 = MagicMock(); d1.Id = "wnd[0]/usr/lbl[2,2]"; d1.Text = "TESTUSER"; d1.Type = "GuiLabel"; d1.col = 2; d1.row = 2
            d2 = MagicMock(); d2.Id = "wnd[0]/usr/lbl[2,15]"; d2.Text = "S4DCLNT100"; d2.Type = "GuiLabel"; d2.col = 15; d2.row = 2
            d3 = MagicMock(); d3.Id = "wnd[0]/usr/lbl[2,30]"; d3.Text = "ZROLE_ABC"; d3.Type = "GuiLabel"; d3.col = 30; d3.row = 2
            
            elements = [h1, h2, h3, d1, d2, d3]
            user_area.Children.Count = len(elements)
            user_area.Children.Element = lambda idx: elements[idx]
            user_area.findById = lambda p: None
        else:
            # Selection screen elements
            f1 = MagicMock(); f1.Id = "wnd[0]/usr/txtBNAME-LOW"; f1.Name = "BNAME-LOW"; f1.Type = "GuiTextField"; f1.Changeable = True
            f2 = MagicMock(); f2.Id = "wnd[0]/usr/txtSUBSYSTEM-LOW"; f2.Name = "SUBSYSTEM-LOW"; f2.Type = "GuiCTextField"; f2.Changeable = True
            
            elements = [f1, f2]
            user_area.Children.Count = len(elements)
            user_area.Children.Element = lambda idx: elements[idx]
            user_area.findById = lambda p: None
            
        return user_area

    mock_session.findById = dynamic_find_by_id

    res = se16_query_with_session(
        mock_session,
        table="USLA04",
        filters=[
            {"field": "BNAME", "value": "TESTUSER"},
            {"field": "SUBSYSTEM", "value": "S4DCLNT100"}
        ],
        strict_filters=True
    )
    
    assert res.success is True
    assert len(res.rows) == 1
    assert res.rows[0]["BNAME"] == "TESTUSER"
    assert res.rows[0]["SUBSYSTEM"] == "S4DCLNT100"
    assert res.rows[0]["AGR_NAME"] == "ZROLE_ABC"
    assert "SE16" in navigation_called
    assert okcd_mock.Text == "/nSE16"

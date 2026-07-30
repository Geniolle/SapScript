import sys
import os
from unittest.mock import MagicMock, patch
import pytest

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2")))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2", "worker")))

from hr_data_analysis import search_hr_user_data_rfc


def test_search_hr_empty_query():
    res = search_hr_user_data_rfc("")
    assert res["success"] is False
    assert res["total"] == 0
    assert "vazio" in res["message"]


@patch("hr_data_analysis._open_rfc_connection")
@patch("hr_data_analysis._read_rfc_table")
def test_search_hr_by_pernr_success(mock_read_rfc, mock_open_rfc):
    mock_conn = MagicMock()
    mock_open_rfc.return_value = mock_conn

    def side_effect_read_table(conn, table, fields, filters, max_rows=10):
        if table == "PA0002":
            return [{"PERNR": "00012345", "VORNA": "Clayton", "NACHN": "Silva", "CNAME": "Clayton Silva"}]
        if table == "PA0105":
            return [
                {"PERNR": "00012345", "SUBTY": "0010", "USRID": "", "USRID_LONG": "clayton.silva@salsajeans.com"},
                {"PERNR": "00012345", "SUBTY": "0001", "USRID": "CSILVA", "USRID_LONG": ""},
            ]
        if table == "PA0001":
            return [{"PERNR": "00012345", "ORGEH": "100020", "PLSTX": "Equipa de Arquitetura SAP", "STELL": "300"}]
        return []

    mock_read_rfc.side_effect = side_effect_read_table

    res = search_hr_user_data_rfc("12345", target_system_key="S4PCLNT100")
    assert res["success"] is True
    assert res["total"] == 1
    item = res["data"][0]
    assert item["pernr"] == "00012345"
    assert item["user_id"] == "CSILVA"
    assert item["full_name"] == "Clayton Silva"
    assert item["email"] == "clayton.silva@salsajeans.com"
    assert item["team"] == "Equipa de Arquitetura SAP"


@patch("hr_data_analysis._open_rfc_connection")
@patch("hr_data_analysis._read_rfc_table")
def test_search_hr_fallback_user_master(mock_read_rfc, mock_open_rfc):
    mock_conn = MagicMock()
    mock_open_rfc.return_value = mock_conn

    def side_effect_read_table(conn, table, fields, filters, max_rows=10):
        if table == "PA0002":
            return []
        if table == "USR21":
            return [{"BNAME": "CSILVA", "PERSNUMBER": "000999111", "ADDRNUMBER": "000888222"}]
        if table == "ADRP":
            return [{"PERSNUMBER": "000999111", "NAME_TEXT": "Clayton Silva", "NAME_FIRST": "Clayton", "NAME_LAST": "Silva"}]
        if table == "ADR6":
            return [{"ADDRNUMBER": "000888222", "PERSNUMBER": "000999111", "SMTP_ADDR": "clayton.silva@salsajeans.com"}]
        if table == "USR02":
            return [{"BNAME": "CSILVA", "CLASS": "EQUIPA_IT"}]
        return []

    mock_read_rfc.side_effect = side_effect_read_table

    res = search_hr_user_data_rfc("CSILVA", target_system_key="S4PCLNT100")
    assert res["success"] is True
    assert res["total"] == 1
    item = res["data"][0]
    assert item["user_id"] == "CSILVA"
    assert item["full_name"] == "Clayton Silva"
    assert item["email"] == "clayton.silva@salsajeans.com"
    assert item["team"] == "EQUIPA_IT"

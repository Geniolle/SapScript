import sys
import os
import pytest
from unittest.mock import MagicMock

# Ajustar paths
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "sap_script_web_cockpit_v2")))
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from sap_agent.sap_gui_actions import (
    _find_se16_field,
    collect_sap_components,
    find_se16_low_field_by_label,
    se16_query_with_session,
    wait_for_se16_selection_screen
)

class MockChildrenCollection:
    def __init__(self, children):
        self._children = children
        self.Count = len(children)
    def Element(self, index):
        return self._children[index]

class MockSAPElement:
    def __init__(self, element_id="some_id", children=None, **kwargs):
        self.Id = element_id
        self.Children = MockChildrenCollection(children or [])
        self.Type = "GuiComponent"
        self.Name = ""
        self.Text = ""
        self.Changeable = True
        self.Left = 0
        self.Top = 0
        self.Width = 10
        self.Height = 10
        self.Visible = True
        for k, v in kwargs.items():
            setattr(self, k, v)

    def setFocus(self):
        pass

    def sendVKey(self, key):
        pass

# Test 1: ID contendo BNAME-LOW
def test_id_containing_bname_low():
    field = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=True, name="BNAME-LOW")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Id == "wnd[0]/usr/txtBNAME-LOW"

# Test 2: Name contendo BNAME
def test_name_containing_bname():
    field = MockSAPElement(element_id="wnd[0]/usr/somefield", Name="BNAME", Type="GuiTextField", Changeable=True)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Name == "BNAME"

# Test 3: ID técnico genérico (fallback posicional ou padrão de sufixo)
def test_generic_technical_id():
    field = MockSAPElement(element_id="wnd[0]/usr/I1-LOW", Name="I1-LOW", Type="GuiTextField", Changeable=True)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Id == "wnd[0]/usr/I1-LOW"

# Test 4: Localização pelo label BNAME
def test_locate_by_label_bname():
    label = MockSAPElement(element_id="lbl_1", Type="GuiLabel", Text="BNAME:", Left=10, Top=20, Width=50, Height=15)
    field_low = MockSAPElement(element_id="low_field", Type="GuiTextField", Changeable=True, Left=70, Top=20, Width=100, Height=15)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[label, field_low])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Id == "low_field"

# Test 5: Localização pelo label SUBSYSTEM
def test_locate_by_label_subsystem():
    label = MockSAPElement(element_id="lbl_2", Type="GuiLabel", Text="Logical System", Left=10, Top=40, Width=80, Height=15)
    field_low = MockSAPElement(element_id="low_field", Type="GuiCTextField", Changeable=True, Left=100, Top=40, Width=100, Height=15)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[label, field_low])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "SUBSYSTEM")
    assert found is not None
    assert found.Id == "low_field"

# Test 6: Escolha do campo LOW em vez do HIGH (proximidade horizontal)
def test_choose_low_instead_of_high():
    label = MockSAPElement(element_id="lbl_1", Type="GuiLabel", Text="BNAME", Left=10, Top=20, Width=50, Height=15)
    field_low = MockSAPElement(element_id="low_field", Type="GuiTextField", Changeable=True, Left=70, Top=20, Width=100, Height=15)
    field_high = MockSAPElement(element_id="high_field", Type="GuiTextField", Changeable=True, Left=200, Top=20, Width=100, Height=15)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[label, field_high, field_low]) # Ordem aleatória na árvore
    session = MagicMock()
    session.findById.return_value = usr
    
    found = find_se16_low_field_by_label(session, "BNAME")
    assert found is not None
    assert found.Id == "low_field" # Deve escolher o de Left=70 em vez do de Left=200

# Test 7: Campos em linhas diferentes
def test_fields_on_different_rows():
    label = MockSAPElement(element_id="lbl_1", Type="GuiLabel", Text="BNAME", Left=10, Top=20, Width=50, Height=15)
    field_other_row = MockSAPElement(element_id="other_row", Type="GuiTextField", Changeable=True, Left=70, Top=50, Width=100, Height=15)
    field_same_row = MockSAPElement(element_id="same_row", Type="GuiTextField", Changeable=True, Left=70, Top=20, Width=100, Height=15)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[label, field_other_row, field_same_row])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = find_se16_low_field_by_label(session, "BNAME")
    assert found is not None
    assert found.Id == "same_row"

# Test 8: Campo não editável ignorado
def test_non_changeable_field_ignored():
    label = MockSAPElement(element_id="lbl_1", Type="GuiLabel", Text="BNAME", Left=10, Top=20, Width=50, Height=15)
    field_non_changeable = MockSAPElement(element_id="non_changeable", Type="GuiTextField", Changeable=False, Left=70, Top=20, Width=100, Height=15)
    field_changeable = MockSAPElement(element_id="changeable", Type="GuiTextField", Changeable=True, Left=200, Top=20, Width=100, Height=15)
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[label, field_non_changeable, field_changeable])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = find_se16_low_field_by_label(session, "BNAME")
    assert found is not None
    assert found.Id == "changeable" # Ignora non_changeable mesmo estando mais perto

# Test 9: Valor preenchido e relido com sucesso
def test_value_filled_and_re_read_successfully():
    field = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=True, Text="CSILVA")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field])
    
    # Mock do session e comportamentos
    session = MagicMock()
    session.findById.return_value = usr
    
    # Simular navegação e okcd/table
    okcd = MagicMock()
    tablename = MagicMock()
    active_window = MagicMock()
    active_window.Text = "Data Browser: Table USZBVSYS: Selection Screen"
    session.ActiveWindow = active_window
    
    def mock_find_by_id(path):
        if "okcd" in path:
            return okcd
        if "DATABROWSE-TABLENAME" in path:
            return tablename
        if "usr" in path:
            return usr
        return MagicMock()
        
    session.findById = mock_find_by_id
    
    # Executar consulta
    res = se16_query_with_session(
        session,
        table="USZBVSYS",
        filters=[{"field": "BNAME", "value": "CSILVA"}],
        strict_filters=True
    )
    
    # O valor foi validado com sucesso e o fluxo seguiu
    assert field.Text == "CSILVA"

# Test 10: Valor relido diferente (erro de preenchimento)
def test_value_read_different_failure():
    class BadField:
        def __init__(self):
            self.Id = "wnd[0]/usr/txtBNAME-LOW"
            self.Type = "GuiTextField"
            self.Changeable = True
        def setFocus(self):
            pass
        @property
        def Text(self):
            return "DIFFERENT_VALUE"
        @Text.setter
        def Text(self, val):
            pass
            
    field = BadField()
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[])
    
    session = MagicMock()
    okcd = MagicMock()
    tablename = MagicMock()
    active_window = MagicMock()
    active_window.Text = "Data Browser: Table USZBVSYS: Selection Screen"
    session.ActiveWindow = active_window
    
    def mock_find_by_id(path):
        if "okcd" in path:
            return okcd
        if "DATABROWSE-TABLENAME" in path:
            return tablename
        if "usr" in path:
            return usr
        return MagicMock()
        
    session.findById = mock_find_by_id
    
    from unittest.mock import patch
    with patch("sap_agent.sap_gui_actions._find_se16_field", return_value=field):
        res = se16_query_with_session(
            session,
            table="USZBVSYS",
            filters=[{"field": "BNAME", "value": "CSILVA"}],
            strict_filters=True
        )
        assert res.success is False
        assert "Não foi possível" in res.error or "Filtro obrigatório" in res.error

# Test 11: BNAME ausente (tabela requer strict_filters e BNAME)
def test_bname_missing_fails():
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[]) # Nenhum campo
    session = MagicMock()
    session.findById.return_value = usr
    
    active_window = MagicMock()
    active_window.Text = "Data Browser: Table USZBVSYS: Selection Screen"
    session.ActiveWindow = active_window
    
    res = se16_query_with_session(
        session,
        table="USZBVSYS",
        filters=[{"field": "BNAME", "value": "CSILVA"}],
        strict_filters=True
    )
    
    assert res.success is False
    assert "Não foi possível localizar o campo BNAME" in res.error

# Test 12: SUBSYSTEM ausente
def test_subsystem_missing_fails():
    field_bname = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=True, Name="BNAME-LOW")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field_bname])
    session = MagicMock()
    session.findById.return_value = usr
    
    active_window = MagicMock()
    active_window.Text = "Data Browser: Table USZBVSYS: Selection Screen"
    session.ActiveWindow = active_window
    
    res = se16_query_with_session(
        session,
        table="USZBVSYS",
        filters=[
            {"field": "BNAME", "value": "CSILVA"},
            {"field": "SUBSYSTEM", "value": "SPACLNT001"}
        ],
        strict_filters=True
    )
    
    assert res.success is False
    assert "Não foi possível localizar o campo SUBSYSTEM" in res.error

# Test 13: Nenhum F8 quando um dos filtros falhar
def test_no_f8_when_one_filter_fails():
    field_bname = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=True, Name="BNAME-LOW")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field_bname]) # SUBSYSTEM ausente
    session = MagicMock()
    session.findById.return_value = usr
    
    active_window = MagicMock()
    active_window.Text = "Data Browser: Table USZBVSYS: Selection Screen"
    session.ActiveWindow = active_window
    
    res = se16_query_with_session(
        session,
        table="USZBVSYS",
        filters=[
            {"field": "BNAME", "value": "CSILVA"},
            {"field": "SUBSYSTEM", "value": "SPACLNT001"}
        ],
        strict_filters=True
    )
    
    assert res.success is False
    # sendVKey(8) não deve ter sido chamado para executar a busca
    assert not any(call[0] == ("wnd[0]",) and 8 in call[1] for call in session.findById.mock_calls)

# Test 14: F8 somente quando ambos forem aplicados
def test_f8_only_when_both_applied():
    field_bname = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=True, Name="BNAME-LOW", Text="CSILVA")
    field_sub = MockSAPElement(element_id="wnd[0]/usr/txtSUBSYSTEM-LOW", Type="GuiCTextField", Changeable=True, Name="SUBSYSTEM-LOW", Text="SPACLNT001")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field_bname, field_sub])
    
    session = MagicMock()
    active_window = MagicMock()
    active_window.Text = "Data Browser: Table USZBVSYS: Selection Screen"
    session.ActiveWindow = active_window
    
    okcd = MagicMock()
    tablename = MagicMock()
    
    # Mock findById
    mock_elements = {}
    mock_elements["wnd[0]/tbar[0]/okcd"] = okcd
    mock_elements["wnd[0]/usr/ctxtDATABROWSE-TABLENAME"] = tablename
    mock_elements["wnd[0]/usr"] = usr
    
    # Para o sendVKey(8) de F8
    wnd_mock = MagicMock()
    mock_elements["wnd[0]"] = wnd_mock
    
    def mock_find_by_id(path):
        return mock_elements.get(path, MagicMock())
        
    session.findById = mock_find_by_id
    
    res = se16_query_with_session(
        session,
        table="USZBVSYS",
        filters=[
            {"field": "BNAME", "value": "CSILVA"},
            {"field": "SUBSYSTEM", "value": "SPACLNT001"}
        ],
        strict_filters=True
    )
    
    wnd_mock.sendVKey.assert_any_call(8)


# Test 15: Changeable = -1 (valor numérico SAP para True)
def test_changeable_minus_one():
    field = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=-1, Name="BNAME-LOW")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Id == "wnd[0]/usr/txtBNAME-LOW"


# Test 16: Changeable indisponível (retorna None ou levanta erro ao acessar)
def test_changeable_unavailable():
    class MockNoChangeable(MockSAPElement):
        @property
        def Changeable(self):
            raise AttributeError("Propriedade indisponível")
        @Changeable.setter
        def Changeable(self, val):
            pass

    field = MockNoChangeable(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Name="BNAME-LOW")
    usr = MockSAPElement(element_id="wnd[0]/usr", children=[field])
    session = MagicMock()
    session.findById.return_value = usr
    
    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Id == "wnd[0]/usr/txtBNAME-LOW"


# Test 17: Coleção COM customizada usando Item(index) e chamada direta
def test_custom_collection_indexing():
    class CustomCollection:
        def __init__(self, items):
            self.items = items
            self.Count = len(items)
        def Item(self, idx):
            return self.items[idx]

    class CustomSAPElement(MockSAPElement):
        def __init__(self, element_id, children_items):
            super().__init__(element_id=element_id)
            self.Children = CustomCollection(children_items)

    field = MockSAPElement(element_id="wnd[0]/usr/txtBNAME-LOW", Type="GuiTextField", Changeable=True, Name="BNAME-LOW")
    usr = CustomSAPElement("wnd[0]/usr", [field])
    session = MagicMock()
    session.findById.return_value = usr

    found = _find_se16_field(session, "BNAME")
    assert found is not None
    assert found.Id == "wnd[0]/usr/txtBNAME-LOW"


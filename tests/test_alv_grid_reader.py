import os
import sys
import pytest
from unittest.mock import MagicMock

# Importar o módulo H. CUA_ADICIONAR.py de forma dinâmica devido ao ponto e espaço no nome
import importlib.util
spec = importlib.util.spec_from_file_location(
    "cua_adicionar",
    os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Processos", "Funções PFCG", "H. CUA_ADICIONAR.py"))
)
cua = importlib.util.module_from_spec(spec)
spec.loader.exec_module(cua)

class MockCOMCollection:
    def __init__(self, items):
        self._items = items
        self.Count = len(items)
    def __iter__(self):
        return iter(self._items)
    def Item(self, index):
        return self._items[index]

class MockChildrenCollection:
    def __init__(self, children):
        self._children = children
        self.Count = len(children)
    def __call__(self, index):
        return self._children[index]

class MockElement:
    def __init__(self, element_id="some_id", **kwargs):
        self.Id = element_id
        self.Children = MockChildrenCollection([])
        for k, v in kwargs.items():
            setattr(self, k, v)

# Testes de detecção do ALV Grid (is_alv_grid e find_alv_grid)
def test_is_alv_grid():
    # 1. Grid clássico GuiShell/GridView
    grid1 = MockElement(Type="GuiShell", SubType="GridView")
    assert cua.is_alv_grid(grid1) is True

    # 2. Componente de tipo diferente de GuiShell mas com capacidades/métodos de Grid (fallback)
    grid2 = MockElement(
        Type="GuiGridView",
        RowCount=10,
        ColumnCount=5,
        GetCellValue=lambda r, c: "test"
    )
    assert cua.is_alv_grid(grid2) is True

    # 3. Componente inválido sem propriedades necessárias
    bad_elem = MockElement(Type="GuiLabel")
    assert cua.is_alv_grid(bad_elem) is False

def test_find_alv_grid_direct_id():
    # Mock da sessão que retorna o grid pelo ID direto
    mock_grid = MockElement(element_id="wnd[0]/usr/cntlGRID1/shellcont/shell", Type="GuiShell", SubType="GridView")
    
    mock_session = MagicMock()
    def find_by_id(id_str):
        if "cntlGRID1/shellcont/shell" in id_str:
            return mock_grid
        raise Exception("Not found")
    mock_session.findById = find_by_id

    grid = cua.find_alv_grid(mock_session)
    assert grid is not None
    assert grid.Id == "wnd[0]/usr/cntlGRID1/shellcont/shell"

def test_find_alv_grid_by_limited_search():
    # Mock da sessão onde a busca direta falha, mas a busca recursiva em wnd[0]/usr encontra o grid
    mock_grid = MockElement(element_id="some_nested_grid", Type="GuiShell", SubType="GridView")
    mock_usr = MockElement(element_id="wnd[0]/usr")
    # Coloca o grid como filho de usr
    mock_usr.Children = MockChildrenCollection([mock_grid])

    mock_session = MagicMock()
    def find_by_id(id_str):
        if id_str == "wnd[0]/usr":
            return mock_usr
        raise Exception("Not found")
    mock_session.findById = find_by_id

    grid = cua.find_alv_grid(mock_session)
    assert grid is not None
    assert grid.Id == "some_nested_grid"

# Testes de leitura e mapeamento de colunas (ColumnOrder e GetCellValue)
def test_get_grid_column_ids():
    # 1. ColumnOrder como tuple
    grid_tuple = MockElement(ColumnOrder=("MANDT", "BNAME"))
    assert cua.get_grid_column_ids(grid_tuple) == ["MANDT", "BNAME"]

    # 2. ColumnOrder como coleção COM customizada
    grid_com = MockElement(ColumnOrder=MockCOMCollection(["MANDT", "BNAME", "SUBSYSTEM"]))
    assert cua.get_grid_column_ids(grid_com) == ["MANDT", "BNAME", "SUBSYSTEM"]

    # 3. Fallback usando ColumnCount e GetColumnKey
    grid_fallback = MockElement(
        ColumnCount=2,
        GetColumnKey=lambda idx: ["COL1", "COL2"][idx]
    )
    assert cua.get_grid_column_ids(grid_fallback) == ["COL1", "COL2"]

def test_map_alv_columns():
    grid = MockElement()
    grid.GetColumnTitle = lambda col_id: {
        "MANDT": "Mandante",
        "BNAME": "Utilizador",
        "SUBSYSTEM": "Sistema receptor",
        "AGR_NAME": "Função",
        "FROM_DAT": "Válido de",
        "TO_DAT": "Válido até",
        "ORG_FLAG": "Org"
    }.get(col_id, "")

    cols = ["MANDT", "BNAME", "SUBSYSTEM", "AGR_NAME", "FROM_DAT", "TO_DAT", "ORG_FLAG"]
    col_map = cua.map_alv_columns(grid, cols)
    
    assert col_map["BNAME"] == "BNAME"
    assert col_map["SUBSYSTEM"] == "SUBSYSTEM"
    assert col_map["AGR_NAME"] == "AGR_NAME"

def test_read_alv_grid_success():
    grid = MockElement(RowCount=24, ColumnCount=7)
    grid.ColumnOrder = ("MANDT", "BNAME", "SUBSYSTEM", "AGR_NAME", "FROM_DAT", "TO_DAT", "ORG_FLAG")
    grid.GetColumnTitle = lambda col_id: col_id
    
    # Simular GetCellValue
    db_cells = {
        (0, "MANDT"): "001",
        (0, "BNAME"): "S80001870",
        (0, "SUBSYSTEM"): "S4DCLNT100",
        (0, "AGR_NAME"): "Z_BR_ROLE",
        (0, "FROM_DAT"): "12.07.2026",
        (0, "TO_DAT"): "31.12.9999",
        (0, "ORG_FLAG"): "",
    }
    grid.GetCellValue = lambda r, c: db_cells.get((r, c), "")

    res = cua.read_alv_grid(grid)
    assert len(res) == 1
    assert res[0]["BNAME"] == "S80001870"
    assert res[0]["SUBSYSTEM"] == "S4DCLNT100"
    assert res[0]["ORG_FLAG"] == ""

def test_read_alv_grid_zero_rows():
    grid = MockElement(RowCount=0, ColumnCount=7)
    grid.ColumnOrder = ("MANDT", "BNAME", "SUBSYSTEM", "AGR_NAME", "FROM_DAT", "TO_DAT", "ORG_FLAG")
    grid.GetColumnTitle = lambda col_id: col_id
    
    res = cua.read_alv_grid(grid)
    assert len(res) == 0

def test_read_alv_grid_unexpected_column():
    grid = MockElement(RowCount=1, ColumnCount=3)
    grid.ColumnOrder = ("BNAME", "SUBSYSTEM", "UNEXPECTED_COL")
    grid.GetColumnTitle = lambda col_id: col_id
    grid.GetCellValue = lambda r, c: "VAL"

    res = cua.read_alv_grid(grid)
    assert len(res) == 1
    assert res[0]["BNAME"] == "VAL"
    assert res[0]["SUBSYSTEM"] == "VAL"
    # Campos que faltavam no grid retornam vazios
    assert res[0]["MANDT"] == ""

# Teste de regressão para wnd[0]/sbar vazia e ALV Grid com dados
def test_regression_gui_shell_grid_view_with_data(monkeypatch):
    """
    Garante que se o Grid existir com 24 linhas e 7 colunas, mesmo com sbar
    vazia, o resultado é considerado válido e lido corretamente.
    """
    mock_grid = MockElement(
        element_id="wnd[0]/usr/cntlGRID1/shellcont/shell",
        Type="GuiShell",
        SubType="GridView",
        RowCount=24,
        ColumnCount=7,
        ColumnOrder=("MANDT", "BNAME", "SUBSYSTEM", "AGR_NAME", "FROM_DAT", "TO_DAT", "ORG_FLAG")
    )
    
    mock_grid.GetColumnTitle = lambda col_id: col_id
    
    # 24 registros simulados
    db_cells = {}
    for r in range(24):
        db_cells[(r, "MANDT")] = "001"
        db_cells[(r, "BNAME")] = "TESTUSER"
        db_cells[(r, "SUBSYSTEM")] = "TESTSYS"
        db_cells[(r, "AGR_NAME")] = f"ROLE_{r}"
        db_cells[(r, "FROM_DAT")] = "15.07.2026"
        db_cells[(r, "TO_DAT")] = "31.12.9999"
        db_cells[(r, "ORG_FLAG")] = ""
        
    mock_grid.GetCellValue = lambda r, c: db_cells.get((r, c), "")

    mock_session = MagicMock()
    # Mock do sbar retornando texto vazio
    mock_sbar = MockElement(Text="", MessageType="")
    
    def find_by_id(id_str):
        if "cntlGRID1/shellcont/shell" in id_str:
            return mock_grid
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        raise Exception(f"Not found: {id_str}")
        
    mock_session.findById = find_by_id

    # Executar a leitura polimórfica
    res = cua.ler_resultados_usla04(mock_session, "TESTUSER", "TESTSYS")
    
    assert len(res) == 24
    assert res[0]["BNAME"] == "TESTUSER"
    assert res[23]["AGR_NAME"] == "ROLE_23"

# Testes de detecção de TableControl e consulta vazia (Fase 3)
def test_encontrar_table_control_is_defined():
    """Garante que a função _encontrar_table_control existe e localiza TableControl."""
    mock_tc = MockElement(element_id="some_tc", Type="GuiTableControl")
    mock_usr = MockElement(element_id="wnd[0]/usr")
    mock_usr.Children = MockChildrenCollection([mock_tc])
    
    found = cua._encontrar_table_control(mock_usr)
    assert found is not None
    assert found.Id == "some_tc"

def test_empty_alv_grid_returns_zero_rows_without_error():
    """Garante que um ALV Grid com RowCount=0 é considerado resultado vazio válido sem erro."""
    mock_grid = MockElement(
        element_id="wnd[0]/usr/cntlGRID1/shellcont/shell",
        Type="GuiShell",
        SubType="GridView",
        RowCount=0,
        ColumnCount=7
    )
    mock_session = MagicMock()
    mock_sbar = MockElement(Text="", MessageType="")
    
    def find_by_id(id_str):
        if "cntlGRID1/shellcont/shell" in id_str:
            return mock_grid
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        raise Exception(f"Not found: {id_str}")
    mock_session.findById = find_by_id
    mock_session.Children = MockChildrenCollection([])
    
    res = cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert res == []

def test_status_bar_no_entries_found_portuguese():
    """Garante que mensagem da status bar em português é interpretada como consulta vazia."""
    mock_session = MagicMock()
    mock_sbar = MockElement(Text="Nenhuma entrada encontrada", MessageType="S")
    
    def find_by_id(id_str):
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        raise Exception(f"Not found: {id_str}")
    mock_session.findById = find_by_id
    mock_session.Children = MockChildrenCollection([])
    
    res = cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert res == []

def test_status_bar_no_entries_found_english():
    """Garante que mensagem da status bar em inglês é interpretada como consulta vazia."""
    mock_session = MagicMock()
    mock_sbar = MockElement(Text="No entries selected", MessageType="S")
    
    def find_by_id(id_str):
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        raise Exception(f"Not found: {id_str}")
    mock_session.findById = find_by_id
    mock_session.Children = MockChildrenCollection([])
    
    res = cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert res == []

def test_popup_no_data_found():
    """Garante que se existir um popup contendo 'No data found' ou aviso equivalente, é tratado como vazio."""
    mock_session = MagicMock()
    mock_sbar = MockElement(Text="", MessageType="")
    
    # Popup mockado com texto "No entries selected"
    mock_popup_lbl = MockElement(Text="No entries selected")
    mock_popup = MockElement(element_id="wnd[1]")
    mock_popup.Children = MockChildrenCollection([mock_popup_lbl])
    
    # Sessão possui um popup (Children.Count > 0)
    mock_session.Children = MockCOMCollection([MockElement("wnd[0]"), mock_popup])
    
    def find_by_id(id_str):
        if id_str == "wnd[1]":
            return mock_popup
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        raise Exception(f"Not found: {id_str}")
    mock_session.findById = find_by_id
    
    res = cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert res == []

def test_technical_error_raises_exception():
    """Garante que falha técnica levanta erro explicativo e não é mascarada como lista vazia."""
    mock_session = MagicMock()
    # Força um erro no findById para wnd[0]/sbar
    mock_session.findById.side_effect = RuntimeError("Erro de rede COM no SAP GUI")
    mock_session.Children = MockChildrenCollection([])
    
    with pytest.raises(RuntimeError) as exc_info:
        cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert "Erro de rede COM" in str(exc_info.value)

def test_regression_cua_adicionar_scenario_8_roles():
    """
    Simula o caso de regressão completo:
      - BNAME = S80001870
      - SUBSYSTEM = S4DCLNT100
      - Consulta da USLA04 vazia (sem registos)
      - 8 funções a classificar
    Valida se as 8 funções são classificadas como INEXISTENTE (aptas para inserção).
    """
    # 1. Consulta retorna vazio (simulado por sbar com 'No entries selected')
    mock_session = MagicMock()
    mock_sbar = MockElement(Text="No entries selected", MessageType="S")
    
    def find_by_id(id_str):
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        raise Exception(f"Not found: {id_str}")
    mock_session.findById = find_by_id
    mock_session.Children = MockChildrenCollection([])
    
    resultados_usla04 = cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert resultados_usla04 == []
    
    # 2. Executar a classificação para as 8 roles
    roles_list = [
        "Z_BR_PURCHSERV_SPECIALIST",
        "ZORG_TODAS_EMPRESAS",
        "ZORG_BP_Z001_GENERALPARTNERS",
        "ZORG_BP_Z003_RELATEDPARTNERS",
        "ZORG_BP_GERAL",
        "ZORG_BP_LOGISTICS_CUSTOMER",
        "ZORG_BP_FLVN01_LOGISTICS_VENDO",
        "Z_BR_TYPE_BP_GERAL"
    ]
    
    from datetime import date
    hoje = date(2026, 7, 15)
    
    classificacoes = cua._classificar_linhas_usla04(resultados_usla04, roles_list, hoje)
    
    # Todas as 8 funções devem estar classificadas como INEXISTENTE e ter zero ativas
    for role in roles_list:
        role_up = role.upper()
        assert role_up in classificacoes
        assert classificacoes[role_up]["classe"] == "INEXISTENTE"
        assert classificacoes[role_up]["n_ativas"] == 0


def test_regression_no_table_entries_found_specified_key():
    """
    Garante o comportamento obrigatório da Fase 4:
      - BNAME = S80001870
      - SUBSYSTEM = S4DCLNT100
      - 8 funções no Excel
      - barra de status = "No table entries found for specified key"
      - nenhum Grid disponível
    Resultado esperado:
      - consulta válida;
      - zero registos na USLA04;
      - 8 funções inexistentes;
      - zero erros de validação;
      - 8 funções aptas para inserção.
    """
    # 1. Mock do session
    mock_session = MagicMock()
    mock_sbar = MockElement(Text="No table entries found for specified key", MessageType="S")
    
    def find_by_id(id_str):
        if id_str == "wnd[0]/sbar":
            return mock_sbar
        # Se tentar buscar grid ou outro componente, levantar exceção para garantir que
        # a execução foi interrompida imediatamente conforme requisito 3.
        raise Exception(f"Erro: Tentou buscar '{id_str}' mesmo com a barra de status confirmando ausência de dados.")
        
    mock_session.findById = find_by_id
    mock_session.Children = MockChildrenCollection([])
    
    # 2. Chamar ler_resultados_usla04
    resultados_usla04 = cua.ler_resultados_usla04(mock_session, "S80001870", "S4DCLNT100")
    assert resultados_usla04 == []
    
    # 3. Classificar as 8 roles
    roles_list = [
        "Z_BR_PURCHSERV_SPECIALIST",
        "ZORG_TODAS_EMPRESAS",
        "ZORG_BP_Z001_GENERALPARTNERS",
        "ZORG_BP_Z003_RELATEDPARTNERS",
        "ZORG_BP_GERAL",
        "ZORG_BP_LOGISTICS_CUSTOMER",
        "ZORG_BP_FLVN01_LOGISTICS_VENDO",
        "Z_BR_TYPE_BP_GERAL"
    ]
    from datetime import date
    hoje = date(2026, 7, 15)
    
    classificacoes = cua._classificar_linhas_usla04(resultados_usla04, roles_list, hoje)
    
    assert len(classificacoes) == 8
    for role in roles_list:
        role_up = role.upper()
        assert classificacoes[role_up]["classe"] == "INEXISTENTE"
        assert classificacoes[role_up]["n_ativas"] == 0



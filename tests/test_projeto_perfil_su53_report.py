# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib.util
from pathlib import Path

import openpyxl


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPT_PATH = PROJECT_ROOT / "Processos" / "Projeto Autorizações" / "Pesquisa_Erros_SU53.py"


def carregar_modulo_su53():
    spec = importlib.util.spec_from_file_location("pesquisa_su53_test", SCRIPT_PATH)
    modulo = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(modulo)
    return modulo


def erro(data_hora, servidor, objeto, nome_servico, *, tcode_app="", campos=None):
    return {
        "timestamp": data_hora,
        "data_hora": data_hora,
        "servidor": servidor,
        "objeto": objeto,
        "rc": 4,
        "programa": "PROGRAMA",
        "linha": 3,
        "tcode_app": tcode_app,
        "nome_servico": nome_servico,
        "campos": campos or [("ACTVT", "03")],
        "tcode_falha": "",
    }


def test_filtra_hoje_na_instancia_da_falha_mais_recente():
    mod = carregar_modulo_su53()
    erros = [
        erro("23.10.2026 08:34:48", "app02", "M_EINK_FRG", "ME23N"),
        erro("23.10.2026 08:24:08", "app02", "B_BUP_DCPD", "FAGLB03"),
        erro("23.10.2026 08:20:00", "app01", "B_BUP_DCPD", "BP"),
        erro("22.10.2026 16:28:04", "app02", "F_SKA1_BUK", "FAGLL03H"),
    ]

    filtrados = mod.filtrar_erros_atuais(erros, data_local="23.10.2026")

    assert [(item["nome_servico"], item["objeto"]) for item in filtrados] == [
        ("ME23N", "M_EINK_FRG"),
        ("FAGLB03", "B_BUP_DCPD"),
    ]


def test_agrupa_por_aplicacao_e_objeto_com_fallback_name():
    mod = carregar_modulo_su53()
    erros = [
        erro("23.10.2026 08:34:48", "app02", "M_EINK_FRG", "ME23N"),
        erro("23.10.2026 08:24:08", "app02", "B_BUP_DCPD", "FAGLB03"),
        erro("23.10.2026 08:23:00", "app02", "B_BUP_DCPD", "BP", tcode_app="FAGLB03"),
    ]

    _, objetos = mod.agrupar_erros_para_analise(erros)

    assert [(item["aplicacao"], item["objeto"]) for item in objetos] == [
        ("ME23N", "M_EINK_FRG"),
        ("FAGLB03", "B_BUP_DCPD"),
    ]
    assert objetos[1]["total_ocorrencias"] == 2


def test_proposta_devolve_funcao_descricao_e_linhas(tmp_path):
    mod = carregar_modulo_su53()
    caminho = tmp_path / "perfis.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Proposta"
    ws.append(["FUNÇÃO", "DESCRIÇÃO"])
    ws.append(["Z_ACCOUNTING_REPORT", "Finance General REPORT"])
    ws.append(["FB03", None])
    ws.append(["FAGLB03", None])
    wb.save(caminho)

    roles = mod.buscar_funcao_individual_para_transacao(str(caminho), "FAGLB03")

    assert roles[0]["role"] == "Z_ACCOUNTING_REPORT"
    assert roles[0]["descricao"] == "Finance General REPORT"
    assert roles[0]["linha_funcao"] == 2
    assert roles[0]["linha_transacao"] == 4


def test_formata_relatorio_com_aplicacao_objeto_e_funcao():
    mod = carregar_modulo_su53()
    objeto = {
        "aplicacao": "FAGLB03",
        "nome_aplicacao": "Exibição de saldos de contas do Razão",
        "objeto": "B_BUP_DCPD",
        "data_hora_recente": "23.10.2026 08:24:08",
        "campos_relevantes": [("ACTVT", "03")],
    }
    roles = [{
        "role": "Z_ACCOUNTING_REPORT",
        "descricao": "Finance General REPORT",
        "linha_funcao": 24,
        "linha_transacao": 35,
    }]

    texto = mod.formatar_relatorio_erro_aplicacao(objeto, roles)

    assert "Transação/aplicação: FAGLB03" in texto
    assert "Objeto recusado: B_BUP_DCPD" in texto
    assert "Atividade: 03 — visualizar" in texto
    assert "Função: Z_ACCOUNTING_REPORT - Descrição: Finance General REPORT" in texto
    assert "linha 24" in texto
    assert "linha 35" in texto


def test_enriquece_objeto_apenas_com_funcoes_da_sheet_proposta(monkeypatch):
    mod = carregar_modulo_su53()
    objeto = {"aplicacao": "FAGLB03", "objeto": "B_BUP_DCPD"}
    monkeypatch.setattr(mod, "obter_descricao_transacao_sap", lambda _tcode: "Exibição de saldos")
    monkeypatch.setattr(
        mod,
        "buscar_funcao_individual_para_transacao",
        lambda *_args, **_kwargs: [
            {"role": "Z_ACCOUNTING_REPORT", "descricao": "Finance General REPORT", "fonte": "Folha 'Proposta'"},
            {"role": "Z_OUTRA", "descricao": "SAP", "fonte": "SAP PRD (AGR_1251)"},
        ],
    )

    resultado = mod.enriquecer_objeto_com_proposta(objeto, "perfis.xlsx", "S80002051")

    assert resultado["nome_aplicacao"] == "Exibição de saldos"
    assert [item["role"] for item in resultado["funcoes_proposta"]] == ["Z_ACCOUNTING_REPORT"]
    assert resultado["funcao_proposta"]["role"] == "Z_ACCOUNTING_REPORT"


def test_cockpit_renderiza_aplicacao_e_funcoes_da_proposta():
    js = (
        PROJECT_ROOT
        / "sap_script_web_cockpit_v2"
        / "web_api"
        / "static"
        / "js"
        / "cockpit.agent.js"
    ).read_text(encoding="utf-8")

    assert "o.aplicacao" in js
    assert "o.nome_aplicacao" in js
    assert "o.funcoes_proposta" in js
    assert "linha_funcao" in js
    assert "linha_transacao" in js


def test_separa_blocos_de_objetos_com_linha_pontilhada(monkeypatch, capsys):
    mod = carregar_modulo_su53()
    objetos = [
        {
            "aplicacao": "ME23N",
            "objeto": "M_EINK_FRG",
            "descricao": "Compras",
            "rc_mais_recente": 4,
            "data_hora_recente": "23.10.2026 08:34:48",
            "programa": "PROG1",
            "campos_relevantes": [("FRGGR", "K7")],
            "total_ocorrencias": 12,
        },
        {
            "aplicacao": "FAGLB03",
            "objeto": "B_BUP_DCPD",
            "descricao": "Business Partner",
            "rc_mais_recente": 4,
            "data_hora_recente": "23.10.2026 08:24:08",
            "programa": "PROG2",
            "campos_relevantes": [("ACTVT", "03")],
            "total_ocorrencias": 10,
        },
    ]
    monkeypatch.setattr(
        mod,
        "enriquecer_objeto_com_proposta",
        lambda objeto, *_args: {**objeto, "nome_aplicacao": "Aplicação", "funcoes_proposta": []},
    )

    mod.exibir_analise_sintese([], objetos, "perfis.xlsx", "S80002051")
    saida = capsys.readouterr().out

    assert "\n\n...............................................................................................\n\n  [2] Objeto:" in saida
    assert "[O] Introduzir manualmente outra Transação ou Objeto" not in saida
    assert "[0] Cancelar e sair" not in saida


def test_prepara_atualizacao_objeto_para_todas_funcoes_indicadas():
    mod = carregar_modulo_su53()
    objeto = {
        "objeto": "M_EINK_FRG",
        "campos_relevantes": [
            ("FRGGR", "K7"),
            ("FRGCO", "20"),
            ("FRGCO", "10"),
            ("FRGCO", "20"),
        ],
        "funcoes_proposta": [
            {"role": "Z_PURCHASE_ORDER_CREATE"},
            {"role": "Z_PURCHASE_ORDER_DISPLAY"},
        ],
    }

    plano = mod.preparar_atualizacao_objeto(objeto)

    assert plano == {
        "objeto": "M_EINK_FRG",
        "roles": ["Z_PURCHASE_ORDER_CREATE", "Z_PURCHASE_ORDER_DISPLAY"],
        "campos": {"FRGGR": "K7", "FRGCO": "20,10"},
    }


def test_prepara_atualizacao_objeto_recusa_sem_funcao_ou_valores():
    mod = carregar_modulo_su53()

    assert mod.preparar_atualizacao_objeto({"objeto": "B_BUP_DCPD"}) is None


def test_ativa_scripting_qad_via_rfc_e_confirma_leitura(monkeypatch):
    mod = carregar_modulo_su53()

    class FakeConnection:
        def __init__(self, **_kwargs):
            self.valor = "FALSE"

        def call(self, nome, **kwargs):
            if nome == "TH_GET_PARAMETER":
                return {"PARAMETER_VALUE": self.valor, "RC": 0}
            assert nome == "TH_CHANGE_PARAMETER"
            assert kwargs == {
                "PARAMETER_NAME": "sapgui/user_scripting",
                "PARAMETER_VALUE": "TRUE",
                "CHECK_PARAMETER": 1,
            }
            self.valor = "TRUE"
            return {}

        def close(self):
            pass

    monkeypatch.setattr(mod, "load_project_env", lambda *_args: None)
    monkeypatch.setattr(mod, "build_connection_params_for", lambda env: {"env": env})
    monkeypatch.setattr(mod, "Connection", FakeConnection)

    assert mod.garantir_scripting_gui_qad_rfc() == "TRUE"


def test_classifica_objeto_existente_sem_criar_nova_autorizacao():
    mod = carregar_modulo_su53()

    assert mod.classificar_modo_objeto([]) == "CRIAR"
    assert mod.classificar_modo_objeto([{"DELETED": "X"}]) == "CRIAR"
    assert mod.classificar_modo_objeto([{"DELETED": ""}]) == "ATUALIZAR"
    assert mod.classificar_modo_objeto([{"DELETED": "X"}, {"DELETED": ""}]) == "ATUALIZAR"


def test_asterisco_satisfaz_valor_especifico_na_validacao():
    mod = carregar_modulo_su53()

    assert mod.valor_autorizacao_cobre("*", "03") is True
    assert mod.valor_autorizacao_cobre("03", "03") is True
    assert mod.valor_autorizacao_cobre("02", "03") is False


def test_executa_todos_objetos_com_confirmacao_unica_e_continua_apos_falha(monkeypatch, capsys):
    mod = carregar_modulo_su53()
    objetos = [
        {"aplicacao": "ZSD_PACKING_LIST", "objeto": "S_ALV_LAYR"},
        {"aplicacao": "MB51", "objeto": "S_ALV_LAYO"},
        {"aplicacao": "VL10B", "objeto": "S_ALV_LAYR"},
    ]
    respostas = iter(["S"])
    chamadas = []

    monkeypatch.setattr("builtins.input", lambda _prompt: next(respostas))

    def executar(**kwargs):
        chamadas.append(kwargs)
        return kwargs["obj_info"]["aplicacao"] != "MB51"

    monkeypatch.setattr(mod, "diagnosticar_objeto_autorizacao", executar)

    resultado = mod.executar_todos_objetos_autorizacao(
        usuario="S80000126",
        objetos=objetos,
        caminho_excel="perfis.xlsx",
        target_env="QAD",
    )

    assert [item["obj_info"] for item in chamadas] == objetos
    assert all(item["pedir_confirmacao"] is False for item in chamadas)
    assert resultado == {"sucessos": 2, "falhas": 1}
    saida = capsys.readouterr().out
    assert "2 sucesso(s) | 1 falha(s)" in saida


def test_cancela_execucao_de_todos_objetos_sem_processar(monkeypatch):
    mod = carregar_modulo_su53()
    chamadas = []

    monkeypatch.setattr("builtins.input", lambda _prompt: "N")
    monkeypatch.setattr(mod, "diagnosticar_objeto_autorizacao", lambda **kwargs: chamadas.append(kwargs))

    resultado = mod.executar_todos_objetos_autorizacao(
        usuario="S80000126",
        objetos=[{"aplicacao": "MB51", "objeto": "S_ALV_LAYO"}],
        caminho_excel="perfis.xlsx",
        target_env="PRD",
    )

    assert resultado == {"sucessos": 0, "falhas": 0}
    assert chamadas == []


def criar_excel_departamentos(caminho):
    wb = openpyxl.Workbook()
    controlo = wb.active
    controlo.title = "CONTROLO"
    controlo.append(["DEPARTAMENTO", "STATUS"])
    controlo.append(["Product Distribution", "PROCESSADO"])
    controlo.append(["Legal", "PROCESSADO"])

    ativa = wb.create_sheet("Proposta Ativa")
    ativa.append(["Usuário", "NºPessoa", "Nome Completo", "Email", "Departamento/Direção", "Função", "Departamento"])
    ativa.append(["S80000126", None, "Joana Faria", None, None, None, "Product Distribution"])
    ativa.append(["S6617", None, "Ana Torcato", None, None, None, " product distribution "])
    ativa.append(["S80001813", None, "Inês Vieira", None, None, None, "Legal"])
    ativa.append([None, None, "Sem utilizador", None, None, None, "Legal"])
    wb.save(caminho)


def test_lista_departamentos_controlo_com_contagem_da_proposta_ativa(tmp_path):
    mod = carregar_modulo_su53()
    caminho = tmp_path / "perfis.xlsx"
    criar_excel_departamentos(caminho)

    departamentos = mod.listar_departamentos_controlo(str(caminho))

    assert departamentos == [
        {"departamento": "Product Distribution", "status": "PROCESSADO", "utilizadores": 2},
        {"departamento": "Legal", "status": "PROCESSADO", "utilizadores": 1},
    ]


def test_lista_utilizadores_do_departamento_com_nome_e_sem_duplicados(tmp_path):
    mod = carregar_modulo_su53()
    caminho = tmp_path / "perfis.xlsx"
    criar_excel_departamentos(caminho)

    utilizadores = mod.listar_utilizadores_departamento(str(caminho), "PRODUCT DISTRIBUTION")

    assert utilizadores == [
        {"usuario": "S80000126", "nome": "Joana Faria"},
        {"usuario": "S6617", "nome": "Ana Torcato"},
    ]


def test_pesquisa_departamento_continua_apos_falha_rfc(monkeypatch, capsys):
    mod = carregar_modulo_su53()
    utilizadores = [
        {"usuario": "S1", "nome": "Com erro"},
        {"usuario": "S2", "nome": "Sem erro"},
        {"usuario": "S3", "nome": "RFC falhou"},
    ]

    def consultar(usuario, target_env):
        assert target_env == "PRD"
        if usuario == "S3":
            raise RuntimeError("ligação indisponível")
        return [{"usuario": usuario}] if usuario == "S1" else []

    monkeypatch.setattr(mod, "consultar_erros_su53_rfc", consultar)
    monkeypatch.setattr(
        mod,
        "agrupar_erros_para_analise",
        lambda erros: ([], [{"aplicacao": "MB51", "objeto": "S_ALV_LAYR", "campos_relevantes": [("ACTVT", "23")]}]) if erros else ([], []),
    )
    monkeypatch.setattr(
        mod,
        "enriquecer_objeto_com_proposta",
        lambda objeto, *_args: {**objeto, "funcoes_proposta": [{"role": "Z_STOCK_REPORT"}]},
    )

    resultado = mod.pesquisar_departamento_su53(
        departamento="Product Distribution",
        utilizadores=utilizadores,
        caminho_excel="perfis.xlsx",
        target_env="PRD",
    )

    assert [item["usuario"] for item in resultado["com_erros"]] == ["S1"]
    assert [item["usuario"] for item in resultado["sem_erros"]] == ["S2"]
    assert resultado["falhas_rfc"] == [{"usuario": "S3", "nome": "RFC falhou", "erro": "ligação indisponível"}]
    saida = capsys.readouterr().out
    assert "Z_STOCK_REPORT" in saida
    assert "1 com erro(s) | 1 sem erros | 1 falha(s) RFC" in saida


def test_pesquisa_departamento_continua_apos_falha_na_analise(monkeypatch):
    mod = carregar_modulo_su53()
    utilizadores = [
        {"usuario": "S1", "nome": "Análise falhou"},
        {"usuario": "S2", "nome": "Sem erro"},
    ]

    monkeypatch.setattr(
        mod,
        "consultar_erros_su53_rfc",
        lambda usuario, target_env: [{"usuario": usuario}] if usuario == "S1" else [],
    )
    monkeypatch.setattr(
        mod,
        "agrupar_erros_para_analise",
        lambda _erros: (_ for _ in ()).throw(RuntimeError("descrição RFC indisponível")),
    )

    resultado = mod.pesquisar_departamento_su53(
        departamento="Product Distribution",
        utilizadores=utilizadores,
        caminho_excel="perfis.xlsx",
        target_env="PRD",
    )

    assert [item["usuario"] for item in resultado["sem_erros"]] == ["S2"]
    assert resultado["falhas_rfc"][0]["usuario"] == "S1"
    assert resultado["falhas_rfc"][0]["erro"] == "descrição RFC indisponível"


def test_resumo_departamental_mostra_apenas_funcoes_da_sheet_proposta(monkeypatch, capsys):
    mod = carregar_modulo_su53()
    monkeypatch.setattr(mod, "consultar_erros_su53_rfc", lambda *_args, **_kwargs: [{"erro": True}])
    monkeypatch.setattr(
        mod,
        "agrupar_erros_para_analise",
        lambda _erros: ([{"tcode_falha": "MB51"}], []),
    )
    monkeypatch.setattr(
        mod,
        "buscar_funcao_individual_para_transacao",
        lambda *_args, **_kwargs: [
            {"role": "Z_STOCK_REPORT", "fonte": "Folha 'Proposta'"},
            {"role": "Z_ROLE_SAP", "fonte": "SAP PRD (AGR_1251)"},
        ],
    )

    mod.pesquisar_departamento_su53(
        departamento="Product Distribution",
        utilizadores=[{"usuario": "S1", "nome": "Utilizador"}],
        caminho_excel="perfis.xlsx",
        target_env="PRD",
    )

    saida = capsys.readouterr().out
    assert "Z_STOCK_REPORT" in saida
    assert "Z_ROLE_SAP" not in saida


def test_main_permite_pesquisar_departamento_sem_entrar_no_fluxo_de_atualizacao(monkeypatch, tmp_path, capsys):
    mod = carregar_modulo_su53()
    caminho = tmp_path / "perfis.xlsx"
    criar_excel_departamentos(caminho)
    respostas = iter(["2", "1", "1", "0"])
    chamadas = []

    monkeypatch.setattr(mod, "CAMINHO_EXCEL_PADRAO", str(caminho))
    monkeypatch.setattr(mod.sys, "argv", ["Pesquisa_Erros_SU53.py"])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(respostas))
    monkeypatch.setattr(
        mod,
        "pesquisar_departamento_su53",
        lambda **kwargs: chamadas.append(kwargs) or {
            "com_erros": [{"usuario": "S80000126", "nome": "Joana Faria"}],
            "sem_erros": [{"usuario": "S6617", "nome": "Ana Torcato"}],
            "falhas_rfc": [],
        },
    )
    monkeypatch.setattr(
        mod,
        "consultar_erros_su53_rfc",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(AssertionError("não deve consultar detalhe individual")),
    )

    mod.main()

    assert chamadas[0]["departamento"] == "Product Distribution"
    assert [item["usuario"] for item in chamadas[0]["utilizadores"]] == ["S80000126", "S6617"]
    assert chamadas[0]["target_env"] == "PRD"
    saida = capsys.readouterr().out
    assert "[2] Departamento" in saida
    assert "Product Distribution" in saida
    assert "Selecionar um utilizador com erros" in saida


def opcoes_mistas_su53():
    return [
        {"tipo": "TRANSACAO", "valor": "FDKUSER", "detalhes": {"tcode_falha": "FDKUSER"}},
        {"tipo": "TRANSACAO", "valor": "MM03", "detalhes": {"tcode_falha": "MM03"}},
        {"tipo": "OBJETO", "valor": "F_UKM_SGMT", "detalhes": {"objeto": "F_UKM_SGMT", "aplicacao": "APP1"}},
        {"tipo": "OBJETO", "valor": "T_ADMIN", "detalhes": {"objeto": "T_ADMIN", "aplicacao": "VF03"}},
    ]


def preparar_main_misto(monkeypatch, mod, caminho, escolha):
    respostas = iter(escolha if isinstance(escolha, list) else [escolha])
    monkeypatch.setattr(mod, "CAMINHO_EXCEL_PADRAO", str(caminho))
    monkeypatch.setattr(mod.sys, "argv", ["Pesquisa_Erros_SU53.py", "S5092", "PRD"])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(respostas))
    monkeypatch.setattr(mod, "consultar_erros_su53_rfc", lambda *_args, **_kwargs: [{"erro": True}])
    monkeypatch.setattr(mod, "exibir_tabela_erros", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        mod,
        "agrupar_erros_para_analise",
        lambda _erros: ([{"tcode_falha": "FDKUSER"}, {"tcode_falha": "MM03"}], [{}, {}]),
    )
    monkeypatch.setattr(mod, "exibir_analise_sintese", lambda *_args, **_kwargs: opcoes_mistas_su53())


def test_menu_misto_permite_selecionar_quarto_item_objeto(monkeypatch, tmp_path, capsys):
    mod = carregar_modulo_su53()
    caminho = tmp_path / "perfis.xlsx"
    criar_excel_departamentos(caminho)
    chamadas = []
    preparar_main_misto(monkeypatch, mod, caminho, ["4", "0"])
    monkeypatch.setattr(mod, "diagnosticar_objeto_autorizacao", lambda **kwargs: chamadas.append(kwargs))

    mod.main()

    assert chamadas[0]["obj_info"]["objeto"] == "T_ADMIN"
    saida = capsys.readouterr().out
    assert "[4] Objeto T_ADMIN" in saida


def test_menu_misto_opcao_b_processa_todos_os_objetos(monkeypatch, tmp_path):
    mod = carregar_modulo_su53()
    caminho = tmp_path / "perfis.xlsx"
    criar_excel_departamentos(caminho)
    chamadas = []
    preparar_main_misto(monkeypatch, mod, caminho, "B")
    monkeypatch.setattr(mod, "executar_todos_objetos_autorizacao", lambda **kwargs: chamadas.append(kwargs))
    monkeypatch.setattr(
        mod,
        "diagnosticar_objeto_autorizacao",
        lambda **_kwargs: (_ for _ in ()).throw(AssertionError("não deve processar apenas o primeiro objeto")),
    )

    mod.main()

    assert [item["objeto"] for item in chamadas[0]["objetos"]] == ["F_UKM_SGMT", "T_ADMIN"]

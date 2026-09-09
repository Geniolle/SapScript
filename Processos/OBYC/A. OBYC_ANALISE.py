# -*- coding: utf-8 -*-

###################################################################################
# BLOCO: ANÁLISE OBYC (determinação automática de contas)
# Traz para o cockpit de terminal a mesma análise read-only da OBYC que já
# existia apenas no cockpit Web. Usa sempre RFC_READ_TABLE (nunca escreve em
# SAP) através de sap_rfc/obyc_service.py — o mesmo motor já usado pelo
# cockpit Web. Não é preciso sessão SAP GUI aberta: a leitura é feita por RFC
# diretamente e, se o processo atual não tiver o pyrfc instalado (venv
# "normal" do cockpit), o próprio serviço recorre automaticamente à
# .venv-rfc através de um bridge por subprocesso.
#
# Dois modos disponíveis:
#   1. Ficheiro Excel  - valida em massa as linhas de um Excel contra a
#                        configuração OBYC atual em SAP.
#   2. Individual      - consulta pontual por chave (KTOPL/KTOSL/BWMOD/
#                        KOMOK/BKLAS/...).
###################################################################################

OBYC_TABELAS = ("T030", "T030K", "T030R", "T030B", "T030H")


def executar(ambiente_cockpit, **kwargs):
    import os
    import sys
    import json
    from datetime import datetime

    # --- CORREÇÃO DA ESTRUTURA DE PASTAS PARA IMPORTAR sap_rfc/obyc_service ---
    dir_atual = os.path.dirname(os.path.abspath(__file__))
    dir_processos = os.path.dirname(dir_atual)
    dir_projeto = os.path.dirname(dir_processos)
    for p in (dir_processos, dir_projeto):
        if p not in sys.path:
            sys.path.insert(0, p)

    def agora_ts():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def log(msg):
        print(f"{agora_ts()} | {msg}", flush=True)

    try:
        from sap_rfc.obyc_service import (
            read_obyc_table,
            validate_obyc_excel,
            _load_excel_rows_for_validation,
            OBYC_VALIDATION_PRIMARY_FIELDS,
        )
    except Exception as e:
        log(f"❌ Não foi possível importar o serviço OBYC (sap_rfc/obyc_service.py): {e}")
        return "voltar"

    print("\n📊 Tabelas OBYC disponíveis:")
    for i, t in enumerate(OBYC_TABELAS, 1):
        print(f"  {i}. {t}")

    tabela = None
    while tabela is None:
        escolha = input(f"\nEscolha a tabela [1-{len(OBYC_TABELAS)}] (Enter = T030): ").strip()
        if not escolha:
            tabela = "T030"
            break
        try:
            idx = int(escolha) - 1
            if 0 <= idx < len(OBYC_TABELAS):
                tabela = OBYC_TABELAS[idx]
        except ValueError:
            pass
        if tabela is None:
            print("❌ Opção inválida.")

    print("\nModo de análise:")
    print("  1. Ficheiro Excel (validação em massa)")
    print("  2. Individual (consulta pontual por chave)")
    modo = input("Escolha o modo [1-2] (Enter = 1): ").strip()

    if modo == "2":
        _consulta_individual(ambiente_cockpit, tabela, read_obyc_table, OBYC_VALIDATION_PRIMARY_FIELDS, log)
        return "voltar"

    ###################################################################################
    # MODO: FICHEIRO EXCEL
    ###################################################################################
    caminho_ficheiro = _selecionar_ficheiro_excel(log)
    if not caminho_ficheiro:
        log("❌ Operação cancelada (ficheiro não selecionado).")
        return "voltar"

    log(f"📂 Ficheiro: {caminho_ficheiro}")
    log("⏳ A ler ficheiro Excel...")
    try:
        workbook = _load_excel_rows_for_validation(caminho_ficheiro)
    except Exception as e:
        log(f"❌ Erro ao ler o Excel: {e}")
        return "voltar"

    log(f"📋 Linhas úteis encontradas: {workbook.get('row_count', 0)}")
    if not workbook.get("rows"):
        log("❌ Nenhuma linha útil encontrada no ficheiro.")
        return "voltar"

    log(f"⏳ A validar contra a tabela {tabela} em {ambiente_cockpit} via RFC (read-only)...")
    try:
        resultado_json, resumo_log = validate_obyc_excel(
            {
                "preview_data": workbook,
                "system": ambiente_cockpit,
                "table": tabela,
            }
        )
    except Exception as e:
        log(f"❌ Erro na validação OBYC: {e}")
        return "voltar"

    try:
        resultado = json.loads(resultado_json)
    except Exception:
        resultado = {}

    print(f"\n{resumo_log}")

    issues = resultado.get("issues") or []
    if issues:
        print(f"\n⚠️ Divergências / linhas sem correspondência (até 20):")
        for it in issues:
            linha_num = it.get("row_number")
            motivo = it.get("message") or it.get("reason")
            print(f"  Linha {linha_num}: {motivo}")

    log("✅ Análise OBYC concluída.")
    return "voltar"


def _selecionar_ficheiro_excel(log):
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        caminho = filedialog.askopenfilename(
            title="Selecione o ficheiro Excel da OBYC",
            filetypes=(("Ficheiros Excel", "*.xlsx;*.xlsm"), ("Todos os ficheiros", "*.*")),
        )
        root.destroy()
        return caminho or ""
    except Exception as e:
        log(f"❌ Falha ao abrir popup de ficheiro: {e}")
        return ""


def _consulta_individual(ambiente_cockpit, tabela, read_obyc_table, campos_sugeridos, log):
    print(f"\n🔎 Consulta individual em {tabela} ({ambiente_cockpit}).")
    print("Informe os valores das chaves que quiser filtrar (Enter para ignorar um campo).")
    print(f"Campos sugeridos: {', '.join(campos_sugeridos)}")

    filtros = []
    for campo in campos_sugeridos:
        valor = input(f"  {campo} = ").strip()
        if valor:
            filtros.append({"field": campo, "value": valor})

    if not filtros:
        campo_extra = input(
            "Nenhuma chave sugerida preenchida. Nome de outro campo (Enter para cancelar): "
        ).strip().upper()
        if not campo_extra:
            log("❌ Consulta cancelada (nenhuma chave informada).")
            return
        valor_extra = input(f"  {campo_extra} = ").strip()
        if not valor_extra:
            log("❌ Consulta cancelada (nenhuma chave informada).")
            return
        filtros.append({"field": campo_extra, "value": valor_extra})

    log(f"⏳ A consultar {tabela} via RFC (read-only)...")
    try:
        linhas = read_obyc_table(ambiente_cockpit, tabela, filtros, [])
    except Exception as e:
        log(f"❌ Erro na consulta OBYC: {e}")
        return

    if not linhas:
        log(f"ℹ️ Nenhum registo encontrado em {tabela} para os filtros informados.")
        return

    log(f"✅ {len(linhas)} registo(s) encontrado(s):")
    campos = sorted({campo for linha in linhas for campo in linha.keys()})
    for i, linha in enumerate(linhas, 1):
        print(f"\n  Registo {i}:")
        for campo in campos:
            valor = linha.get(campo, "")
            if valor:
                print(f"    {campo}: {valor}")

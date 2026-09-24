# -*- coding: utf-8 -*-

###################################################################################
# BLOCO: VALIDAR FLUXO DE CAIXA / ITEM DE LIQUIDEZ (New Cash Management)
# Apoia a investigação read-only de documentos financeiros que recebem um
# Item de Liquidez (LQA...) incorreto ou caem no item técnico de eliminação
# (LQAT9999). Usa sempre RFC_READ_TABLE (nunca escreve em SAP) através de
# sap_rfc/cash_flow_validation_service.py.
#
# Contexto: SKB1-FIPOS não está em uso neste sistema. A classificação real
# vem do New Cash Management (FQM) e é calculada em runtime — não é possível
# lê-la diretamente por RFC_READ_TABLE. Este processo traz o documento e a
# respetiva cadeia de compensação (AUGBL) para permitir "subir até a origem"
# manualmente, tal como o motor de derivação (FQM_CLEARING /
# FQMS_LIQUPOS_DERIV) deveria fazer.
#
# Ver docs/CONHECIMENTO_FLUXO_DE_CAIXA_LIQUIDEZ.md para o contexto completo
# desta investigação (caso de origem: documento 8020024067, empresa 2010).
###################################################################################


def executar(ambiente_cockpit, **kwargs):
    import os
    import sys
    from datetime import datetime

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
        from sap_rfc.cash_flow_validation_service import (
            fetch_clearing_group,
            fetch_document,
            lookup_liquidity_item,
        )
    except Exception as e:
        log(f"❌ Não foi possível importar o serviço (sap_rfc/cash_flow_validation_service.py): {e}")
        return "voltar"

    print("\n💰 Validar Fluxo de Caixa / Item de Liquidez (read-only)")
    bukrs = input("Empresa (BUKRS), ex. 2010: ").strip()
    belnr = input("Nº do documento (BELNR): ").strip()
    gjahr = input("Exercício (GJAHR), ex. 2025: ").strip()

    if not (bukrs and belnr and gjahr):
        log("❌ Empresa, documento e exercício são obrigatórios.")
        return "voltar"

    log(f"⏳ A ler documento {bukrs}/{belnr}/{gjahr} em {ambiente_cockpit} via RFC (read-only)...")
    try:
        doc = fetch_document(ambiente_cockpit, bukrs, belnr, gjahr)
    except Exception as e:
        log(f"❌ Erro ao ler o documento: {e}")
        return "voltar"

    if not doc.get("found"):
        log("ℹ️ Documento não encontrado.")
        return "voltar"

    hdr = doc["header"]
    print(f"\n📄 Cabeçalho: BLART={hdr.get('BLART')}  BLDAT={hdr.get('BLDAT')}  BUDAT={hdr.get('BUDAT')}"
          f"  TCODE={hdr.get('TCODE')}  USNAM={hdr.get('USNAM')}")
    print(f"   Texto: {hdr.get('BKTXT') or '(vazio)'}")

    print("\n📋 Itens:")
    augbl_vistos = set()
    for linha in doc["lines"]:
        print(
            f"   Item {linha.get('BUZEI')}: conta {linha.get('HKONT')} ({linha.get('HKONT_DESC') or 's/ descrição'})"
            f" | {linha.get('SHKZG')} {linha.get('DMBTR')} | BSCHL {linha.get('BSCHL')}"
            f" | texto: {linha.get('SGTXT') or '(vazio)'}"
            f" | compensado por AUGBL={linha.get('AUGBL') or '-'} em {linha.get('AUGDT') or '-'}"
        )
        if linha.get("AUGBL"):
            augbl_vistos.add(linha["AUGBL"])

    if not augbl_vistos:
        log("ℹ️ Nenhum documento em aberto/compensado (sem AUGBL). Fim da investigação.")
        return "voltar"

    for augbl in sorted(augbl_vistos):
        print(f"\n🔗 Grupo de compensação AUGBL={augbl} (todos os documentos fechados por esta compensação):")
        try:
            grupo = fetch_clearing_group(ambiente_cockpit, bukrs, augbl, gjahr)
        except Exception as e:
            log(f"   ❌ Erro ao ler o grupo de compensação: {e}")
            continue

        if not grupo:
            print("   (nenhum item encontrado)")
            continue

        for linha in grupo:
            marcador = "⭐" if linha.get("BELNR") == belnr.zfill(10) else "  "
            print(
                f"   {marcador} Doc {linha.get('BELNR')} item {linha.get('BUZEI')}: conta {linha.get('HKONT')}"
                f" ({linha.get('HKONT_DESC') or 's/ descrição'}) | {linha.get('SHKZG')} {linha.get('DMBTR')}"
                f" | texto: {linha.get('SGTXT') or '(vazio)'}"
            )

        contas_distintas = {l.get("HKONT") for l in grupo}
        textos_distintos = {l.get("SGTXT") for l in grupo if l.get("SGTXT")}
        if len(contas_distintas) == 1:
            print("   ⚠️  Todas as linhas compensadas caem na MESMA conta — compensação sem contrapartida"
                  " noutra conta (potencial causa de item de liquidez técnico/eliminação).")
        if len(textos_distintos) > 1:
            print(f"   ⚠️  Textos diferentes no mesmo grupo de compensação: {sorted(textos_distintos)}"
                  " — indício de origens de negócio distintas a serem fechadas juntas.")

    print("\n🔎 Consultar significado de um código de Item de Liquidez (ex.: LQAO0141). Enter para saltar.")
    codigo = input("Código: ").strip()
    if codigo:
        try:
            info = lookup_liquidity_item(ambiente_cockpit, codigo)
        except Exception as e:
            log(f"❌ Erro ao consultar o item de liquidez: {e}")
            info = None
        if info:
            print(f"   {info['LIQUIDITYITEM']}: {info['LIQUIDITYITEMNAME']}")
        else:
            print("   Código não encontrado em ALIQITEMTXT.")

    log("✅ Validação de fluxo de caixa concluída (read-only).")
    return "voltar"

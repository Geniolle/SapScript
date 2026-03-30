###################################################################################
# BLOCO 1: IMPORTAÇÕES E DEFINIÇÕES INICIAIS
###################################################################################
def executar(ambiente_cockpit):
    import os
    import pandas as pd
    import win32com.client
    import time

    tempo_inicio = time.time()
    mapa_sistema = {"DEV": "S4D", "QAD": "S4Q", "PRD": "S4P"}
    sistema_desejado = mapa_sistema.get(ambiente_cockpit)

    caminho_pasta = r"C:\SAP Script\Processos\Reverter Documento"
    ficheiro_nome = "Script_Anular_Compensação_FBRA.xlsx"
    caminho_ficheiro = os.path.join(caminho_pasta, ficheiro_nome)

###################################################################################
# BLOCO 2: LEITURA DO FICHEIRO EXCEL
###################################################################################
    try:
        df = pd.read_excel(caminho_ficheiro)
    except Exception as e:
        print(f"❌ Erro ao abrir o ficheiro '{ficheiro_nome}': {e}")
        return

    if 'MSG' not in df.columns:
        df['MSG'] = ''
    df['MSG'] = df['MSG'].astype(str)
    df['STATUS'] = df['STATUS'].fillna('').astype(str).str.strip().str.upper()

    for col in ['ID', 'Nº DOCUMENTO', 'EMPRESA', 'EXERCÍCIO', 'MOTIVO', 'DATA DO LANÇAMENTO', 'STATUS']:
        if col not in df.columns:
            print(f"❌ Coluna obrigatória em falta: {col}")
            return

    df_filtrado = df[(df['ID'].notna()) & (df['STATUS'] != 'CONCLUÍDO')].sort_values(by='ID')
    if df_filtrado.empty:
        print("\n⚠️ Nenhuma linha válida encontrada.")
        return

    print(f"\n📋 Documentos a anular ({len(df_filtrado)}):")
    print(df_filtrado.fillna('').to_string(index=False))

    resposta = input("\nDeseja anular estas compensações no SAP? [S/N]: ").strip().upper()
    if resposta != 'S':
        print("❌ Anulação cancelada pelo utilizador.")
        return

###################################################################################
# BLOCO 3: CONEXÃO AO SAP GUI
###################################################################################
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        session = None

        for conn in application.Children:
            for sess in conn.Children:
                try:
                    if sess.Info.SystemName.upper() == sistema_desejado:
                        session = sess
                        break
                except:
                    continue
            if session:
                break

        if not session:
            print(f"\n❌ Nenhuma sessão encontrada para o ambiente '{ambiente_cockpit}' (esperado: {sistema_desejado}).")
            return

        print(f"\n✅ Conectado ao sistema SAP: {session.Info.SystemName} (ambiente {ambiente_cockpit})")
        print(f"👤 Utilizador SAP: {session.Info.User} | Cliente: {session.Info.Client}")

    except Exception as e:
        print(f"❌ Erro ao conectar SAP: {e}")
        return

    session.findById("wnd[0]").resizeWorkingPane(92, 28, False)

###################################################################################
# BLOCO 4: ANULAÇÃO DAS COMPENSAÇÕES
###################################################################################
    total = len(df_filtrado)
    for i, (index, row) in enumerate(df_filtrado.iterrows(), start=1):
        doc_num = str(row['Nº DOCUMENTO']).strip()
        empresa = str(row['EMPRESA']).strip()
        exercicio = str(row['EXERCÍCIO']).strip()

        try:
            motivo = str(int(float(row['MOTIVO']))).zfill(2)
        except:
            motivo = str(row['MOTIVO']).strip().zfill(2)

        try:
            data_lancamento = pd.to_datetime(row['DATA DO LANÇAMENTO'], dayfirst=True).strftime("%d.%m.%Y")
        except:
            data_lancamento = ""

        print(f"\n🔧 {i}/{total} - Anulando compensação do documento {doc_num} - Empresa {empresa}")
        print(f"🔎 MOTIVO formatado: {motivo}")
        print(f"🔎 DATA LANÇAMENTO formatada: {data_lancamento}")

        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/NFBRA"
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/usr/txtRF05R-AUGBL").text = doc_num
            session.findById("wnd[0]/usr/ctxtRF05R-BUKRS").text = empresa
            session.findById("wnd[0]/usr/txtRF05R-GJAHR").text = exercicio
            session.findById("wnd[0]/usr/txtRF05R-GJAHR").setFocus
            session.findById("wnd[0]/usr/txtRF05R-GJAHR").caretPosition = 4
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]").sendVKey(0)
            #session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[1]/usr/btnSPOP-VAROPTION2").press()
            session.findById("wnd[1]/usr/ctxtRF05R-STGRD").text = motivo
            session.findById("wnd[1]/usr/ctxtRF05R-BUDAT").text = data_lancamento
            session.findById("wnd[1]/usr/ctxtRF05R-BUDAT").setFocus
            session.findById("wnd[1]/usr/ctxtRF05R-BUDAT").caretPosition = 10
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]").sendVKey(0)
            #session.findById("wnd[0]").sendVKey(0)
            session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
            #session.findById("wnd[0]").sendVKey(0)

            # Captura da última mensagem do SAP
            try:
                msg_final = session.findById("wnd[0]/sbar").Text.strip()
            except:
                msg_final = "Compensação anulada com sucesso (sem mensagem)"

            print(f"📢 Mensagem SAP: {msg_final}")

            df.at[index, 'STATUS'] = 'Concluído'
            df.at[index, 'MSG'] = msg_final
            print(f"✅ Sucesso: {msg_final}")

        except Exception as e:
            df.at[index, 'STATUS'] = 'Erro no processamento'
            try:
                msg_erro = session.findById("wnd[0]/sbar").Text.strip()
            except:
                msg_erro = f"Erro desconhecido: {str(e)}"
            df.at[index, 'MSG'] = msg_erro
            print(f"❌ Erro ao anular compensação do documento {doc_num}. Motivo: {msg_erro}")

###################################################################################
# BLOCO 5: GRAVAÇÃO DO RESULTADO
###################################################################################
    try:
        df.to_excel(caminho_ficheiro, index=False)
        print("💾 Ficheiro atualizado.")
    except Exception as erro:
        print(f"❌ Erro ao salvar ficheiro: {erro}")

###################################################################################
# BLOCO 6: RESUMO FINAL DE ERROS
###################################################################################
    erros = df[df['STATUS'] != 'Concluído']
    if not erros.empty:
        print("\n📌 Resumo final de erros:")
        for _, linha in erros.iterrows():
            doc = linha['Nº DOCUMENTO']
            motivo = linha['MOTIVO']
            msg = linha['MSG']
            print(f"❌ {doc} - {motivo} → {msg}")
    else:
        print("\n🎉 Todas as compensações foram anuladas com sucesso.")

    fim = time.time()
    m, s = divmod(int(fim - tempo_inicio), 60)
    print(f"\n⏱️ Tempo total: {m:02d}:{s:02d}")

###################################################################################
# BLOCO 7: FINALIZAÇÃO E RETORNO AO COCKPIT
###################################################################################
    print("\n🔁 A voltar ao cockpit principal...")
    return True

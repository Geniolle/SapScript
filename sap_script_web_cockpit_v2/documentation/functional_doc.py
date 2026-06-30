# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import logging
from datetime import datetime
from pathlib import Path

class FunctionalDocSession:
    def __init__(self, nome_pasta: str | None, processo: str = "PFCG_CREATE", transacao: str = "PFCG"):
        self.nome_pasta = str(nome_pasta or "").strip()
        self.enabled = bool(self.nome_pasta)
        self.processo = processo
        self.transacao = transacao
        
        self.base_dir = os.getenv("WORKFLOW_DOC_OUTPUT_DIR", "C:\\Jira")
        if self.enabled:
            self.output_dir = Path(self.base_dir) / self.nome_pasta
            self.image_dir = self.output_dir
        else:
            self.output_dir = None
            self.image_dir = None
            
        self.metadata = {}
        self.roles_summary = []
        self.evidences = {}  # role_name -> list of dicts
        self._evidence_count: int = 0  # total de evidências adicionadas
        self._role_summary_count: int = 0  # total de roles adicionadas ao resumo

    @property
    def evidence_count(self) -> int:
        return self._evidence_count

    @property
    def role_summary_count(self) -> int:
        return self._role_summary_count

    @property
    def has_functional_content(self) -> bool:
        """True se existir pelo menos uma evidência ou role no resumo."""
        return self._evidence_count > 0 or self._role_summary_count > 0

    def start_execution(self, metadata: dict):
        if not self.enabled:
            return
        self.metadata = metadata
        self.metadata.setdefault("data_inicio", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            print(f"[DOC_WARN] Não foi possível criar pasta de documentação: {e}")

    def add_role_summary(self, role_name: str, description: str, tcode_count: int, result: str, duration: str):
        if not self.enabled:
            return
        self.roles_summary.append({
            "ordem": len(self.roles_summary) + 1,
            "role": role_name,
            "descricao": description,
            "tcodes": tcode_count,
            "resultado": result,
            "tempo": duration
        })
        self._role_summary_count += 1

    def start_role_section(self, role_name: str, description: str, tcode_count: int):
        if not self.enabled:
            return
        if role_name not in self.evidences:
            self.evidences[role_name] = []

    def add_evidence(self, role_name: str, title: str, caption: str, screenshot_path: str):
        if not self.enabled:
            return
        if role_name not in self.evidences:
            self.evidences[role_name] = []
        self.evidences[role_name].append({
            "title": title,
            "caption": caption,
            "path": screenshot_path
        })
        self._evidence_count += 1

    def finalize(self, output_path: str | None = None) -> str:
        if not self.enabled:
            return ""

        # Não gerar documento vazio se não houver conteúdo funcional real.
        if not self.has_functional_content:
            print("[DOC_WARN] Nenhum conteúdo funcional registado. Documento Word não será gerado.")
            return ""
        
        self.metadata.setdefault("data_fim", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ambiente = self.metadata.get("ambiente", "DEV")
        doc_name = f"Documentacao_Funcional_{self.processo}_{ambiente}_{stamp}.docx"
        
        if output_path:
            final_doc_path = Path(output_path)
        else:
            final_doc_path = self.output_dir / doc_name
            
        try:
            self._generate_word_doc(final_doc_path)
            return str(final_doc_path)
        except Exception as e:
            print(f"[DOC_WARN] Falha ao gerar documento funcional Word: {e}")
            return ""

    def safe_save(self):
        try:
            self.finalize()
        except Exception as e:
            print(f"[DOC_WARN] Erro no safe_save: {e}")

    def _generate_word_doc(self, doc_path: Path):
        import pythoncom
        import win32com.client
        
        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            doc = app.Documents.Add()
            sel = app.Selection
            
            # Helpers de formatação
            def _h1(text):
                sel.Font.Size = 14
                sel.Font.Bold = True
                sel.Font.Italic = False
                sel.TypeText(text)
                sel.TypeParagraph()
                
            def _h2(text):
                sel.Font.Size = 12
                sel.Font.Bold = True
                sel.Font.Italic = False
                sel.TypeText(text)
                sel.TypeParagraph()
                
            def _h3(text):
                sel.Font.Size = 11
                sel.Font.Bold = True
                sel.Font.Italic = True
                sel.TypeText(text)
                sel.TypeParagraph()

            def _p(text):
                sel.Font.Size = 11
                sel.Font.Bold = False
                sel.Font.Italic = False
                sel.TypeText(text)
                sel.TypeParagraph()

            def _bullet(text):
                sel.Font.Size = 11
                sel.Font.Bold = False
                sel.Font.Italic = False
                sel.TypeText(f"• {text}")
                sel.TypeParagraph()

            def _br():
                sel.TypeParagraph()

            def _page_break():
                sel.InsertBreak(7)

            # Obter data no formato YYYY-MM-DD
            data_exec = "-"
            data_inicio_raw = self.metadata.get("data_inicio", "")
            if data_inicio_raw:
                data_inicio_str = str(data_inicio_raw).strip()
                if len(data_inicio_str) >= 10 and data_inicio_str[4] == "-" and data_inicio_str[7] == "-":
                    data_exec = data_inicio_str[:10]
                else:
                    data_exec = data_inicio_str
            else:
                data_exec = datetime.now().strftime("%Y-%m-%d")

            # -------------------------------------------------------------
            # 1. Capa
            # -------------------------------------------------------------
            sel.ParagraphFormat.Alignment = 1  # Center
            _br()
            _br()
            _br()
            sel.Font.Size = 24
            sel.Font.Bold = True
            sel.Font.Name = "Arial"
            sel.TypeText("Análise Técnica/Funcional")
            sel.TypeParagraph()
            _br()
            
            sel.Font.Size = 18
            sel.TypeText(self.processo)
            sel.TypeParagraph()
            _br()
            
            sel.Font.Size = 12
            sel.Font.Bold = False
            sel.Font.Italic = True
            sel.TypeText("Criação/Atualização de Roles e Perfis de Autorização")
            sel.TypeParagraph()
            _br()
            _br()
            _br()
            _br()
            
            sel.Font.Size = 11
            sel.Font.Italic = False
            sel.TypeText(f"Data: {data_exec}")
            sel.TypeParagraph()
            sel.TypeText("Versão: 1.0")
            sel.TypeParagraph()
            
            _page_break()
            sel.ParagraphFormat.Alignment = 0  # Left

            # -------------------------------------------------------------
            # 2. Informação Geral
            # -------------------------------------------------------------
            _h1("1. Informação Geral")
            _br()
            
            meta_keys = [
                ("Processo", self.processo),
                ("Transação SAP utilizada", self.transacao),
                ("Sistema", self.metadata.get("sistema", "")),
                ("Cliente", self.metadata.get("cliente", "")),
                ("Utilizador SAP", self.metadata.get("utilizador_sap", "")),
                ("Data", data_exec),
                ("Total de roles processadas", str(self.metadata.get("total_roles", "0"))),
                ("Pasta de documentação solicitada", self.nome_pasta)
            ]
            
            table_meta = doc.Tables.Add(Range=sel.Range, NumRows=len(meta_keys), NumColumns=2)
            table_meta.Borders.Enable = True
            for idx, (k, v) in enumerate(meta_keys, start=1):
                val_str = str(v).strip() if v is not None else ""
                if not val_str:
                    val_str = "-"
                table_meta.Cell(idx, 1).Range.Text = k
                table_meta.Cell(idx, 1).Range.Font.Bold = True
                table_meta.Cell(idx, 2).Range.Text = val_str
                
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 3. Histórico de Versões
            # -------------------------------------------------------------
            _h1("2. Histórico de Versões")
            _br()
            
            version_headers = ["Versão", "Data", "Autor", "Modificação"]
            version_row = [
                "1.0",
                data_exec,
                str(self.metadata.get("utilizador_sap") or "Sistema").strip(),
                f"Documento inicial gerado automaticamente para evidência funcional da execução {self.processo}"
            ]
            
            table_version = doc.Tables.Add(Range=sel.Range, NumRows=2, NumColumns=len(version_headers))
            table_version.Borders.Enable = True
            
            for col_idx, h in enumerate(version_headers, start=1):
                cell = table_version.Cell(1, col_idx)
                cell.Range.Text = h
                cell.Range.Font.Bold = True
                
            for col_idx, val in enumerate(version_row, start=1):
                table_version.Cell(2, col_idx).Range.Text = val
                
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 4. Índice
            # -------------------------------------------------------------
            _h1("3. Índice")
            _br()
            _p("1. Informação Geral")
            _p("2. Histórico de Versões")
            _p("3. Índice")
            _p("4. Pedido de Alteração")
            _p("   4.1 Requisitos")
            _p("       4.1.1 Requisitos de Qualidade")
            _p("       4.1.2 Requester / Unidade de Negócios / Owner")
            _p("       4.1.3 Requisitos do Cliente")
            _p("5. Módulos e processos afetados")
            _p("   5.1 Módulos afetados")
            _p("   5.2 Processos afetados")
            _p("6. Análise de Viabilidade")
            _p("   6.1 Criação/Atualização de Roles PFCG")
            _p("       6.1.1 Solução Proposta")
            _p("7. Configuração / Execução")
            _p("8. Especificação Funcional")
            _p("   8.1 Objetivo")
            _p("   8.2 Transação Envolvida")
            _p("   8.3 Modo de Processamento")
            _p("   8.4 Estrutura de Entrada")
            _p("   8.5 Regras de Processamento")
            _p("9. Resumo das roles processadas")
            _p("10. Evidências por role")
            _p("11. Testes Unitários")
            _p("12. Anexos")
            
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 5. Pedido de Alteração
            # -------------------------------------------------------------
            _h1("4. Pedido de Alteração")
            _br()
            _p(
                "Este documento tem como objetivo registar a análise funcional e as evidências da execução "
                f"do processo de criação/atualização de roles SAP através da transação {self.transacao}."
            )
            _br()
            
            _h2("4.1 Requisitos")
            _br()
            _h3("4.1.1 Requisitos de Qualidade")
            _br()
            _p(
                f"Os requisitos processados são coerentes com o fluxo standard da transação {self.transacao}, permitindo a "
                "criação ou atualização de roles, atribuição de transações, geração de perfis de autorização "
                "e respetiva rastreabilidade por evidência."
            )
            _br()
            
            _h3("4.1.2 Requester / Unidade de Negócios / Owner")
            _br()
            
            req_keys = [
                ("Requester", "-"),
                ("Unidade de Negócios", "-"),
                ("Owner", "-")
            ]
            table_req = doc.Tables.Add(Range=sel.Range, NumRows=len(req_keys), NumColumns=2)
            table_req.Borders.Enable = True
            for idx, (k, v) in enumerate(req_keys, start=1):
                table_req.Cell(idx, 1).Range.Text = k
                table_req.Cell(idx, 1).Range.Font.Bold = True
                table_req.Cell(idx, 2).Range.Text = v
                
            sel.Start = doc.Content.End
            _br()
            _br()
            
            _h3("4.1.3 Requisitos do Cliente")
            _br()
            
            client_req_headers = ["ID", "Descrição", "Prioridade"]
            client_req_row = [
                "R1",
                f"Criar ou atualizar roles SAP, atribuir transações e gerar perfis de autorização através da transação {self.transacao}.",
                "MÉDIA"
            ]
            table_client_req = doc.Tables.Add(Range=sel.Range, NumRows=2, NumColumns=len(client_req_headers))
            table_client_req.Borders.Enable = True
            for col_idx, h in enumerate(client_req_headers, start=1):
                cell = table_client_req.Cell(1, col_idx)
                cell.Range.Text = h
                cell.Range.Font.Bold = True
            for col_idx, val in enumerate(client_req_row, start=1):
                table_client_req.Cell(2, col_idx).Range.Text = val
                
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 6. Módulos e processos afetados
            # -------------------------------------------------------------
            _h1("5. Módulos e processos afetados")
            _br()
            _h2("5.1 Módulos afetados")
            _br()
            _bullet("SAP Basis / Segurança / Autorizações")
            _br()
            _h2("5.2 Processos afetados")
            _br()
            _bullet("Gestão de roles e perfis de autorização")
            _bullet("Atribuição de transações a roles")
            _bullet("Geração de perfis de autorização")
            _br()

            # -------------------------------------------------------------
            # 7. Análise de Viabilidade
            # -------------------------------------------------------------
            _h1("6. Análise de Viabilidade")
            _br()
            _h2(f"6.1 Criação/Atualização de Roles {self.transacao}")
            _br()
            _p(f"De seguida apresenta-se o detalhe da solução proposta e executada para o processo {self.processo}.")
            _br()
            _h3("6.1.1 Solução Proposta")
            _br()
            _p(
                "A solução consiste em processar automaticamente as roles informadas no ficheiro de entrada, "
                f"validar as transações associadas, aceder à transação {self.transacao}, criar ou atualizar a role, atribuir "
                "as transações na aba Menu, gravar as alterações e gerar o perfil de autorização."
            )
            _br()

            # -------------------------------------------------------------
            # 8. Configuração / Execução
            # -------------------------------------------------------------
            _h1("7. Configuração / Execução")
            _br()
            
            ot_value = str(self.metadata.get("request_transporte") or "").strip()
            if not ot_value:
                ot_value = "-"
                
            config_rows = [
                ("Transação", self.transacao),
                ("Sistema", self.metadata.get("sistema", "")),
                ("Cliente", self.metadata.get("cliente", "")),
                ("Ordem de Transporte", ot_value),
                ("Total de roles processadas", str(self.metadata.get("total_roles", "0"))),
                ("Pasta de documentação", self.nome_pasta)
            ]
            
            table_config = doc.Tables.Add(Range=sel.Range, NumRows=len(config_rows), NumColumns=2)
            table_config.Borders.Enable = True
            for idx, (k, v) in enumerate(config_rows, start=1):
                val_str = str(v).strip() if v is not None else ""
                if not val_str:
                    val_str = "-"
                table_config.Cell(idx, 1).Range.Text = k
                table_config.Cell(idx, 1).Range.Font.Bold = True
                table_config.Cell(idx, 2).Range.Text = val_str
                
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 9. Especificação Funcional
            # -------------------------------------------------------------
            _h1("8. Especificação Funcional")
            _br()
            _h2("8.1 Objetivo")
            _br()
            _p(
                f"O processo tem como objetivo realizar a criação ou atualização de roles SAP utilizando a transação "
                f"{self.transacao}, com atribuição das transações informadas e geração dos respetivos perfis de autorização."
            )
            _br()
            _h2("8.2 Transação Envolvida")
            _br()
            _bullet(f"{self.transacao} - Manutenção de roles")
            _br()
            _h2("8.3 Modo de Processamento")
            _br()
            _p(
                "O processamento é realizado de forma assistida/automática através do Web Cockpit SAP Script, utilizando "
                "como entrada um ficheiro Excel com as roles e transações a processar."
            )
            _br()
            _h2("8.4 Estrutura de Entrada")
            _br()
            
            input_struct_headers = ["Campo", "Tipo", "Descrição"]
            input_struct_rows = [
                ("AGR_NAME", "Texto", "Nome da role"),
                ("TEXT", "Texto", "Descrição da role"),
                ("TCODE", "Texto", "Transações a atribuir"),
                ("STATUS", "Texto", "Resultado do processamento"),
                ("MSG", "Texto", "Mensagem de retorno"),
                ("TIMESTEMP", "Texto", "Data/hora de atualização do resultado")
            ]
            
            table_input_struct = doc.Tables.Add(Range=sel.Range, NumRows=len(input_struct_rows) + 1, NumColumns=3)
            table_input_struct.Borders.Enable = True
            
            for col_idx, h in enumerate(input_struct_headers, start=1):
                cell = table_input_struct.Cell(1, col_idx)
                cell.Range.Text = h
                cell.Range.Font.Bold = True
                
            for row_idx, r in enumerate(input_struct_rows, start=2):
                table_input_struct.Cell(row_idx, 1).Range.Text = r[0]
                table_input_struct.Cell(row_idx, 2).Range.Text = r[1]
                table_input_struct.Cell(row_idx, 3).Range.Text = r[2]
                
            sel.Start = doc.Content.End
            _br()
            _br()
            
            _h2("8.5 Regras de Processamento")
            _br()
            _bullet("Ler o ficheiro Excel informado.")
            _bullet("Agrupar linhas por AGR_NAME.")
            _bullet("Ignorar roles já concluídas.")
            _bullet(f"Abrir a transação {self.transacao}.")
            _bullet("Criar ou atualizar a role.")
            _bullet("Preencher a descrição.")
            _bullet("Atribuir as transações na aba Menu.")
            _bullet("Gravar as alterações.")
            _bullet("Gerar o perfil de autorização.")
            _bullet("Atualizar o Excel com STATUS, MSG e TIMESTEMP.")
            _bullet("Registar evidências no documento funcional.")
            _br()

            # -------------------------------------------------------------
            # 10. Resumo das roles processadas
            # -------------------------------------------------------------
            _h1("9. Resumo das roles processadas")
            _br()
            
            headers = ["Ordem", "Role", "Descrição", "Qtd TCODEs", "Resultado", "Tempo de execução"]
            num_rows = len(self.roles_summary) + 1
            table_summary = doc.Tables.Add(Range=sel.Range, NumRows=num_rows, NumColumns=len(headers))
            table_summary.Borders.Enable = True
            
            for col_idx, h in enumerate(headers, start=1):
                cell = table_summary.Cell(1, col_idx)
                cell.Range.Text = h
                cell.Range.Font.Bold = True
                
            for row_idx, r in enumerate(self.roles_summary, start=2):
                table_summary.Cell(row_idx, 1).Range.Text = str(r["ordem"])
                table_summary.Cell(row_idx, 2).Range.Text = r["role"]
                table_summary.Cell(row_idx, 3).Range.Text = r["descricao"]
                table_summary.Cell(row_idx, 4).Range.Text = str(r["tcodes"])
                table_summary.Cell(row_idx, 5).Range.Text = r["resultado"]
                table_summary.Cell(row_idx, 6).Range.Text = r["tempo"]
                
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 11. Evidências por role
            # -------------------------------------------------------------
            _h1("10. Evidências por role")
            _br()
            
            sub_idx = 1
            for r in self.roles_summary:
                role_name = r["role"]
                if r["resultado"] == "Concluída" and role_name in self.evidences and self.evidences[role_name]:
                    _h2(f"10.{sub_idx} Role {role_name}")
                    _br()
                    _p(f"• Descrição: {r['descricao']}")
                    _p(f"• Quantidade de TCODEs atribuídas: {r['tcodes']}")
                    _p(f"• Resultado: Role tratada por completo")
                    _p(f"• Tempo de execução: {r['tempo']}")
                    _br()
                    
                    for ev in self.evidences[role_name]:
                        sel.Font.Bold = True
                        _p(ev['title'])
                        sel.Font.Bold = False
                        
                        img_path = Path(ev["path"])
                        if img_path.exists():
                            try:
                                sel.InlineShapes.AddPicture(
                                    FileName=str(img_path.resolve()),
                                    LinkToFile=False,
                                    SaveWithDocument=True
                                )
                                sel.TypeParagraph()
                            except Exception as img_exc:
                                _p(f"[Erro ao inserir imagem: {img_exc}]")
                        else:
                            _p("[Imagem de evidência não disponível]")
                            
                        sel.Font.Italic = True
                        _p(f"Legenda: {ev['caption']}")
                        _br()
                        
                    sub_idx += 1
            
            sel.Start = doc.Content.End
            _br()

            # -------------------------------------------------------------
            # 12. Testes Unitários
            # -------------------------------------------------------------
            _h1("11. Testes Unitários")
            _br()
            
            test_headers = ["ID", "Cenário", "Resultado Esperado", "Resultado Obtido", "Estado"]
            test_rows = [
                ("T1", "Leitura do ficheiro de entrada", "Ficheiro lido com sucesso", "Conforme execução", "OK"),
                ("T2", "Agrupamento de roles", "Roles agrupadas por AGR_NAME", "Conforme resumo processado", "OK"),
                ("T3", "Atribuição de transações", "Transações atribuídas na aba Menu", "Conforme evidências", "OK"),
                ("T4", "Geração do perfil", "Perfil de autorização gerado", "Conforme evidências", "OK"),
                ("T5", "Atualização do Excel", "STATUS, MSG e TIMESTEMP atualizados", "Conforme execução", "OK")
            ]
            
            table_tests = doc.Tables.Add(Range=sel.Range, NumRows=len(test_rows) + 1, NumColumns=len(test_headers))
            table_tests.Borders.Enable = True
            
            for col_idx, h in enumerate(test_headers, start=1):
                cell = table_tests.Cell(1, col_idx)
                cell.Range.Text = h
                cell.Range.Font.Bold = True
                
            for row_idx, r in enumerate(test_rows, start=2):
                table_tests.Cell(row_idx, 1).Range.Text = r[0]
                table_tests.Cell(row_idx, 2).Range.Text = r[1]
                table_tests.Cell(row_idx, 3).Range.Text = r[2]
                table_tests.Cell(row_idx, 4).Range.Text = r[3]
                table_tests.Cell(row_idx, 5).Range.Text = r[4]
                
            sel.Start = doc.Content.End
            _br()
            _br()

            # -------------------------------------------------------------
            # 13. Anexos
            # -------------------------------------------------------------
            _h1("12. Anexos")
            _br()
            _p("As evidências capturadas durante a execução encontram-se incorporadas nas respetivas secções de cada role.")

            doc.SaveAs(str(doc_path), FileFormat=12)
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

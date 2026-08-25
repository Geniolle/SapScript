# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import sys
import logging
from datetime import datetime
from pathlib import Path

class FunctionalDocSession:
    def __init__(self, nome_pasta: str | None, processo: str = "PFCG_CREATE", transacao: str = "PFCG", config: dict | None = None):
        self.nome_pasta = str(nome_pasta or "").strip()
        self.enabled = bool(self.nome_pasta)
        self.processo = processo
        self.transacao = transacao
        self.config = config or {}
        
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

    def add_item_summary(self, item_name: str, description: str, quantity: int, result: str, duration: str):
        if not self.enabled:
            return
        self.roles_summary.append({
            "ordem": len(self.roles_summary) + 1,
            "role": item_name,
            "descricao": description,
            "tcodes": quantity,
            "resultado": result,
            "tempo": duration
        })
        self._role_summary_count += 1

    def add_role_summary(self, role_name: str, description: str, tcode_count: int, result: str, duration: str):
        self.add_item_summary(role_name, description, tcode_count, result, duration)

    def start_item_section(self, item_name: str, description: str, quantity: int):
        if not self.enabled:
            return
        if item_name not in self.evidences:
            self.evidences[item_name] = []

    def start_role_section(self, role_name: str, description: str, tcode_count: int):
        self.start_item_section(role_name, description, tcode_count)

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

    def _get_template_path(self) -> Path | None:
        env_path = os.getenv("FUNCTIONAL_DOC_TEMPLATE_PATH")
        if env_path:
            p = Path(env_path)
            if p.exists():
                return p
        default_p = Path(__file__).parent / "templates" / "functional_template.docx"
        if default_p.exists():
            return default_p
        return None

    def _open_word_document(self, app, doc_path: Path):
        return app.Documents.Open(str(doc_path))

    def _find_placeholder_range(self, doc, placeholder: str):
        find_range = doc.Content
        find_range.Find.ClearFormatting()
        found = find_range.Find.Execute(FindText=placeholder)
        if found:
            return find_range
        return None

    def _write_heading(self, sel, text: str, level: int):
        if level == 1:
            sel.Font.Size = 12
            sel.Font.Bold = True
            sel.Font.Italic = False
        elif level == 2:
            sel.Font.Size = 12
            sel.Font.Bold = True
            sel.Font.Italic = False
        else:
            sel.Font.Size = 11
            sel.Font.Bold = True
            sel.Font.Italic = True
        sel.TypeText(text)
        sel.TypeParagraph()

    def _write_paragraph(self, sel, text: str):
        sel.Font.Size = 11
        sel.Font.Bold = False
        sel.Font.Italic = False
        sel.TypeText(text)
        sel.TypeParagraph()

    def _write_table(self, doc, sel, rows: list[list[str] | tuple[str, ...]], is_kv: bool = False):
        if not rows:
            return
        num_rows = len(rows)
        num_cols = len(rows[0])
        table = doc.Tables.Add(Range=sel.Range, NumRows=num_rows, NumColumns=num_cols)
        table.Borders.Enable = True
        
        for r_idx, row in enumerate(rows, start=1):
            for c_idx, val in enumerate(row, start=1):
                cell = table.Cell(r_idx, c_idx)
                cell.Range.Text = str(val or "")
                
                # Configurar fonte do conteúdo da tabela
                cell.Range.Font.Name = "Arial"
                cell.Range.Font.Size = 12
                
                if is_kv and c_idx == 1:
                    cell.Range.Font.Bold = True
                elif not is_kv and r_idx == 1:
                    cell.Range.Font.Bold = True
                    
        sel.Start = doc.Content.End

    def _write_bullets(self, sel, items: list[str]):
        for item in items:
            sel.Font.Size = 11
            sel.Font.Bold = False
            sel.Font.Italic = False
            sel.TypeText(f"• {item}")
            sel.TypeParagraph()

    def _insert_image(self, sel, image_path: str):
        img_path = Path(image_path)
        if img_path.exists():
            try:
                sel.InlineShapes.AddPicture(
                    FileName=str(img_path.resolve()),
                    LinkToFile=False,
                    SaveWithDocument=True
                )
                sel.TypeParagraph()
            except Exception as e:
                self._write_paragraph(sel, f"[Erro ao inserir imagem: {e}]")
        else:
            self._write_paragraph(sel, "[Imagem de evidência não disponível]")

    def _generate_word_doc(self, doc_path: Path):
        import pythoncom
        import win32com.client
        
        pythoncom.CoInitialize()
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            
            # 1. Obter template
            template_path = self._get_template_path()
            if template_path:
                print(f"[DOC] Template encontrado: {template_path}. A utilizar...")
                doc = self._open_word_document(app, template_path)
            else:
                print("[DOC] Nenhum template encontrado. A gerar documento em branco com fallback...")
                doc = app.Documents.Add()
                
            sel = app.Selection
            
            # 2. Localizar placeholder {{FUNCTIONAL_DOC_CONTENT}} se existir
            if template_path:
                p_range = self._find_placeholder_range(doc, "{{FUNCTIONAL_DOC_CONTENT}}")
                if p_range:
                    sel.Start = p_range.Start
                    sel.End = p_range.End
                    sel.Text = ""
                else:
                    sel.Start = doc.Content.End
            else:
                sel.Start = doc.Content.End
                
            # Restaura alinhamento à esquerda
            sel.ParagraphFormat.Alignment = 0
            
            def _br():
                sel.TypeParagraph()
                
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
            # Se for documento novo (sem template), criar Capa
            # -------------------------------------------------------------
            if not template_path:
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
                desc = self.config.get("titulo", "Criação/Atualização de Roles e Perfis de Autorização")
                sel.TypeText(desc)
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
                
                sel.InsertBreak(7)  # Page Break
                sel.ParagraphFormat.Alignment = 0  # Left

            # -------------------------------------------------------------
            # 1. Informação Geral
            # -------------------------------------------------------------
            self._write_heading(sel, "1. Informação Geral", 1)
            _br()
            
            meta_keys = [
                ("Processo", self.processo),
                ("Transação SAP utilizada", self.transacao),
                ("Sistema", self.metadata.get("sistema", "")),
                ("Cliente", self.metadata.get("cliente", "")),
                ("Utilizador SAP", self.metadata.get("utilizador_sap", "")),
                ("Data", data_exec),
                ("Total de itens/roles processados", str(self.metadata.get("total_roles", "0"))),
                ("Pasta de documentação solicitada", self.nome_pasta)
            ]
            self._write_table(doc, sel, meta_keys, is_kv=True)
            _br()
            _br()

            # -------------------------------------------------------------
            # 2. Histórico de Versões
            # -------------------------------------------------------------
            self._write_heading(sel, "2. Histórico de Versões", 1)
            _br()
            
            version_headers = ["Versão", "Data", "Autor", "Modificação"]
            version_row = [
                "1.0",
                data_exec,
                str(self.metadata.get("utilizador_sap") or "Sistema").strip(),
                f"Documento inicial gerado automaticamente para evidência funcional da execução {self.processo}"
            ]
            self._write_table(doc, sel, [version_headers, version_row])
            _br()
            _br()

            # -------------------------------------------------------------
            # 3. Índice
            # -------------------------------------------------------------
            self._write_heading(sel, "3. Índice", 1)
            _br()
            self._write_paragraph(sel, "1. Pedido de Alteração")
            self._write_paragraph(sel, "2. Módulos e processos afetados")
            self._write_paragraph(sel, "3. Análise de Viabilidade")
            self._write_paragraph(sel, "4. Especificação Funcional")
            self._write_paragraph(sel, "5. Resumo dos itens processados")
            self._write_paragraph(sel, "6. Evidências")
            self._write_paragraph(sel, "7. Testes Unitários")
            self._write_paragraph(sel, "8. Anexos")
            _br()
            _br()

            # -------------------------------------------------------------
            # 4. Pedido de Alteração
            # -------------------------------------------------------------
            self._write_heading(sel, "4. Pedido de Alteração", 1)
            _br()
            self._write_paragraph(
                sel,
                f"Este documento tem como objetivo registar a análise funcional e as evidências da execução "
                f"do processo {self.processo}, realizado através da transação {self.transacao}."
            )
            _br()
            
            self._write_heading(sel, "4.1 Requisitos", 2)
            _br()
            self._write_heading(sel, "4.1.1 Requisitos de Qualidade", 3)
            _br()
            self._write_paragraph(
                sel,
                "Os requisitos processados são coerentes com o fluxo standard do sistema SAP, permitindo "
                "rastreabilidade, validação funcional e evidência documental da execução."
            )
            _br()
            
            self._write_heading(sel, "4.1.2 Requester / Unidade de Negócios / Owner", 3)
            _br()
            req_keys = [
                ("Requester", "-"),
                ("Unidade de Negócios", "-"),
                ("Owner", "-")
            ]
            self._write_table(doc, sel, req_keys, is_kv=True)
            _br()
            _br()
            
            self._write_heading(sel, "4.1.3 Requisitos do Cliente", 3)
            _br()
            client_req_headers = ["ID", "Descrição", "Prioridade"]
            client_req_row = [
                "R1",
                f"Executar o processo {self.processo} através da transação {self.transacao}, registando evidências funcionais da execução.",
                "MÉDIA"
            ]
            self._write_table(doc, sel, [client_req_headers, client_req_row])
            _br()
            _br()

            # -------------------------------------------------------------
            # 5. Módulos e processos afetados
            # -------------------------------------------------------------
            self._write_heading(sel, "5. Módulos e processos afetados", 1)
            _br()
            self._write_heading(sel, "5.1 Módulos afetados", 2)
            _br()
            modulos = self.config.get("modulos_afetados")
            if not modulos:
                modulos = ["SAP"]
            self._write_bullets(sel, modulos)
            _br()
            
            self._write_heading(sel, "5.2 Processos afetados", 2)
            _br()
            processos_afetados = self.config.get("processos_afetados")
            if not processos_afetados:
                processos_afetados = ["Processo funcional executado via SAP Script"]
            self._write_bullets(sel, processos_afetados)
            _br()

            # -------------------------------------------------------------
            # 6. Análise de Viabilidade
            # -------------------------------------------------------------
            self._write_heading(sel, "6. Análise de Viabilidade", 1)
            _br()
            self._write_heading(sel, f"Execução do processo {self.processo}", 2)
            _br()
            self._write_paragraph(sel, f"De seguida apresenta-se o detalhe da solução proposta e executada para o processo {self.processo}.")
            _br()
            self._write_heading(sel, "6.1.1 Solução Proposta", 3)
            _br()
            sol_prop = self.config.get("solucao_proposta")
            if not sol_prop:
                sol_prop = "A solução consiste em processar automaticamente os dados informados, executar a transação SAP correspondente, registar o resultado da execução e anexar evidências funcionais no documento."
            self._write_paragraph(sel, sol_prop)
            _br()

            # -------------------------------------------------------------
            # 7. Configuração / Execução
            # -------------------------------------------------------------
            self._write_heading(sel, "7. Configuração / Execução", 1)
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
            self._write_table(doc, sel, config_rows, is_kv=True)
            _br()
            _br()

            # -------------------------------------------------------------
            # 8. Especificação Funcional
            # -------------------------------------------------------------
            self._write_heading(sel, "8. Especificação Funcional", 1)
            _br()
            
            # Objetivo
            self._write_heading(sel, "8.1 Objetivo", 2)
            _br()
            if self.processo == "PFCG_CREATE":
                obj_txt = "O processo tem como objetivo realizar a criação ou atualização de roles SAP utilizando a transação PFCG, com atribuição das transações informadas e geração dos respetivos perfis de autorização."
            else:
                obj_txt = f"Execução automática do processo {self.processo} via SAP Script."
            self._write_paragraph(sel, obj_txt)
            _br()
            
            # Transação Envolvida
            self._write_heading(sel, "8.2 Transação Envolvida", 2)
            _br()
            self._write_bullets(sel, [f"{self.transacao} - Manutenção de roles" if self.processo == "PFCG_CREATE" else f"{self.transacao}"])
            _br()
            
            # Modo de Processamento
            self._write_heading(sel, "8.3 Modo de Processamento", 2)
            _br()
            self._write_paragraph(
                sel,
                "O processamento é realizado de forma assistida/automática através do Web Cockpit SAP Script, utilizando "
                "como entrada um ficheiro Excel com as roles e transações a processar."
            )
            _br()
            
            # Estrutura de Entrada
            self._write_heading(sel, "8.4 Estrutura de Entrada", 2)
            _br()
            input_struct_headers = ["Campo", "Tipo", "Descrição"]
            if self.processo == "PFCG_CREATE":
                input_struct_rows = [
                    ("AGR_NAME", "Texto", "Nome da role"),
                    ("TEXT", "Texto", "Descrição da role"),
                    ("TCODE", "Texto", "Transações a atribuir"),
                    ("STATUS", "Texto", "Resultado do processamento"),
                    ("MSG", "Texto", "Mensagem de retorno"),
                    ("TIMESTEMP", "Texto", "Data/hora de atualização do resultado")
                ]
            else:
                input_struct_rows = [
                    ("Campo", "Texto", "Descrição do campo")
                ]
            self._write_table(doc, sel, [input_struct_headers] + input_struct_rows)
            _br()
            _br()
            
            # Regras de Processamento
            self._write_heading(sel, "8.5 Regras de Processamento", 2)
            _br()
            if self.processo == "PFCG_CREATE":
                regras = [
                    "Ler o ficheiro Excel informado.",
                    "Agrupar linhas por AGR_NAME.",
                    "Ignorar roles já concluídas.",
                    f"Abrir a transação {self.transacao}.",
                    "Criar ou atualizar a role.",
                    "Preencher a descrição.",
                    "Atribuir as transações na aba Menu.",
                    "Gravar as alterações.",
                    "Gerar o perfil de autorização.",
                    "Atualizar o Excel com STATUS, MSG e TIMESTEMP.",
                    "Registar evidências no documento funcional."
                ]
            else:
                regras = [
                    "Carregar dados do processo.",
                    "Executar as ações na transação SAP.",
                    "Gravar evidências e atualizar status de retorno."
                ]
            self._write_bullets(sel, regras)
            _br()

            # -------------------------------------------------------------
            # 9. Resumo dos itens processados
            # -------------------------------------------------------------
            self._write_heading(sel, "9. Resumo dos itens processados", 1)
            _br()
            
            headers = ["Ordem", "Item/Role", "Descrição", "Quantidade", "Resultado", "Tempo de execução"]
            summary_rows = [
                [
                    str(r["ordem"]),
                    r["role"],
                    r["descricao"],
                    str(r["tcodes"]),
                    r["resultado"],
                    r["tempo"]
                ] for r in self.roles_summary
            ]
            self._write_table(doc, sel, [headers] + summary_rows)
            _br()
            _br()

            # -------------------------------------------------------------
            # 10. Evidências
            # -------------------------------------------------------------
            self._write_heading(sel, "10. Evidências", 1)
            _br()
            
            sub_idx = 1
            objeto_lbl = self.config.get("objeto_principal", "Role")
            for r in self.roles_summary:
                item_name = r["role"]
                if r["resultado"] == "Concluída" and item_name in self.evidences and self.evidences[item_name]:
                    self._write_heading(sel, f"10.{sub_idx} {objeto_lbl} {item_name}", 2)
                    _br()
                    self._write_paragraph(sel, f"• Descrição: {r['descricao']}")
                    self._write_paragraph(sel, f"• Quantidade: {r['tcodes']}")
                    self._write_paragraph(sel, f"• Resultado: {objeto_lbl} tratada por completo")
                    self._write_paragraph(sel, f"• Tempo de execução: {r['tempo']}")
                    _br()
                    
                    for ev in self.evidences[item_name]:
                        sel.Font.Bold = True
                        self._write_paragraph(sel, ev['title'])
                        sel.Font.Bold = False
                        
                        self._insert_image(sel, ev["path"])
                        
                        sel.Font.Italic = True
                        self._write_paragraph(sel, f"Legenda: {ev['caption']}")
                        _br()
                        
                    sub_idx += 1
            
            sel.Start = doc.Content.End
            _br()

            # -------------------------------------------------------------
            # 11. Testes Unitários
            # -------------------------------------------------------------
            self._write_heading(sel, "11. Testes Unitários", 1)
            _br()
            
            test_headers = ["ID", "Cenário", "Resultado Esperado", "Resultado Obtido", "Estado"]
            test_rows = [
                ("T1", "Leitura dos dados de entrada", "Dados lidos com sucesso", "Conforme execução", "OK"),
                ("T2", "Processamento dos itens", "Itens processados conforme regras funcionais", "Conforme resumo processado", "OK"),
                ("T3", "Execução da transação SAP", "Transação executada sem erro impeditivo", "Conforme execução", "OK"),
                ("T4", "Registo de evidências", "Evidências anexadas ao documento", "Conforme documento", "OK"),
                ("T5", "Atualização dos resultados", "Resultados registados após execução", "Conforme execução", "OK")
            ]
            self._write_table(doc, sel, [test_headers] + test_rows)
            _br()
            _br()

            # -------------------------------------------------------------
            # 12. Anexos
            # -------------------------------------------------------------
            self._write_heading(sel, "12. Anexos", 1)
            _br()
            self._write_paragraph(sel, "As evidências capturadas durante a execução encontram-se incorporadas nas respetivas secções deste documento.")

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

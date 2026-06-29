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
            
            # Título
            sel.ParagraphFormat.Alignment = 1  # Center
            sel.Font.Size = 16
            sel.Font.Bold = True
            sel.Font.Name = "Arial"
            sel.TypeText("Documento Funcional de Configuração SAP")
            sel.TypeParagraph()
            sel.TypeParagraph()
            
            # Restaura alinhamento à esquerda
            sel.ParagraphFormat.Alignment = 0  # Left
            sel.Font.Size = 11
            sel.Font.Bold = False
            
            # Secção 2: Dados da execução
            sel.Font.Size = 13
            sel.Font.Bold = True
            sel.TypeText("2. Dados da execução")
            sel.TypeParagraph()
            sel.Font.Size = 11
            sel.Font.Bold = False
            
            meta_keys = [
                ("Processo executado", self.processo),
                ("Transação SAP utilizada", self.transacao),
                ("Ambiente", self.metadata.get("ambiente", "")),
                ("Sistema", self.metadata.get("sistema", "")),
                ("Cliente", self.metadata.get("cliente", "")),
                ("Utilizador SAP", self.metadata.get("utilizador_sap", "")),
                ("Data/hora de início", self.metadata.get("data_inicio", "")),
                ("Data/hora de fim", self.metadata.get("data_fim", "")),
                ("Total de roles processadas", str(self.metadata.get("total_roles", "0"))),
                ("Ficheiro Excel utilizado", self.metadata.get("excel_utilizado", "")),
                ("Pasta de documentação solicitada", self.nome_pasta)
            ]
            
            table_meta = doc.Tables.Add(Range=sel.Range, NumRows=len(meta_keys), NumColumns=2)
            table_meta.Borders.Enable = True
            for idx, (k, v) in enumerate(meta_keys, start=1):
                table_meta.Cell(idx, 1).Range.Text = k
                table_meta.Cell(idx, 1).Range.Font.Bold = True
                table_meta.Cell(idx, 2).Range.Text = str(v or "")
                
            # Mover seleção para baixo da tabela
            sel.Start = doc.Content.End
            sel.TypeParagraph()
            sel.TypeParagraph()
            
            # Secção 3: Objetivo da execução
            sel.Font.Size = 13
            sel.Font.Bold = True
            sel.TypeText("3. Objetivo da execução")
            sel.TypeParagraph()
            sel.Font.Size = 11
            sel.Font.Bold = False
            sel.TypeText(
                "Foi utilizada a transação PFCG para criar/atualizar roles, atribuir as respetivas "
                "transações, gravar as alterações no SAP e gerar os perfis de autorização correspondentes."
            )
            sel.TypeParagraph()
            sel.TypeParagraph()
            
            # Secção 4: Resumo das roles processadas
            sel.Font.Size = 13
            sel.Font.Bold = True
            sel.TypeText("4. Resumo das roles processadas")
            sel.TypeParagraph()
            sel.Font.Size = 11
            sel.Font.Bold = False
            
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
            sel.TypeParagraph()
            sel.TypeParagraph()
            
            # Secção 5: Evidências por role
            sel.Font.Size = 13
            sel.Font.Bold = True
            sel.TypeText("5. Evidências por role")
            sel.TypeParagraph()
            sel.Font.Size = 11
            sel.Font.Bold = False
            
            sub_idx = 1
            for r in self.roles_summary:
                role_name = r["role"]
                # Apenas secções detalhadas para roles concluídas com sucesso e que tenham evidências
                if r["resultado"] == "Concluída" and role_name in self.evidences and self.evidences[role_name]:
                    sel.Font.Size = 12
                    sel.Font.Bold = True
                    sel.TypeText(f"5.{sub_idx} Role {role_name}")
                    sel.TypeParagraph()
                    sel.Font.Size = 11
                    sel.Font.Bold = False
                    
                    sel.TypeText(f"• Descrição: {r['descricao']}")
                    sel.TypeParagraph()
                    sel.TypeText(f"• Quantidade de TCODEs atribuídas: {r['tcodes']}")
                    sel.TypeParagraph()
                    sel.TypeText(f"• Resultado: Role tratada por completo")
                    sel.TypeParagraph()
                    sel.TypeText(f"• Tempo de execução: {r['tempo']}")
                    sel.TypeParagraph()
                    sel.TypeParagraph()
                    
                    for ev in self.evidences[role_name]:
                        sel.Font.Bold = True
                        sel.TypeText(f"{ev['title']}")
                        sel.TypeParagraph()
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
                                sel.TypeText(f"[Erro ao inserir imagem: {img_exc}]")
                                sel.TypeParagraph()
                        else:
                            sel.TypeText("[Imagem de evidência não disponível]")
                            sel.TypeParagraph()
                            
                        sel.Font.Italic = True
                        sel.TypeText(f"Legenda: {ev['caption']}")
                        sel.Font.Italic = False
                        sel.TypeParagraph()
                        sel.TypeParagraph()
                        
                    sub_idx += 1
            
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

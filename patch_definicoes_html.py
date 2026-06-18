#!/usr/bin/env python3
"""
Patch: Adiciona a view Definições com CRUD de regras de contexto do Agente SAP.
"""
import re

HTML_FILE = r"c:\workspace\sap-script\sap_script_web_cockpit_v2\web_api\templates\index.html"

with open(HTML_FILE, "r", encoding="utf-8") as f:
    content = f.read()

# ── 1. Nav item: ligar "Definições" à view correta ─────────────────────────
OLD_NAV = '''        <a class="nav-item" onclick="switchView('visao-geral')">Definições</a>'''
NEW_NAV = '''        <a class="nav-item" id="nav-item-definicoes" onclick="switchView('definicoes')">⚙️ Definições</a>'''
assert OLD_NAV in content, "NAV not found"
content = content.replace(OLD_NAV, NEW_NAV, 1)

# ── 2. Inserir view-definicoes antes do comentário END SAP AGENT VIEW ──────
VIEW_ANCHOR = "        <!-- ====== END SAP AGENT VIEW ====== -->"

VIEW_DEFINICOES = '''        <!-- ====== DEFINIÇÕES VIEW ====== -->
        <div id="view-definicoes" style="display: none;">
          <div style="width: 100%;">

            <!-- Header -->
            <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; flex-wrap: wrap; gap: 12px;">
              <div style="display: flex; align-items: center; gap: 14px;">
                <div style="width: 42px; height: 42px; border-radius: 12px; background: linear-gradient(135deg, #6366f1, #4f46e5); display: flex; align-items: center; justify-content: center; font-size: 20px; box-shadow: 0 4px 12px rgba(99,102,241,0.35);">⚙️</div>
                <div>
                  <h2 style="margin: 0; font-size: 1.3rem; font-weight: 800; color: var(--text-primary);">Definições</h2>
                  <p style="margin: 0; font-size: 0.75rem; color: var(--text-secondary);">Parâmetros de contexto que enriquecem a análise do Agente SAP.</p>
                </div>
              </div>
            </div>

            <!-- ══ Agente SAP — Regras de Contexto ══ -->
            <div class="card" style="padding: 0; overflow: hidden; border-radius: 16px; border: 1px solid #e2e8f0; box-shadow: 0 4px 20px rgba(0,0,0,0.06);">
              <!-- Card Header -->
              <div style="padding: 16px 22px; background: linear-gradient(135deg, rgba(99,102,241,0.07), rgba(16,185,129,0.04)); border-bottom: 1px solid #f1f5f9; display: flex; align-items: center; justify-content: space-between;">
                <div style="display: flex; align-items: center; gap: 10px;">
                  <span style="font-size: 1.1rem;">🤖</span>
                  <div>
                    <div style="font-size: 0.85rem; font-weight: 800; color: #334155;">Regras de Contexto — Agente SAP</div>
                    <div style="font-size: 0.72rem; color: #64748b; margin-top: 1px;">Defina informações complementares que o agente usa automaticamente ao analisar tickets.</div>
                  </div>
                </div>
                <button type="button" id="def-btn-nova-regra" onclick="defOpenModal()" style="display: inline-flex; align-items: center; gap: 7px; padding: 8px 16px; background: linear-gradient(135deg, #6366f1, #4f46e5); color: white; border: none; border-radius: 10px; font-size: 0.8rem; font-weight: 700; cursor: pointer; box-shadow: 0 4px 12px rgba(99,102,241,0.35); transition: all 0.2s;" onmouseover="this.style.transform='translateY(-1px)';this.style.boxShadow='0 6px 16px rgba(99,102,241,0.4)'" onmouseout="this.style.transform='';this.style.boxShadow='0 4px 12px rgba(99,102,241,0.35)'">
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
                  Nova Regra
                </button>
              </div>

              <!-- Info Banner -->
              <div style="padding: 12px 22px; background: rgba(99,102,241,0.04); border-bottom: 1px solid #f1f5f9; display: flex; align-items: flex-start; gap: 10px;">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0; margin-top:1px;"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>
                <div style="font-size: 0.75rem; color: #475569; line-height: 1.5;">
                  Quando o Agente SAP analisa um ticket, verifica se algum campo do ticket corresponde a uma regra aqui configurada.
                  Se houver correspondência, a <strong>Transação SAP</strong> e as <strong>Notas</strong> são mostradas automaticamente nos <em>Sinais Identificados</em>.
                  <br><span style="font-size:0.7rem; color:#94a3b8; margin-top:3px; display:block;">Exemplo: se o ticket tem <strong>IT SALSA - Categoria SAP = "FI Código IVA"</strong> → o agente sugere a transação <strong>FTXP</strong>.</span>
                </div>
              </div>

              <!-- Rules Table -->
              <div style="overflow-x: auto;">
                <table id="def-rules-table" style="width: 100%; border-collapse: collapse; font-size: 0.78rem; color: #334155;">
                  <thead>
                    <tr style="background: #f8fafc; border-bottom: 2px solid #e2e8f0;">
                      <th style="text-align: left; padding: 10px 16px; font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; white-space: nowrap;">Campo</th>
                      <th style="text-align: left; padding: 10px 16px; font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em;">Valor</th>
                      <th style="text-align: left; padding: 10px 16px; font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; white-space: nowrap;">Transação SAP</th>
                      <th style="text-align: left; padding: 10px 16px; font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em;">Tags</th>
                      <th style="text-align: left; padding: 10px 16px; font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; max-width: 260px;">Notas</th>
                      <th style="text-align: center; padding: 10px 16px; font-size: 0.68rem; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; white-space: nowrap;">Ações</th>
                    </tr>
                  </thead>
                  <tbody id="def-rules-tbody">
                    <tr id="def-rules-loading-row">
                      <td colspan="6" style="text-align: center; padding: 32px; color: #94a3b8; font-size: 0.8rem;">A carregar regras...</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

          </div>
        </div>
        <!-- ====== END DEFINIÇÕES VIEW ====== -->

'''

assert VIEW_ANCHOR in content, "VIEW_ANCHOR not found"
content = content.replace(VIEW_ANCHOR, VIEW_DEFINICOES + VIEW_ANCHOR, 1)

# ── 3. Modal de criar/editar regra (antes de </body>) ─────────────────────
MODAL_ANCHOR = "</body>"

MODAL_HTML = '''  <!-- ══ Modal: Criar/Editar Regra de Contexto ══ -->
  <div id="def-rule-modal" style="display:none; position:fixed; inset:0; z-index:10000; display:none; align-items:center; justify-content:center; background:rgba(0,0,0,0.55); backdrop-filter:blur(4px);">
    <div style="background:#ffffff; border-radius:18px; box-shadow:0 25px 60px rgba(0,0,0,0.25); width:540px; max-width:95vw; max-height:90vh; overflow-y:auto; animation: modalSlideIn 0.22s ease;">
      <!-- Modal Header -->
      <div style="padding:20px 24px 14px; border-bottom:1px solid #f1f5f9; display:flex; align-items:center; justify-content:space-between;">
        <div style="display:flex; align-items:center; gap:10px;">
          <div style="width:34px; height:34px; border-radius:10px; background:linear-gradient(135deg,#6366f1,#4f46e5); display:flex; align-items:center; justify-content:center; font-size:16px; box-shadow:0 4px 10px rgba(99,102,241,0.3);">⚙️</div>
          <h3 id="def-modal-title" style="margin:0; font-size:0.95rem; font-weight:800; color:#1e293b;">Nova Regra de Contexto</h3>
        </div>
        <button onclick="defCloseModal()" style="background:none; border:none; cursor:pointer; color:#94a3b8; font-size:1.2rem; padding:4px 8px; border-radius:6px; transition:all 0.15s;" onmouseover="this.style.background='#f1f5f9';this.style.color='#475569'" onmouseout="this.style.background='none';this.style.color='#94a3b8'">✕</button>
      </div>

      <!-- Modal Body -->
      <div style="padding:22px 24px;">
        <input type="hidden" id="def-rule-id">

        <!-- Campo -->
        <div style="margin-bottom:16px;">
          <label for="def-rule-campo" style="display:flex; align-items:center; gap:6px; font-size:11px; font-weight:700; color:#64748b; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">
            <span style="display:inline-block; width:6px; height:6px; border-radius:50%; background:#6366f1;"></span>
            Campo JIRA <span style="color:#ef4444">*</span>
          </label>
          <select id="def-rule-campo" style="width:100%; padding:10px 36px 10px 14px; font-size:0.85rem; font-weight:500; color:#334155; background-color:#f8fafc; background-image:url('data:image/svg+xml;charset=UTF-8,%3csvg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 24 24\' fill=\'none\' stroke=\'%2364748b\' stroke-width=\'2\' stroke-linecap=\'round\' stroke-linejoin=\'round\'%3e%3cpolyline points=\'6 9 12 15 18 9\'%3e%3c/polyline%3e%3c/svg%3e'); background-repeat:no-repeat; background-position:right 12px center; background-size:14px; border:1px solid #e2e8f0; border-radius:10px; appearance:none; -webkit-appearance:none; outline:none; cursor:pointer; transition:all 0.2s; box-sizing:border-box;">
            <option value="IT SALSA - Categoria SAP">IT SALSA - Categoria SAP</option>
            <option value="Tipo de Ticket">Tipo de Ticket</option>
            <option value="Stream">Stream</option>
          </select>
        </div>

        <!-- Valor -->
        <div style="margin-bottom:16px;">
          <label for="def-rule-valor" style="display:flex; align-items:center; gap:6px; font-size:11px; font-weight:700; color:#64748b; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">
            <span style="display:inline-block; width:6px; height:6px; border-radius:50%; background:#6366f1;"></span>
            Valor do Campo <span style="color:#ef4444">*</span>
          </label>
          <input type="text" id="def-rule-valor" placeholder="Ex: FI Código IVA" style="width:100%; padding:10px 14px; font-size:0.85rem; font-weight:500; color:#334155; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; outline:none; transition:all 0.2s; box-sizing:border-box;" onfocus="this.style.borderColor='#6366f1';this.style.boxShadow='0 0 0 4px rgba(99,102,241,0.12)'" onblur="this.style.borderColor='#e2e8f0';this.style.boxShadow='none'">
        </div>

        <!-- Transação SAP -->
        <div style="margin-bottom:16px;">
          <label for="def-rule-transacao" style="display:flex; align-items:center; gap:6px; font-size:11px; font-weight:700; color:#64748b; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">
            <span style="display:inline-block; width:6px; height:6px; border-radius:50%; background:#10b981;"></span>
            Transação SAP
          </label>
          <input type="text" id="def-rule-transacao" placeholder="Ex: FTXP, SE38, SM30..." style="width:100%; padding:10px 14px; font-size:0.85rem; font-weight:600; color:#334155; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; outline:none; transition:all 0.2s; box-sizing:border-box; font-family:monospace;" onfocus="this.style.borderColor='#10b981';this.style.boxShadow='0 0 0 4px rgba(16,185,129,0.12)'" onblur="this.style.borderColor='#e2e8f0';this.style.boxShadow='none'">
          <div style="font-size:0.68rem; color:#94a3b8; margin-top:4px;">Código da transação SAP relevante para este cenário (ex: FTXP para códigos IVA)</div>
        </div>

        <!-- Tags -->
        <div style="margin-bottom:16px;">
          <label for="def-rule-tags" style="display:flex; align-items:center; gap:6px; font-size:11px; font-weight:700; color:#64748b; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">
            <span style="display:inline-block; width:6px; height:6px; border-radius:50%; background:#f59e0b;"></span>
            Tags
          </label>
          <input type="text" id="def-rule-tags" placeholder="Ex: IVA, Impostos, FI (separadas por vírgula)" style="width:100%; padding:10px 14px; font-size:0.85rem; font-weight:500; color:#334155; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; outline:none; transition:all 0.2s; box-sizing:border-box;" onfocus="this.style.borderColor='#f59e0b';this.style.boxShadow='0 0 0 4px rgba(245,158,11,0.12)'" onblur="this.style.borderColor='#e2e8f0';this.style.boxShadow='none'">
        </div>

        <!-- Notas -->
        <div style="margin-bottom:4px;">
          <label for="def-rule-notas" style="display:flex; align-items:center; gap:6px; font-size:11px; font-weight:700; color:#64748b; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">
            <span style="display:inline-block; width:6px; height:6px; border-radius:50%; background:#3b82f6;"></span>
            Notas / Contexto Adicional
          </label>
          <textarea id="def-rule-notas" rows="4" placeholder="Descreva o cenário, passos de diagnóstico, verificações comuns..." style="width:100%; padding:10px 14px; font-size:0.82rem; font-weight:500; color:#334155; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; outline:none; transition:all 0.2s; box-sizing:border-box; resize:vertical; font-family:inherit; line-height:1.5;" onfocus="this.style.borderColor='#3b82f6';this.style.boxShadow='0 0 0 4px rgba(59,130,246,0.12)'" onblur="this.style.borderColor='#e2e8f0';this.style.boxShadow='none'"></textarea>
        </div>
      </div>

      <!-- Modal Footer -->
      <div style="padding:14px 24px 20px; border-top:1px solid #f1f5f9; display:flex; align-items:center; justify-content:flex-end; gap:10px;">
        <button type="button" onclick="defCloseModal()" style="padding:9px 18px; font-size:0.82rem; font-weight:600; color:#64748b; background:#f1f5f9; border:1px solid #e2e8f0; border-radius:10px; cursor:pointer; transition:all 0.15s;" onmouseover="this.style.background='#e2e8f0'" onmouseout="this.style.background='#f1f5f9'">Cancelar</button>
        <button type="button" id="def-modal-save-btn" onclick="defSaveRule()" style="padding:9px 22px; font-size:0.82rem; font-weight:700; color:white; background:linear-gradient(135deg,#6366f1,#4f46e5); border:none; border-radius:10px; cursor:pointer; box-shadow:0 4px 12px rgba(99,102,241,0.35); transition:all 0.2s;" onmouseover="this.style.transform='translateY(-1px)'" onmouseout="this.style.transform=''">💾 Guardar Regra</button>
      </div>
    </div>
  </div>

'''

assert MODAL_ANCHOR in content, "MODAL_ANCHOR not found"
content = content.replace(MODAL_ANCHOR, MODAL_HTML + MODAL_ANCHOR, 1)

# ── 4. Adicionar CSS de animação do modal (antes de </style> no head) ─────
STYLE_ANCHOR = "    @keyframes pulse-opacity {"
MODAL_CSS = """    @keyframes modalSlideIn {
      from { opacity: 0; transform: scale(0.95) translateY(-10px); }
      to   { opacity: 1; transform: scale(1) translateY(0); }
    }
    #def-rule-modal { display: flex !important; }
    #def-rule-modal[style*="display:none"], #def-rule-modal[style*="display: none"] { display: none !important; }
    .def-rule-row:hover { background: #f8fafc; }
    .def-rule-row td { padding: 11px 16px; border-bottom: 1px solid #f1f5f9; vertical-align: top; }

"""
assert STYLE_ANCHOR in content, "STYLE_ANCHOR not found"
content = content.replace(STYLE_ANCHOR, MODAL_CSS + STYLE_ANCHOR, 1)

with open(HTML_FILE, "w", encoding="utf-8") as f:
    f.write(content)

print("SUCCESS: HTML structure patches applied")
print(f"File size: {len(content)} bytes")

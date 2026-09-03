/*
 * cockpit.core.js - estado global, wiring do DOM, fila de jobs e dashboard.
 * Carrega ANTES de cockpit.agent.js. Ver Fase 2b em docs/REFACTOR_AGENTE_SALSA_IT.md.
 */
    // ── SAP AGENT filter sidebar state (declared at top for global scope) ──
    let saAllTickets = [];
    let saFilteredTickets = [];
    let saActiveTicketKey = null;

    function formatTimeOnly(date) {
        return [
            String(date.getHours()).padStart(2, '0'),
            String(date.getMinutes()).padStart(2, '0'),
            String(date.getSeconds()).padStart(2, '0')
        ].join(':');
    }
    function getTimestampForLine(job, lineIndex, totalLines) {
        const start = new Date(job.created_at);
        const end = (job.state === 'running' || job.state === 'pending') ? new Date() : new Date(job.updated_at);
        const diffMs = Math.max(0, end - start);
        if (totalLines <= 1) {
            return formatTimeOnly(start);
        }
        const stepMs = diffMs / (totalLines - 1 || 1);
        const lineTime = new Date(start.getTime() + (lineIndex * stepMs));
        return formatTimeOnly(lineTime);
    }
    function getRawDateForLine(job, lineIndex, totalLines) {
        const start = new Date(job.created_at);
        const end = (job.state === 'running' || job.state === 'pending') ? new Date() : new Date(job.updated_at);
        const diffMs = Math.max(0, end - start);
        if (totalLines <= 1) {
            return start;
        }
        const stepMs = diffMs / (totalLines - 1 || 1);
        return new Date(start.getTime() + (lineIndex * stepMs));
    }
    function formatDuration(startStr, endStr) {
        const start = new Date(startStr);
        const end = endStr ? new Date(endStr) : new Date();
        const diffMs = Math.max(0, end - start);
        const diffSecs = Math.floor(diffMs / 1000);
        const hours = Math.floor(diffSecs / 3600);
        const minutes = Math.floor((diffSecs % 3600) / 60);
        const seconds = diffSecs % 60;
        return [
            String(hours).padStart(2, '0'),
            String(minutes).padStart(2, '0'),
            String(seconds).padStart(2, '0')
        ].join(':');
    }
    function formatJiraDate(dateStr) {
        if (!dateStr) return '<span style="color: #9ca3af;">-</span>';
        try {
            const date = new Date(dateStr);
            if (isNaN(date.getTime())) return escapeHtml(dateStr);
            const pad = n => String(n).padStart(2, '0');
            return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}`;
        } catch (e) {
            return escapeHtml(dateStr);
        }
    }
    function parseStageLabel(text, index) {
        const textLower = text.toLowerCase();
        
        if (textLower === 'inicialização e logon sap') {
            return { title: 'Inicialização', sub: 'Login SAP' };
        }
        if (textLower === 'abertura e verificação de tela') {
            return { title: 'Verificação de Tela', sub: 'Abertura e validação' };
        }
        if (textLower === 'processamento principal / ações gui' || textLower.includes('processamento principal')) {
            return { title: 'Processamento GUI', sub: 'Execução das ações' };
        }
        if (textLower === 'conclusão / geração de logs' || textLower.includes('conclusão')) {
            return { title: 'Conclusão / Logs', sub: 'Geração de logs' };
        }
        
        if (textLower === 'preparação e dados') {
            return { title: 'Preparação', sub: 'Dados e regras' };
        }
        if (textLower === 'atribuição de tcods') {
            return { title: 'Atribuição', sub: 'Transações e permissões' };
        }
        if (textLower === 'gerar perfil') {
            return { title: 'Perfil SAP', sub: 'Geração do perfil' };
        }
        if (textLower === 'ordem de transporte') {
            return { title: 'Transporte', sub: 'Liberar Request' };
        }
        if (textLower === 'leitura do excel') {
            return { title: 'Leitura Excel', sub: 'Carregar dados' };
        }
        if (textLower === 'pesquisa de utilizadores') {
            return { title: 'Pesquisa', sub: 'Validar utilizadores' };
        }
        if (textLower === 'atribuição no sap cua') {
            return { title: 'Atribuição CUA', sub: 'Processar perfis' };
        }
        if (textLower === 'gravação de resultados' || textLower === 'atualização do excel') {
            return { title: 'Resultados', sub: 'Salvar ficheiro' };
        }
        if (textLower === 'acesso ao sap') {
            return { title: 'Conexão SAP', sub: 'Abrir sessão' };
        }
        if (textLower === 'eliminação das funções') {
            return { title: 'Eliminar funções', sub: 'Remover no SAP' };
        }
        if (textLower === 'validação da role') {
            return { title: 'Validar Role', sub: 'Verificar status' };
        }
        if (textLower === 'inserção massval') {
            return { title: 'Ações MASSVAL', sub: 'Inserir dados' };
        }
        
        if (text.includes('/')) {
            const parts = text.split('/');
            return { title: parts[0].trim(), sub: parts[1].trim() };
        }
        if (text.includes('|')) {
            const parts = text.split('|');
            return { title: parts[0].trim(), sub: parts[1].trim() };
        }
        if (text.includes(' e ')) {
            const parts = text.split(' e ');
            return { title: parts[0].trim(), sub: parts[1].trim() };
        }
        
        return { title: text, sub: `Etapa ${index + 1}` };
    }
    const STORAGE_KEY_EXCEL_PATH = 'sap_script_web_last_excel_path';
    const JIRA_POLL_SECONDS = (window.__COCKPIT__ && window.__COCKPIT__.pollSeconds) || 60;
    const FI_DEFAULTS = (window.__COCKPIT__ && window.__COCKPIT__.fiDefaults) || {};
    let allJobs = [];
    let isConnectingWorker = false;
    let connectingTimeout = null;

    // --- Sidebar: SAP Dropdown ---
    let sapMenuOpen = false;

    function toggleSapMenu() {
      const toggle = document.getElementById('nav-sap-toggle');
      const sub = document.getElementById('nav-sap-sub');
      sapMenuOpen = !sapMenuOpen;
      toggle.classList.toggle('open', sapMenuOpen);
      sub.classList.toggle('open', sapMenuOpen);
    }

    // --- Sidebar: Processos Dropdown (2 níveis) ---
    let processosLoaded = false;
    let processosMenuOpen = false;
    const subprocessosCache = {};

    function toggleProcessosMenu() {
      const toggle = document.getElementById('nav-processos-toggle');
      const sub = document.getElementById('nav-processos-sub');
      processosMenuOpen = !processosMenuOpen;
      toggle.classList.toggle('open', processosMenuOpen);
      sub.classList.toggle('open', processosMenuOpen);
      if (processosMenuOpen && !processosLoaded) {
        loadProcessosMenu();
      }
    }

    function safeId(nome) {
      return nome.replace(/[^a-zA-Z0-9]/g, '_');
    }

    async function loadProcessosMenu() {
      const sub = document.getElementById('nav-processos-sub');
      const loading = document.getElementById('nav-processos-loading');
      try {
        const res = await fetch('/api/processes', { cache: 'no-store' });
        const data = await res.json();
        const processos = data.processes || [];
        if (loading) loading.remove();
        if (processos.length === 0) {
          sub.innerHTML = '<div class="nav-sub-loading">Nenhum processo encontrado.</div>';
        } else {
          sub.innerHTML = processos.map(p => `
            <div>
              <div class="nav-sub-item" id="proc-item-${safeId(p.nome)}" title="${escapeHtml(p.nome)}" onclick="toggleSubprocessosMenu('${p.nome.replace(/'/g,'\\&apos;')}', this)">
                <span>📁 ${escapeHtml(p.nome)}</span>
                <span class="nav-sub-arrow">&#9658;</span>
              </div>
              <div class="nav-sub-sub-list" id="subproc-list-${safeId(p.nome)}">
                <div class="nav-sub-loading" id="subproc-loading-${safeId(p.nome)}">A carregar...</div>
              </div>
            </div>
          `).join('');
        }
        processosLoaded = true;
      } catch (err) {
        if (loading) loading.textContent = 'Erro ao carregar processos.';
      }
    }

    async function toggleSubprocessosMenu(nome, el) {
      const subList = document.getElementById(`subproc-list-${safeId(nome)}`);
      const isOpen = subList.classList.contains('open');

      // Fecha todos os outros processos abertos
      document.querySelectorAll('.nav-sub-item.open').forEach(item => item.classList.remove('open'));
      document.querySelectorAll('.nav-sub-sub-list.open').forEach(list => list.classList.remove('open'));

      if (isOpen) return; // era o mesmo → já fechou acima

      el.classList.add('open');
      subList.classList.add('open');

      if (subprocessosCache[nome]) return; // já carregado

      try {
        const res = await fetch(`/api/subprocesses?processo=${encodeURIComponent(nome)}`, { cache: 'no-store' });
        const data = await res.json();
        const subs = data.subprocesses || [];
        const loading = document.getElementById(`subproc-loading-${safeId(nome)}`);
        if (loading) loading.remove();
        if (subs.length === 0) {
          subList.innerHTML = '<div class="nav-sub-loading">Sem subprocessos.</div>';
        } else {
          subList.innerHTML = subs.map(s => `
            <div class="nav-sub-sub-item" title="${escapeHtml(s.nome)}" onclick="abrirSubprocessoModal('${nome.replace(/'/g,'\\&apos;')}', '${s.nome.replace(/'/g,'\\&apos;')}')">
              ⚙️ ${escapeHtml(s.nome.replace(/^[A-Z]\.\s*/, '').replace(/\.py$/i, ''))}
            </div>
          `).join('');
        }
        subprocessosCache[nome] = true;
      } catch (err) {
        const loading = document.getElementById(`subproc-loading-${safeId(nome)}`);
        if (loading) loading.textContent = 'Erro ao carregar.';
      }
    }

    async function abrirSubprocessoModal(processo, subprocesso) {
      const modal = document.getElementById('modal-novo-job');
      if (!modal) return;
      document.getElementById('job-form').reset();
      _resetWebParams();
      const procSelect = document.getElementById('processo-select');
      if (procSelect) {
        procSelect.value = processo;
        _lockModalMenuFields();
        await loadSubprocessos(subprocesso);
        const reqSel = document.getElementById('request-option-select');
        if (reqSel) reqSel.dispatchEvent(new Event('change'));
      }
      // Carregar WEB_PARAMS do subprocesso (sem executar código SAP)
      try {
        const res = await fetch(`/api/subprocess-web-params?processo=${encodeURIComponent(processo)}&subprocesso=${encodeURIComponent(subprocesso)}`);
        const data = await res.json();
        _applyWebParamsToModal(data.params || null, data.config || null);
      } catch (e) {
        // Em caso de falha, mostrar formulário genérico normal
      }
      modal.classList.add('active');
    }

    function _applyWebParamsToModal(params, config) {
      const fileSec   = document.getElementById('file-picker-section');
      const reqSec    = document.getElementById('request-options-section');
      const ambSel    = document.getElementById('ambiente-select');
      const container = document.getElementById('web-params-container');

      // Aplicar WEB_CONFIG
      if (config) {
        if (config.show_file_picker === false && fileSec) fileSec.hidden = true;
        if (config.show_request_options === false && reqSec) reqSec.hidden = true;
        if (config.ambiente && ambSel) {
          ambSel.value = config.ambiente;
          const ambLabel = ambSel.closest('label');
          if (ambLabel) ambLabel.hidden = true;
          // badge de bloqueio para ambiente
          const badgeId = 'lock-badge-ambiente-select';
          if (!document.getElementById(badgeId)) {
            const badge = document.createElement('div');
            badge.id = badgeId;
            badge.className = 'select-lock-badge';
            badge.innerHTML = `🔒 Ambiente fixo: ${config.ambiente}`;
            ambSel.parentNode.insertBefore(badge, ambSel.nextSibling);
          }
        }
      }

      // Renderizar WEB_PARAMS
      if (!params || !params.length || !container) return;
      const html = params.map(p => {
        const req = p.required ? ' <span style="color:#ef4444">*</span>' : '';
        if (p.type === 'select') {
          const opts = (p.options || []).map(o =>
            `<option value="${_esc(o.value)}">${_esc(o.label)}</option>`
          ).join('');
          return `<label class="web-param-field" data-param-name="${_esc(p.name)}">
            ${_esc(p.label)}${req}
            <select data-web-param="${_esc(p.name)}" ${p.required ? 'required' : ''}>
              <option value="">Selecione...</option>
              ${opts}
            </select>
          </label>`;
        }
        const xform = p.transform === 'uppercase' ? 'style="text-transform:uppercase"' : '';
        return `<label class="web-param-field" data-param-name="${_esc(p.name)}">
          ${_esc(p.label)}${req}
          <input type="${p.type === 'password' ? 'password' : 'text'}"
            data-web-param="${_esc(p.name)}"
            placeholder="${_esc(p.placeholder || '')}"
            ${p.required ? 'required' : ''}
            ${xform}
            autocomplete="off">
        </label>`;
      }).join('');
      container.innerHTML = html;

      // uppercase transform: converter em tempo real
      container.querySelectorAll('input[data-web-param]').forEach(inp => {
        const label = inp.closest('label');
        const pName = label ? label.dataset.paramName : null;
        if (!pName) return;
        const pDef = params.find(p => p.name === pName);
        if (pDef && pDef.transform === 'uppercase') {
          inp.addEventListener('input', () => { inp.value = inp.value.toUpperCase(); });
        }
      });
    }

    function _esc(str) {
      return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    function _resetWebParams() {
      const container = document.getElementById('web-params-container');
      if (container) container.innerHTML = '';
      const fileSec = document.getElementById('file-picker-section');
      if (fileSec) fileSec.hidden = false;
      const reqSec = document.getElementById('request-options-section');
      if (reqSec) reqSec.hidden = false;
      const ambSel = document.getElementById('ambiente-select');
      if (ambSel) {
        const ambLabel = ambSel.closest('label');
        if (ambLabel) ambLabel.hidden = false;
      }
      const badge = document.getElementById('lock-badge-ambiente-select');
      if (badge) badge.remove();
    }

    function _lockModalMenuFields() {
      ['processo-select', 'subprocesso-select'].forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        el.classList.add('select-locked');
        // Oculta o label/campo correspondente no popup
        const label = el.closest('label');
        if (label) {
          label.hidden = true;
        }
        // Adiciona badge de bloqueio se ainda não existir
        const badgeId = `lock-badge-${id}`;
        if (!document.getElementById(badgeId)) {
          const badge = document.createElement('div');
          badge.id = badgeId;
          badge.className = 'select-lock-badge';
          badge.innerHTML = '🔒 Preenchido via menu lateral';
          el.parentNode.insertBefore(badge, el.nextSibling);
        }
      });
    }

    function resetModalLock() {
      ['processo-select', 'subprocesso-select'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
          el.classList.remove('select-locked');
          const label = el.closest('label');
          if (label) {
            label.hidden = false;
          }
        }
        const badge = document.getElementById(`lock-badge-${id}`);
        if (badge) badge.remove();
      });
      _resetWebParams();
    }

    let currentTab = 'running';
    let activeJobId = null;
    let activeJobDetailsExpanded = true;
    let expandedJobs = {};
    let activeLogQuery = '';
    let activeLogLevel = 'ALL';
    let activeLogState = 'ALL';
    let activeLogTime = 'ALL';
    let expandedLogItems = {};

    // Elementos DOM
    const modal = document.getElementById('modal-novo-job');
    const btnNewJob = document.getElementById('new-job-btn');
    const btnCloseModal = document.getElementById('close-modal-btn');
    const btnCancelModal = document.getElementById('cancel-modal-btn');
    btnNewJob.onclick = () => {
        resetModalLock();
        form.reset();
        requestOptionSelect.dispatchEvent(new Event('change'));
        processoSelect.dispatchEvent(new Event('change'));
        modal.classList.add('active');
    };
    btnCloseModal.onclick = () => {
        modal.classList.remove('active');
        resetModalLock();
    };
    btnCancelModal.onclick = () => {
        modal.classList.remove('active');
        resetModalLock();
    };

    // KPI Modal elements & close handlers
    const modalKpi = document.getElementById('modal-kpi-jobs');
    const btnCloseKpiModal = document.getElementById('close-kpi-modal-btn');
    const btnCloseKpiAction = document.getElementById('close-kpi-modal-action-btn');
    btnCloseKpiModal.onclick = () => modalKpi.classList.remove('active');
    btnCloseKpiAction.onclick = () => modalKpi.classList.remove('active');

    // Ligar Worker click listener
    const btnStartWorker = document.getElementById('start-worker-btn');
    if (btnStartWorker) {
        btnStartWorker.addEventListener('click', () => {
            // 1. Atualizar o estado e a UI primeiro
            isConnectingWorker = true;
            btnStartWorker.disabled = true;
            btnStartWorker.innerHTML = '<span class="server-icon">⏳</span><span>A ligar...</span>';
            btnStartWorker.style.opacity = '0.7';
            
            const ind = document.getElementById('worker-status-indicator');
            const txt = document.getElementById('worker-status-text');
            if (ind) ind.style.background = '#f59e0b'; // Amber/orange
            if (txt) txt.innerHTML = '⏳ A aguardar ligação...';

            if (connectingTimeout) clearTimeout(connectingTimeout);
            connectingTimeout = setTimeout(() => {
                if (isConnectingWorker) {
                    isConnectingWorker = false;
                    loadJobs();
                }
            }, 30000);

            // 2. Disparar o protocolo customizado após a atualização do DOM
            setTimeout(() => {
                window.location.href = 'sap-worker://start';
            }, 150);
        });
    }

    // Formulário
    const form = document.getElementById('job-form');
    const formMessage = document.getElementById('form-message');
    const processoSelect = document.getElementById('processo-select');
    const subprocessoSelect = document.getElementById('subprocesso-select');
    const caminhoFicheiroInput = document.getElementById('caminho-ficheiro-input');
    const selectExcelButton = document.getElementById('select-excel-button');
    const browserFileInput = document.getElementById('browser-file-input');
    const clearExcelButton = document.getElementById('clear-excel-button');
    const requestOptionSelect = document.getElementById('request-option-select');
    const requestNumberField = document.getElementById('request-number-field');
    const requestDescField = document.getElementById('request-desc-field');
    const requestTypeField = document.getElementById('request-type-field');
    const requestNumberInput = document.getElementById('request-number-input');
    const requestDescInput = document.getElementById('request-desc-input');
    const requestTypeSelect = document.getElementById('request-type-select');
    const ambienteSelect = document.getElementById('ambiente-select');
    const envDisplay = document.getElementById('env-display');

    // SAP Requests Modal DOM elements
    const modalSapRequests = document.getElementById('modal-sap-requests');
    const btnSearchRequest = document.getElementById('btn-search-request');
    const btnCloseSapRequests = document.getElementById('close-sap-requests-btn');
    const btnCloseSapRequestsBottom = document.getElementById('close-sap-requests-bottom-btn');
    const sapRequestsLoading = document.getElementById('sap-requests-loading');
    const sapRequestsResults = document.getElementById('sap-requests-results');
    const sapRequestsEmpty = document.getElementById('sap-requests-empty');
    const sapRequestsEmptyMsg = document.getElementById('sap-requests-empty-msg');
    const sapRequestsTableBody = document.getElementById('sap-requests-table-body');
    const sapRequestsSubtitle = document.getElementById('sap-requests-subtitle');
    
    let searchPollInterval = null;
    
    btnCloseSapRequests.onclick = () => {
        closeSapRequestsModal();
    };
    btnCloseSapRequestsBottom.onclick = () => {
        closeSapRequestsModal();
    };
    
    function closeSapRequestsModal() {
        modalSapRequests.classList.remove('active');
        if (searchPollInterval) {
            clearInterval(searchPollInterval);
            searchPollInterval = null;
        }
    }
    
    btnSearchRequest.onclick = async () => {
        const env = ambienteSelect.value;
        if (!env) {
            alert('Por favor, selecione primeiro o ambiente (DEV/QAD/PRD) para pesquisar as requests!');
            return;
        }
        
        const envLabel = ambienteSelect.options[ambienteSelect.selectedIndex].text;
        sapRequestsSubtitle.textContent = `Consultando a tabela E070 no SAP GUI para o ambiente: ${envLabel}...`;
        
        // Show modal and loading state
        modalSapRequests.classList.add('active');
        sapRequestsLoading.style.display = 'block';
        sapRequestsResults.style.display = 'none';
        sapRequestsEmpty.style.display = 'none';
        sapRequestsTableBody.innerHTML = '';
        
        try {
            // Queue the sap_search_requests task
            const payload = {
                task: 'sap_search_requests',
                params: { ambiente: env }
            };
            const response = await fetch('/api/jobs', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (!response.ok) throw new Error('Erro ao solicitar pesquisa de requests no worker.');
            const data = await response.json();
            const jobId = data.id;
            
            // Poll for job completion
            searchPollInterval = setInterval(async () => {
                try {
                    const statusRes = await fetch(`/api/jobs/${jobId}`);
                    if (!statusRes.ok) throw new Error('Erro ao consultar status do job de pesquisa.');
                    const job = await statusRes.json();
                    
                    if (job.state === 'succeeded') {
                        clearInterval(searchPollInterval);
                        searchPollInterval = null;
                        
                        let requestsList = [];
                        try {
                            requestsList = JSON.parse(job.status);
                        } catch (e) {
                            requestsList = [];
                        }
                        
                        sapRequestsLoading.style.display = 'none';
                        
                        if (!Array.isArray(requestsList) || requestsList.length === 0) {
                            sapRequestsEmptyMsg.innerHTML = `⚠️ Nenhuma request pendente ou ativa encontrada para o seu usuário no sistema <b>${job.params?.sap_system || 'SAP'}</b>.`;
                            sapRequestsEmpty.style.display = 'block';
                        } else {
                            sapRequestsTableBody.innerHTML = '';
                            requestsList.forEach(item => {
                                const tr = document.createElement('tr');
                                tr.className = 'kpi-table-row';
                                tr.style.cursor = 'pointer';
                                
                                // Make the entire row clickable to select
                                tr.onclick = () => {
                                    requestNumberInput.value = item.trkorr;
                                    closeSapRequestsModal();
                                };
                                
                                tr.innerHTML = `
                                    <td style="padding: 10px 8px; font-family: monospace; font-weight: bold; color: var(--text-primary);">${item.trkorr}</td>
                                    <td style="padding: 10px 8px; color: var(--text-secondary); max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.as4text}">${item.as4text || '<i>Sem descrição</i>'}</td>
                                    <td style="padding: 10px 8px; text-align: center;">
                                        <button class="btn btn-primary" style="padding: 4px 10px; font-size: 11px; border-radius: 6px;">Selecionar</button>
                                    </td>
                                `;
                                sapRequestsTableBody.appendChild(tr);
                            });
                            sapRequestsResults.style.display = 'block';
                        }
                    } else if (job.state === 'failed') {
                        clearInterval(searchPollInterval);
                        searchPollInterval = null;
                        sapRequestsLoading.style.display = 'none';
                        sapRequestsEmptyMsg.innerHTML = `❌ Falha ao consultar o SAP GUI:<br><span style="font-size: 11px; color: var(--danger); font-family: monospace;">${job.status || 'Erro desconhecido'}</span>`;
                        sapRequestsEmpty.style.display = 'block';
                    }
                } catch (err) {
                    console.error(err);
                }
            }, 1000);
            
        } catch (err) {
            sapRequestsLoading.style.display = 'none';
            sapRequestsEmptyMsg.textContent = `Erro ao iniciar pesquisa: ${err.message}`;
            sapRequestsEmpty.style.display = 'block';
        }
    };

    ambienteSelect.addEventListener('change', () => {
       envDisplay.textContent = ambienteSelect.options[ambienteSelect.selectedIndex].text || '---';
    });

    restoreLastExcelPath();

    processoSelect.addEventListener('change', loadSubprocessos);
    
    selectExcelButton.addEventListener('click', () => {
      browserFileInput.value = '';
      browserFileInput.click();
    });

    browserFileInput.addEventListener('change', async () => {
      const file = browserFileInput.files && browserFileInput.files[0];
      if (!file) return;
      formMessage.textContent = 'A enviar ficheiro...';
      try {
        const payload = new FormData();
        payload.append('file', file);
        const response = await fetch('/api/upload-file', { method: 'POST', body: payload });
        if (!response.ok) throw new Error('Erro ao enviar');
        const data = await response.json();
        setExcelPath(data.windows_path);
        formMessage.textContent = 'Ficheiro guardado.';
      } catch (err) {
        formMessage.textContent = err.message;
      }
    });

    clearExcelButton.addEventListener('click', () => setExcelPath(''));

    requestOptionSelect.addEventListener('change', updateRequestFieldsVisibility);

    function updateRequestFieldsVisibility() {
      const opt = requestOptionSelect.value;
      requestNumberField.hidden = opt !== '1';
      requestNumberInput.disabled = opt !== '1';
      requestNumberInput.required = opt === '1';

      requestDescField.hidden = opt !== '2';
      requestDescInput.disabled = opt !== '2';
      
      requestTypeField.hidden = opt !== '2';
      requestTypeSelect.disabled = opt !== '2';
    }

    function setExcelPath(path) {
      caminhoFicheiroInput.value = path || '';
      localStorage.setItem(STORAGE_KEY_EXCEL_PATH, path || '');
    }

    function restoreLastExcelPath() {
      setExcelPath(localStorage.getItem(STORAGE_KEY_EXCEL_PATH));
    }

    async function loadSubprocessos(selectedSub = '') {
      if (selectedSub instanceof Event) {
        selectedSub = '';
      }
      const p = processoSelect.value;
      subprocessoSelect.innerHTML = '<option value="">A carregar...</option>';
      if (!p) return subprocessoSelect.innerHTML = '<option value="">Selecione primeiro o processo</option>';
      try {
        const res = await fetch(`/api/subprocesses?processo=${encodeURIComponent(p)}`);
        const data = await res.json();
        subprocessoSelect.innerHTML = '<option value="">Selecione o subprocesso</option>';
        (data.subprocesses || []).forEach(sub => {
           const selectedAttr = (selectedSub && sub.nome === selectedSub) ? ' selected' : '';
           subprocessoSelect.innerHTML += `<option value="${sub.nome}"${selectedAttr}>${sub.label}</option>`;
        });
        if (selectedSub) {
          subprocessoSelect.value = selectedSub;
          subprocessoSelect.dispatchEvent(new Event('change'));
        }
      } catch {
        subprocessoSelect.innerHTML = '<option value="">Erro ao carregar</option>';
      }
    }

    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      document.getElementById('submit-job-btn').disabled = true;
      formMessage.textContent = 'A processar...';
      try {
        const fd = new FormData(form);
        // Incluir valores dos campos dinâmicos WEB_PARAMS
        document.querySelectorAll('#web-params-container [data-web-param]').forEach(el => {
          const name = el.dataset.webParam;
          if (name) fd.set(name, el.value || '');
        });
        const res = await fetch('/jobs', { method: 'POST', body: fd });
        if (!res.ok) throw new Error('Erro');
        const novoJob = await res.json();
        activeJobId = novoJob.id;
        formMessage.textContent = 'Criado com sucesso!';
        setTimeout(() => modal.classList.remove('active'), 1000);
        await loadJobs();
      } catch (err) {
        formMessage.textContent = 'Falha ao criar.';
      }
      document.getElementById('submit-job-btn').disabled = false;
    });

    document.getElementById('refresh-button').addEventListener('click', async () => {
      const btn = document.getElementById('refresh-button');
      if (btn.disabled) return;
      btn.disabled = true;
      btn.innerHTML = '<span style="display:inline-block; animation: spin 0.7s linear infinite;">↻</span>&nbsp;A atualizar...';
      try {
        await loadJobs();
      } finally {
        btn.disabled = false;
        btn.innerHTML = '↻ Atualizar';
      }
    });

    // Gestão de Abas da Queue
    document.querySelectorAll('.c-tab').forEach(tab => {
      tab.addEventListener('click', (e) => {
        document.querySelectorAll('.c-tab').forEach(t => t.classList.remove('active'));
        e.target.classList.add('active');
        currentTab = e.target.dataset.q;
        renderQueue();
      });
    });

    async function loadJobs() {
      try {
        const fetchOptions = { cache: 'no-store' };
        const [jobsRes, statusRes] = await Promise.all([
          fetch('/api/jobs?limit=50', fetchOptions),
          fetch('/api/worker/status', fetchOptions)
        ]);

        if (jobsRes.ok) {
          const data = await jobsRes.json();
          allJobs = data.jobs || [];
          calculateKPIs();
          renderQueue();
        }

        if (statusRes.ok) {
          const sData = await statusRes.json();
          const ind = document.getElementById('worker-status-indicator');
          const txt = document.getElementById('worker-status-text');
          const btn = document.getElementById('start-worker-btn');
          
          if (sData.status === 'online') {
            isConnectingWorker = false;
            if (connectingTimeout) {
                clearTimeout(connectingTimeout);
                connectingTimeout = null;
            }
            ind.style.background = 'var(--success)';
            txt.textContent = 'Online';
            if (btn) {
              btn.style.display = 'none';
              btn.disabled = false;
              btn.innerHTML = '<span class="server-icon">🖥️</span><span>Ligar Worker</span>';
              btn.style.opacity = '1';
            }
          } else {
            if (!isConnectingWorker) {
              ind.style.background = 'var(--danger)';
              txt.textContent = 'Offline';
              if (btn) {
                btn.style.display = 'flex';
                btn.disabled = false;
                btn.innerHTML = '<span class="server-icon">🖥️</span><span>Ligar Worker</span>';
                btn.style.opacity = '1';
              }
            }
          }
        }
      } catch(e) {}
    }

    function calculateKPIs() {
      const now = new Date();
      let running = 0, pending = 0, successToday = 0, errorToday = 0;
      let latestRunningProcess = '';
      
      let totalCompletedAllTime = 0;
      let successAllTime = 0;
      let totalDurationMs = 0;
      let latestJobTime = null;

      allJobs.forEach(job => {
        const isToday = new Date(job.created_at).toDateString() === now.toDateString();
        
        // Track latest job timestamp
        const jobDate = new Date(job.created_at);
        if (!latestJobTime || jobDate > latestJobTime) {
            latestJobTime = jobDate;
        }

        if (job.state === 'running') {
            running++;
            if (!latestRunningProcess) latestRunningProcess = job.params?.subprocesso || job.task;
        }
        if (job.state === 'pending') pending++;
        
        if (isToday && job.state === 'succeeded') successToday++;
        if (isToday && job.state === 'failed') errorToday++;
        
        // General stats for all-time rates
        if (job.state === 'succeeded') {
            successAllTime++;
            totalCompletedAllTime++;
            
            // Calculate duration
            const duration = new Date(job.updated_at) - new Date(job.created_at);
            if (duration > 0) {
                totalDurationMs += duration;
            }
        } else if (job.state === 'failed') {
            totalCompletedAllTime++;
        }
      });

      document.getElementById('kpi-running').textContent = running;
      document.getElementById('kpi-running-sub').textContent = running > 0 ? '+ ' + latestRunningProcess : 'Sistema livre';
      
      document.getElementById('kpi-pending').textContent = pending;
      
      document.getElementById('kpi-success').textContent = successToday;
      const totalCompletedToday = successToday + errorToday;
      const todayRate = totalCompletedToday > 0 ? Math.round((successToday / totalCompletedToday) * 100) : 100;
      document.getElementById('kpi-success-rate').textContent = `${todayRate}% sem erro`;

      document.getElementById('kpi-failed').textContent = errorToday;
      
      // Calculate avg duration
      let avgDurationStr = '0s';
      if (successAllTime > 0) {
          const avgMs = totalDurationMs / successAllTime;
          const avgSecs = Math.round(avgMs / 1000);
          if (avgSecs < 60) {
              avgDurationStr = `${avgSecs}s`;
          } else {
              const mins = Math.floor(avgSecs / 60);
              const secs = avgSecs % 60;
              avgDurationStr = `${mins}m ${secs}s`;
          }
      }
      document.getElementById('kpi-avg-time').textContent = avgDurationStr;
      
      // General success rate
      const generalRate = totalCompletedAllTime > 0 ? Math.round((successAllTime / totalCompletedAllTime) * 100) : 100;
      document.getElementById('kpi-success-rate-general').textContent = `${generalRate}%`;
      
      // Latest time
      let latestTimeStr = '---';
      if (latestJobTime) {
          latestTimeStr = latestJobTime.toLocaleDateString('pt-PT') + ' ' + latestJobTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
      }
      document.getElementById('kpi-latest-time').textContent = latestTimeStr;
    }

    function renderQueue() {
      const qContainer = document.getElementById('queue-list-container');
      if (!qContainer) {
          renderActiveJob();
          return;
      }
      qContainer.innerHTML = '';
      
      let filtered = [];
      if (currentTab === 'running') {
          filtered = allJobs.filter(j => j.state === 'running' || j.state === 'pending');
          
          // Se estamos na aba Ativos e o job ativo já não está a correr, limpa a seleção
          if (activeJobId) {
              const activeJobStillRunning = filtered.find(j => j.id === activeJobId);
              if (!activeJobStillRunning) {
                  activeJobId = null;
              }
          }
      } else {
          filtered = allJobs.filter(j => j.state === 'succeeded' || j.state === 'failed');
      }

      if (!filtered.length) {
         qContainer.innerHTML = '<div class="empty-state" style="padding: 20px;">Lista vazia</div>';
         if (!activeJobId && allJobs.length > 0 && currentTab === 'running') {
             // Try to select first history job if nothing running
             activeJobId = allJobs.find(j => j.state === 'succeeded' || j.state === 'failed')?.id || null;
         }
      }

      filtered.forEach(job => {
         const div = document.createElement('div');
         div.className = 'queue-item';
         div.style.cursor = 'pointer';
          let dataFila = new Date(job.created_at).toLocaleString('pt-PT', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' });
          let badgeClass = '';
          let badgeText = job.state;
          if (job.state === 'running') { badgeClass = 'badge-running'; badgeText = 'running'; }
          else if (job.state === 'pending') { badgeClass = 'badge-pending'; badgeText = 'pending'; }
          else if (job.state === 'succeeded') { badgeClass = 'badge-success'; badgeText = 'sucesso'; }
          else if (job.state === 'failed') { badgeClass = 'badge-error'; badgeText = 'erro'; }
          
          let proc = job.params ? (job.params.subprocesso || job.params.processo || job.task) : job.task;
          let isActive = (job.id === activeJobId) ? 'active' : '';
          let shortId = job.id.substring(0, 8);
          
          let html = `
          <div class="fila-job-status ${isActive}">
              <div class="fila-job-info">
                  <div class="fila-job-title">${escapeHtml(proc)}</div>
                  <div class="fila-job-meta"><span style="font-family:monospace">#${shortId}</span> · ${escapeHtml(job.ambiente || 'DEV')} · ${escapeHtml(job.worker_name ? (job.worker_name.split('-')[0]) : '')} · ${dataFila}</div>
              </div>
              <div class="status-badge ${badgeClass}">${badgeText}</div>
          </div>
          `; 
          div.innerHTML = html;
         
         div.onclick = () => {
             activeJobId = job.id;
             renderQueue(); // Re-render to highlight
         };
         qContainer.appendChild(div);

         // Auto-select first running job
         if (!activeJobId && job.state === 'running') {
             activeJobId = job.id;
         }
      });

      renderActiveJob();
    }

    function renderActiveJob() {
        const container = document.getElementById('active-jobs-container');
        const logContainer = document.getElementById('realtime-log-container');
        if (!container) return;
        // Guardar as posições de scroll atuais antes de limpar e renderizar
        const savedScrolls = {};
        container.querySelectorAll('.activity-log-scroll').forEach(el => {
            const isAtBottom = (el.scrollHeight - el.clientHeight - el.scrollTop) < 25;
            savedScrolls[el.id] = {
                scrollTop: el.scrollTop,
                isAtBottom: isAtBottom
            };
        });
        let savedRealtimeScroll = null;
        let realtimeIsAtBottom = true;
        if (logContainer) {
            realtimeIsAtBottom = (logContainer.scrollHeight - logContainer.clientHeight - logContainer.scrollTop) < 25;
            savedRealtimeScroll = logContainer.scrollTop;
        }

        if (activeJobId && !allJobs.some(j => j.id === activeJobId)) {
            activeJobId = null;
        }

        const active = allJobs.find(j => j.state === 'running' || j.state === 'pending');
        if (active) {
            activeJobId = active.id;
        } else if (!activeJobId && allJobs.length > 0) {
            activeJobId = allJobs[0].id;
        }

        // Find all active (running/pending) jobs
        let jobsToRender = allJobs.filter(j => j.state === 'running' || j.state === 'pending');

        // Always render the focused job, inserting it at the top if it is not already in the active list
        if (activeJobId && !jobsToRender.find(j => j.id === activeJobId)) {
            const selectedJob = allJobs.find(j => j.id === activeJobId);
            if (selectedJob) {
                jobsToRender.unshift(selectedJob);
            }
        }

        // If no jobs at all to display
        if (jobsToRender.length === 0) {
           container.innerHTML = `
              <div class="card" style="min-height: 200px; display: flex; align-items: center; justify-content: center;">
                 <div class="empty-state">Nenhum job em execução ou pendente no momento.</div>
              </div>
           `;
           if (logContainer) logContainer.innerHTML = 'Nenhum log disponível...';
           return;
       }

        // Clear children that are not in the new jobsToRender list or empty state
        const activeIds = jobsToRender.map(j => `card-job-${j.id}`);
        Array.from(container.children).forEach(child => {
            if (child.classList && child.classList.contains('card') && child.querySelector('.empty-state')) {
                container.innerHTML = '';
            } else if (child.id && child.id.startsWith('card-job-') && !activeIds.includes(child.id)) {
                container.removeChild(child);
            }
        });
        // Render each job
        jobsToRender.forEach((job, idx) => {
           const existingCard = document.getElementById(`card-job-${job.id}`);
           if (existingCard) {
               const lastUpdated = existingCard.getAttribute('data-updated-at');
               const lastState = existingCard.getAttribute('data-state');
               if (lastUpdated === job.updated_at && lastState === job.state) {
                   // Keep the card in the correct order in container if needed
                   if (container.children[idx] !== existingCard) {
                       container.insertBefore(existingCard, container.children[idx]);
                   }
                   return; // Skip rendering/updating this card!
               }
           }
          const p = job.params || {};
          const proc = p.processo || '';
          const sub = p.subprocesso || '';
          
          let displayTitleHtml = '';
          if (proc && sub) {
              displayTitleHtml = `
                 <span style="font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.05em; padding: 4px 10px; background: rgba(59, 130, 246, 0.08); border: 1px solid rgba(59, 130, 246, 0.15); color: #3B82F6; border-radius: 4px; display: inline-flex; align-items: center; gap: 4px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">📂 ${escapeHtml(proc)}</span>
                 <span style="color: var(--text-secondary); opacity: 0.4; font-size: 14px; margin: 0 4px; font-weight: normal;">/</span>
                 <span style="font-size: 17px; font-weight: 800; color: var(--text-primary); letter-spacing: -0.01em;">${escapeHtml(sub)}</span>
              `;
          } else if (proc || sub) {
              displayTitleHtml = `<span style="font-size: 17px; font-weight: 800; color: var(--text-primary); letter-spacing: -0.01em;">${escapeHtml(proc || sub)}</span>`;
          } else {
              displayTitleHtml = `<span style="font-size: 17px; font-weight: 800; color: var(--text-primary); letter-spacing: -0.01em;">${escapeHtml(job.task)}</span>`;
          }

          const data = new Date(job.created_at).toLocaleString('pt-PT');
          
          let sapInfoBadgesHtml = '';
          if (p.sap_user) {
              sapInfoBadgesHtml = `
                <span style="opacity: 0.4">|</span>
                <span><strong>Ambiente:</strong> ${escapeHtml(p.ambiente || job.ambiente || 'DEV')}</span>
                <span style="opacity: 0.4">|</span>
                <span><strong>Sistema:</strong> ${escapeHtml(p.sap_system || '')}</span>
                <span style="opacity: 0.4">|</span>
                <span><strong>Cliente:</strong> ${escapeHtml(p.sap_client || '')}</span>
                <span style="opacity: 0.4">|</span>
                <span><strong>Utilizador:</strong> ${escapeHtml(p.sap_user || '')}</span>
              `;
          }

          // Timeline steps HTML
          let logStr = job.log || '';
          logStr = logStr.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n');
          const lines = logStr.split('\n').filter(l => l.trim() !== '');

          let activeIndex = -1;
          if (job.state === 'running') {
              activeIndex = 0; // default a inicializar
              const searchString = logStr.toLowerCase();
              if (searchString.includes('sap gui scripting esta ativo') || searchString.includes('sessão sap pronta')) {
                  activeIndex = 1;
              }
              if (searchString.includes('processo selecionado') || searchString.includes('role a processar')) {
                  activeIndex = 2;
              }
              if (searchString.includes('iniciando processamento das roles') || searchString.includes('roles a processar') || searchString.includes('roles concluídas') || searchString.includes('role:')) {
                  activeIndex = 2;
              }
          } else if (job.state === 'succeeded') {
              activeIndex = 4;
          }

          const isCadeiaProcess = proc.toLowerCase().includes('cadeia') || sub.toLowerCase().includes('cadeia') || job.task.toLowerCase().includes('cadeia');
          let totalRoles = p.roles_count || 0;
          let currentRoleName = '';
          let currentRoleIndex = 0;
          let rolesFromLogs = [];
          let insideSummary = false;
          let completedRolesList = [];
          let currentRoleInLoop = '';
          let roleMetadata = {};
          let sapSbarText = '';
          let errorCount = 0;

          // Parse and scan logs for roles, progress, and status
          lines.forEach(line => {
              const tr = line.trim();
              if (!tr) return;
              // Check for error lines to increment errorCount
              if (tr.includes('🔴 ERRO') || tr.includes('❌ SAP Erro:') || tr.includes('❌ Erro') || tr.includes('❌ Falha') || tr.startsWith('❌')) {
                  errorCount++;
              }
              // Parse total chains to verify from logs if not set
              const totalCadeiasMatch = tr.match(/🔍\s*Total de cadeias a verificar:\s*(\d+)/i);
              if (totalCadeiasMatch) {
                  totalRoles = parseInt(totalCadeiasMatch[1]);
              }
              // Parse search chains for Validar Cadeia de Pesquisa.py
              const cadeiaVerifyMatch = tr.match(/^([✅❌])\s*([^-\n]+)\s*-\s*(.*)$/);
              if (cadeiaVerifyMatch) {
                  const isSuccess = (cadeiaVerifyMatch[1] === '✅');
                  const cadeiaName = cadeiaVerifyMatch[2].trim();
                  const cadeiaStatus = cadeiaVerifyMatch[3].trim();
                  
                  // Keep the chain name as requested by the user, incorporating the status
                  const displayCadeiaName = `${cadeiaName} - ${cadeiaStatus}`;
                  
                  if (!rolesFromLogs.some(r => r.name === displayCadeiaName)) {
                      rolesFromLogs.push({ name: displayCadeiaName });
                  }
                  
                  if (isSuccess) {
                      completedRolesList.push(displayCadeiaName);
                  }
                  
                  if (!roleMetadata[displayCadeiaName]) {
                      roleMetadata[displayCadeiaName] = {
                          tcodes: null,
                          actions: 1,
                          duration: '',
                          statusText: cadeiaStatus,
                          success: isSuccess
                      };
                  }
                  
                  // Update progress index to count how many chains we've verified
                  currentRoleIndex = rolesFromLogs.length;
              }
              
              // 1. Parse the full checklist of roles from the summary in logs if available
              if (tr.includes('Roles a processar') || tr.includes('Resumo das Roles') || tr.includes('Roles a eliminar')) {
                  insideSummary = true;
              } else if (tr.includes('Deseja lançar') || tr.includes('INICIANDO ROLE') || tr.includes('A processar role') || tr.includes('Executing') || tr.includes('Bloco 2') || tr.includes('===') || tr.includes('tratar_popup') || tr.includes('[Etapa') || tr.includes('Acesso ao SAP')) {
                  insideSummary = false;
              }
              
              if (insideSummary) {
                  const matchRole = tr.match(/^(?:-\s*|(?:\d+)\.\s*)([A-Za-z0-9_\\\-]+)/);
                  if (matchRole) {
                      const rName = matchRole[1];
                      if (!rolesFromLogs.some(r => r.name === rName)) {
                          rolesFromLogs.push({ name: rName });
                      }
                  }
              }
              
              // 2. Parse active role, index, and total roles from bracket logs like ▶ [1/3]
              const bracketMatch = tr.match(/▶\s*\[(\d+)\/(\d+)\]/);
              if (bracketMatch) {
                  const idx = parseInt(bracketMatch[1]);
                  const tot = parseInt(bracketMatch[2]);
                  if (idx > currentRoleIndex) {
                      currentRoleIndex = idx;
                  }
                  if (tot > totalRoles) {
                      totalRoles = tot;
                  }
                  
                  const nameMatch = tr.match(/(?:ROLE|CADEIA):\s*([A-Za-z0-9_\\\-]+)/i);
                  if (nameMatch && nameMatch[1]) {
                      currentRoleInLoop = nameMatch[1];
                      currentRoleName = currentRoleInLoop;
                      if (!rolesFromLogs.some(r => r.name === currentRoleInLoop)) {
                          rolesFromLogs.push({ name: currentRoleInLoop });
                      }

                      // Initialize metadata for this role
                      const tcodesMatch = tr.match(/TCODEs:\s*(\d+)/i);
                      const tcodesVal = tcodesMatch ? parseInt(tcodesMatch[1]) : null;
                      if (!roleMetadata[currentRoleInLoop]) {
                          roleMetadata[currentRoleInLoop] = {
                              tcodes: tcodesVal,
                              actions: 0,
                              duration: ''
                          };
                      } else if (tcodesVal !== null) {
                          roleMetadata[currentRoleInLoop].tcodes = tcodesVal;
                      }
                  }
              }
              
              // Track action counts
              if (currentRoleInLoop && (tr.startsWith('├─') || tr.startsWith('└─'))) {
                  if (!roleMetadata[currentRoleInLoop]) {
                      roleMetadata[currentRoleInLoop] = { tcodes: null, actions: 0, duration: '' };
                  }
                  roleMetadata[currentRoleInLoop].actions++;
              }

              // Track duration / execution times
              if (currentRoleInLoop && (tr.includes('🟢 SUCESSO') || tr.includes('🔴 ERRO') || tr.includes('Role concluida') || tr.includes('Role concluída') || tr.includes('tratada por completo'))) {
                  const durationMatch = tr.match(/\(Tempo:\s*([^)]+)\)/i);
                  if (durationMatch) {
                      if (!roleMetadata[currentRoleInLoop]) {
                          roleMetadata[currentRoleInLoop] = { tcodes: null, actions: 0, duration: '' };
                      }
                      roleMetadata[currentRoleInLoop].duration = durationMatch[1].trim();
                  }
              }

              // 3. Track role success completion status
              if (tr.includes('🟢 SUCESSO') || tr.includes('Role concluida:') || tr.includes('Role concluída:') || tr.includes('[OK] Role concluida:') || tr.includes('[OK] Role concluída:')) {
                  if (currentRoleInLoop) {
                      completedRolesList.push(currentRoleInLoop);
                  }
                  const parts = tr.split(':');
                  if (parts[1] && (parts[0].includes('Role concluida') || parts[0].includes('Role concluída'))) {
                      completedRolesList.push(parts[1].trim());
                  }
              }

              // Parse SAP status bar message if present
              if (tr.includes('[SAP_SBAR]')) {
                  const sVal = tr.split('[SAP_SBAR]')[1].trim();
                  if (sVal) {
                      sapSbarText = sVal;
                  }
              } else if (tr.toUpperCase().startsWith('STATUS:')) {
                  const sVal = tr.substring(7).trim();
                  if (sVal) {
                      sapSbarText = sVal;
                  }
              }
          });
          if (totalRoles === 0 && rolesFromLogs.length > 0) {
              totalRoles = rolesFromLogs.length;
          }
          if (job.state === 'failed' && errorCount === 0) {
              errorCount = 1;
          }

          // Unique values in completed list
          completedRolesList = [...new Set(completedRolesList)];

          // 4. Fallback search for current active role from end of logs if not captured by bracket log yet
          if (job.state === 'running' && !currentRoleName) {
              for (let i = lines.length - 1; i >= 0; i--) {
                  const line = lines[i];
                  if (line.includes('Role [') || line.includes('role:')) {
                      const match = line.match(/(?:Role\s+\[|role:\s*)([A-Za-z0-9_\\-]+)/i);
                      if (match && match[1]) {
                          currentRoleName = match[1];
                          break;
                      }
                  }
              }
              if (!currentRoleName) {
                  for (let i = lines.length - 1; i >= 0; i--) {
                      if (lines[i].includes('processar') && lines[i].includes('Roles')) {
                          break;
                      }
                      if (lines[i].trim().startsWith('- ')) {
                          const parts = lines[i].split(':');
                          currentRoleName = parts[0].replace('-', '').trim();
                          break;
                      }
                  }
              }

              for (let i = lines.length - 1; i >= 0; i--) {
                  const line = lines[i];
                  const idxMatch = line.match(/\b(\d+)\/(\d+)\b/);
                  if (idxMatch) {
                      currentRoleIndex = parseInt(idxMatch[1]);
                      if (!totalRoles) {
                          totalRoles = parseInt(idxMatch[2]);
                      }
                      break;
                  }
              }
          }

          const concludedRoles = (job.state === 'succeeded') ? totalRoles : currentRoleIndex;

          let roleHistoryHtml = '';
          let totalActions = 0;
          if (totalRoles > 0) {
              let rows = '';
              const allRolesParam = (p.roles && p.roles.length > 0) ? p.roles : rolesFromLogs;
              allRolesParam.forEach((r, idx) => {
                  const isConcluida = completedRolesList.includes(r.name) || (job.state === 'succeeded');
                  const isActive = (job.state === 'running' && r.name === currentRoleName);
                  let statusBadge = '';
                  
                  if (isCadeiaProcess) {
                      const meta = roleMetadata[r.name] || {};
                      if (meta.success === true) {
                          statusBadge = '<span class="role-table-badge success">✓ ENCONTRADA</span>';
                      } else if (meta.success === false) {
                          statusBadge = '<span class="role-table-badge failed">✗ NÃO ENCONTRADA</span>';
                      } else {
                          statusBadge = '<span class="role-table-badge pending">PENDENTE</span>';
                      }
                  } else {
                      if (isConcluida) {
                          statusBadge = '<span class="role-table-badge success">✓ CONCLUÍDA</span>';
                      } else if (isActive) {
                          statusBadge = '<span class="role-table-badge active">⚙️ PROCESSANDO</span>';
                      } else {
                          statusBadge = '<span class="role-table-badge pending">PENDENTE</span>';
                      }
                  }
                  const meta = roleMetadata[r.name] || {};
                  const actionsCount = meta.actions || 0;
                  totalActions += actionsCount;
                  
                  let completedTimeStr = '---';
                  const isProcessed = isConcluida || (isCadeiaProcess && meta.success !== undefined);
                  if (isProcessed) {
                      let compLineIdx = -1;
                      for (let i = 0; i < lines.length; i++) {
                          const tr = lines[i].trim();
                          if (isCadeiaProcess) {
                              if (tr.includes(r.name.split(' - ')[0])) {
                                  compLineIdx = i;
                                  break;
                              }
                          } else {
                              if ((tr.includes('🟢 SUCESSO') || tr.includes('Role concluida:') || tr.includes('Role concluída:')) && tr.includes(r.name)) {
                                  compLineIdx = i;
                                  break;
                              }
                          }
                      }
                      if (compLineIdx !== -1) {
                          const ts = getTimestampForLine(job, compLineIdx, lines.length);
                          const startDay = new Date(job.created_at).toLocaleDateString('pt-PT');
                          completedTimeStr = `${startDay}, ${ts}`;
                      } else {
                          const dateObj = new Date(job.updated_at);
                          completedTimeStr = `${dateObj.toLocaleDateString('pt-PT')}, ${formatTimeOnly(dateObj)}`;
                      }
                  } else if (isActive) {
                      completedTimeStr = 'A processar...';
                  }
                  
                  if (!isCadeiaProcess) {
                      rows += `
                         <tr class="role-history-table-row ${isActive ? 'active-row' : ''}">
                            <td style="padding: 8px; font-weight: bold; color: var(--text-secondary); width: 60px;">${String(idx+1).padStart(2, '0')}</td>
                            <td style="padding: 8px; font-weight: 600; color: ${isActive ? 'var(--primary)' : 'var(--text-primary)'};">${escapeHtml(r.name)}</td>
                            <td style="padding: 8px; font-family: monospace; color: var(--text-secondary);">${actionsCount}</td>
                            <td style="padding: 8px;">${statusBadge}</td>
                            <td style="padding: 8px; color: var(--text-secondary); font-size: 11px;">${completedTimeStr}</td>
                         </tr>
                      `;
                  }
              });
              
              if (!isCadeiaProcess) {
                  roleHistoryHtml = `
                     <div style="flex: 1.2; min-width: 320px;">
                        <div style="display:flex; justify-content:space-between; margin-bottom:12px; align-items: center;">
                           <span style="font-weight:bold; color:var(--text-primary); font-size:14px;">Histórico de Execução das Roles</span>
                           <span style="font-size:11px; font-weight:bold; color:var(--success);">${concludedRoles}/${totalRoles} Concluídas</span>
                        </div>
                        <div class="role-history-scroll custom-scroll" style="max-height:320px; overflow-y:auto; border: 1px solid var(--border-color); border-radius: 8px; background: rgba(0,0,0,0.01);">
                           <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 12px;">
                              <thead>
                                 <tr style="background: rgba(0,0,0,0.02); border-bottom: 1px solid var(--border-color); font-weight: bold; color: var(--text-secondary);">
                                    <th style="padding: 8px; width: 60px;">Ordem</th>
                                    <th style="padding: 8px;">Role</th>
                                    <th style="padding: 8px; width: 60px;">Ações</th>
                                    <th style="padding: 8px; width: 110px;">Status</th>
                                    <th style="padding: 8px; width: 140px;">Concluído em</th>
                                 </tr>
                              </thead>
                              <tbody>
                                 ${rows}
                              </tbody>
                           </table>
                        </div>
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 12px; font-size: 11px; color: var(--text-secondary);">
                           <span>Exibindo ${allRolesParam.length} de ${allRolesParam.length} roles</span>
                           <a href="#" style="color: var(--primary); text-decoration: none; font-weight: 600;" onclick="event.preventDefault();">Ver todas as roles &rarr;</a>
                        </div>
                     </div>
                  `;
              }
          }
          let stageLabels = ['Inicialização e Logon SAP', 'Abertura e Verificação de Tela', 'Processamento Principal / Ações GUI', 'Conclusão / Geração de Logs'];
          const subLower = sub.toLowerCase();
          if (subLower.includes('pfcg_composta') || (subLower.includes('pfcg') && subLower.includes('composta'))) {
              stageLabels = ['Preparação e Dados', 'Atribuir Roles', 'Ordem de Transporte'];
          } else if (subLower.includes('pfcg_create') || (subLower.includes('pfcg') && subLower.includes('create'))) {
              stageLabels = ['Preparação e Dados', 'Atribuição de TCODs', 'Gerar Perfil', 'Ordem de Transporte'];
          } else if (subLower.includes('pfcg_delete') || (subLower.includes('pfcg') && subLower.includes('delete'))) {
              stageLabels = ['Leitura do Excel', 'Acesso ao SAP', 'Eliminação das Funções', 'Atualização do Excel'];
          } else if (subLower.includes('pfcg_authority') || (subLower.includes('pfcg') && subLower.includes('authority'))) {
              stageLabels = ['Validação da Role', 'Inserção MASSVAL', 'Ordem de Transporte'];
          } else if (subLower.includes('cua_adicionar') || subLower.includes('adicionar')) {
              stageLabels = ['Leitura do Excel', 'Pesquisa de Utilizadores', 'Atribuição no SAP CUA', 'Gravação de Resultados'];
          } else if (subLower.includes('cua_enddate') || subLower.includes('enddate') || subLower.includes('validade')) {
              stageLabels = ['Leitura do Excel', 'Acesso ao SAP CUA', 'Bloqueio / Data Fim', 'Gravação de Resultados'];
          } else if (subLower.includes('cua_remove') || subLower.includes('remove')) {
              stageLabels = ['Leitura do Excel', 'Acesso ao SAP CUA', 'Remoção de Perfis', 'Gravação de Resultados'];
          }
          lines.forEach(line => {
              for (let s = 1; s <= 10; s++) {
                  const tag = `[Etapa ${s}]`;
                  if (line.includes(tag)) {
                      const idx = line.indexOf(tag);
                      const labelText = line.substring(idx + tag.length).trim();
                      if (labelText && labelText.length < 60) {
                          while (stageLabels.length < s) {
                              stageLabels.push(`Etapa ${stageLabels.length + 1}`);
                          }
                          stageLabels[s - 1] = labelText;
                      }
                  }
              }
          });
          const numStages = stageLabels.length;
          const stepClasses = Array(numStages).fill('pending');
          let currentRoleLogs = [];
          let currentRoleStartIndex = -1;
          for (let i = lines.length - 1; i >= 0; i--) {
              if (lines[i].toUpperCase().includes('INICIANDO ROLE')) {
                  currentRoleStartIndex = i;
                  break;
              }
          }
          if (currentRoleStartIndex !== -1) {
              currentRoleLogs = lines.slice(currentRoleStartIndex);
          } else {
              currentRoleLogs = lines;
          }
          if (job.state === 'running') {
              let activeSubstep = 0;
              let currentRoleHasError = false;
              let currentRoleHasSuccess = false;
              
              currentRoleLogs.forEach(line => {
                  for (let s = 1; s <= numStages; s++) {
                      if (line.includes(`[Etapa ${s}]`)) {
                          activeSubstep = s;
                      }
                  }
                  
                  if (line.includes('🔴 ERRO') || line.includes('❌ SAP Erro:') || line.includes('❌ Erro') || line.includes('❌ Falha')) {
                      currentRoleHasError = true;
                  }
                  if (line.includes('🟢 SUCESSO') || line.includes('Role concluida:') || line.includes('Role concluída:') || line.includes('Role tratada por completo')) {
                      currentRoleHasSuccess = true;
                  }
              });
              
              if (currentRoleHasSuccess) {
                  for (let s = 0; s < numStages; s++) stepClasses[s] = 'completed';
              } else if (currentRoleHasError) {
                  for (let s = 0; s < numStages; s++) {
                      if (s + 1 < activeSubstep) {
                          stepClasses[s] = 'completed';
                      } else if (s + 1 === activeSubstep) {
                          stepClasses[s] = 'failed';
                      } else {
                          stepClasses[s] = 'pending';
                      }
                  }
                  if (activeSubstep === 0) {
                      stepClasses[0] = 'failed';
                  }
              } else {
                  for (let s = 0; s < numStages; s++) {
                      if (s + 1 < activeSubstep) {
                          stepClasses[s] = 'completed';
                      } else if (s + 1 === activeSubstep) {
                          stepClasses[s] = 'active';
                      } else {
                          stepClasses[s] = 'pending';
                      }
                  }
                  if (activeSubstep === 0) {
                      stepClasses[0] = 'active';
                  }
              }
          } else if (job.state === 'succeeded') {
              for (let s = 0; s < numStages; s++) stepClasses[s] = 'completed';
          } else if (job.state === 'failed') {
              let activeSubstep = 0;
              lines.forEach(line => {
                  for (let s = 1; s <= numStages; s++) {
                      if (line.includes(`[Etapa ${s}]`)) {
                          activeSubstep = s;
                      }
                  }
              });
              for (let s = 0; s < numStages; s++) {
                  if (s + 1 < activeSubstep) {
                      stepClasses[s] = 'completed';
                  } else if (s + 1 === activeSubstep) {
                      stepClasses[s] = 'failed';
                  } else {
                      stepClasses[s] = 'pending';
                  }
              }
              if (activeSubstep === 0) {
                  stepClasses[0] = 'failed';
              }
          }
          // Group lines into steps
          let taskLogCount = 0;
          let steps = [];
          let currentStep = null;
          
          let cmdArgs = [];
          if (p.ambiente) cmdArgs.push(`--ambiente ${p.ambiente}`);
          if (p.sap_system) cmdArgs.push(`--system ${p.sap_system}`);
          if (p.sap_client) cmdArgs.push(`--client ${p.sap_client}`);
          let cmdStr = `python Processos/${job.task} ${cmdArgs.join(' ')}`;
          const initTs = getTimestampForLine(job, 0, lines.length);
          
          currentStep = {
              title: 'Inicializar Subprocesso',
              dotColor: 'var(--primary)',
              inContent: cmdStr,
              outLines: [],
              lineIndex: 0,
              timestamp: initTs
          };
          steps.push(currentStep);
          
           lines.forEach((line, idx) => {
               const tr = line.trim();
               if (!tr) return;
               
               const lower = tr.toLowerCase();
               
               // Skip raw divider lines (equal signs, hyphens, etc.)
               if (/^[=\-_ ]+$/.test(tr)) {
                   return;
               }
               
               // Skip unnecessary system/redundant logs
               if (lower.includes('ficheiro .env') || 
                   lower.includes('credenciais carregadas') || 
                   lower.includes('chave_password') ||
                   lower.includes('verificar disponibilidade do sap gui') ||
                   (lower.includes('sap gui scripting') && lower.includes('ativo')) ||
                   lower.includes('sessão sap encontrada') ||
                   (lower.includes('[sap_sbar]') && lower.includes('visão de atualização'))) {
                   return;
               }
               
               let cleanLine = line;
               
               // Clean up Windows COM exception tuples and trailing commas
               cleanLine = cleanLine.replace(/\(-?\d+,\s*'([^']*)'.*?\)/g, '$1');
               cleanLine = cleanLine.replace(/\('([^']*)',\)/g, '$1');
               
               let isNewStep = false;
               let newStepTitle = '';
               let newStepColor = 'var(--primary)';
               
               const roleMatch = cleanLine.match(/▶\s*\[\d+\/\d+\]\s*(?:ROLE|CADEIA):\s*([A-Za-z0-9_\\-]+)/i);
               const chainMatch = cleanLine.match(/^([✅❌])\s*([^-]+)\s*-\s*(.*)$/);
               const phaseMatch = cleanLine.match(/===+\s*(FASE\s+\d+:\s+[^=]+)\s*===+/i) || cleanLine.match(/(FASE\s+\d+:\s+.*)/i);
               
               if (roleMatch) {
                   isNewStep = true;
                   newStepTitle = cleanLine.trim();
                   newStepColor = 'var(--primary)';
               } else if (phaseMatch) {
                   isNewStep = true;
                   let titleStr = phaseMatch[1].trim();
                   if (titleStr.toUpperCase() === titleStr) {
                       titleStr = titleStr.toLowerCase().replace(/^(fase\s+\d+:?)\s*(.*)$/, (match, p1, p2) => {
                           const part1 = p1.charAt(0).toUpperCase() + p1.slice(1);
                           const prepositions = ['de', 'do', 'da', 'dos', 'das', 'em', 'para', 'com', 'por', 'a', 'o', 'e'];
                           const words = p2.split(/\s+/).map((w, index) => {
                               if (index === 0) return w.charAt(0).toUpperCase() + w.slice(1);
                               if (prepositions.includes(w)) return w;
                               return w.charAt(0).toUpperCase() + w.slice(1);
                           });
                           return part1 + ' ' + words.join(' ');
                       });
                   }
                   newStepTitle = titleStr;
                   newStepColor = 'var(--primary)';
               } else if (chainMatch) {
                   isNewStep = true;
                   newStepTitle = `Validar Cadeia: ${chainMatch[2].trim()}`;
                   newStepColor = chainMatch[1] === '✅' ? 'var(--success)' : 'var(--danger)';
               } else if (cleanLine.startsWith('[Etapa ') || cleanLine.includes('[Etapa ')) {
                   isNewStep = true;
                   newStepTitle = cleanLine.trim();
                   newStepColor = 'var(--primary)';
               } else if (cleanLine.includes('🔴 ERRO') || cleanLine.includes('❌ SAP Erro:') || cleanLine.includes('❌ Erro') || cleanLine.includes('❌ Falha') || cleanLine.startsWith('Traceback') || cleanLine.startsWith('ERRO:')) {
                   if (currentStep) {
                       currentStep.dotColor = 'var(--danger)';
                   }
               }
               
               if (isNewStep) {
                   currentStep = {
                       title: newStepTitle,
                       dotColor: newStepColor,
                       inContent: cleanLine.trim(),
                       outLines: [],
                       lineIndex: idx,
                       timestamp: getTimestampForLine(job, idx, lines.length)
                   };
                   steps.push(currentStep);
               } else {
                   if (currentStep) {
                       currentStep.outLines.push(cleanLine);
                       
                       const lowerLine = cleanLine.toLowerCase();
                       if (cleanLine.includes('🟢 SUCESSO') || cleanLine.includes('✅') || lowerLine.includes('sucesso') || lowerLine.includes('concluída') || lowerLine.includes('concluido')) {
                           currentStep.dotColor = 'var(--success)';
                       } else if (cleanLine.includes('🔴 ERRO') || cleanLine.includes('❌') || lowerLine.includes('falha') || lowerLine.includes('erro') || lowerLine.includes('exception') || lowerLine.includes('traceback') || cleanLine.startsWith('ERRO:')) {
                           currentStep.dotColor = 'var(--danger)';
                       } else if (cleanLine.includes('⚠️') || lowerLine.includes('warning') || lowerLine.includes('aviso') || cleanLine.startsWith('[WARN]')) {
                           if (currentStep.dotColor !== 'var(--danger)' && currentStep.dotColor !== 'var(--success)') {
                               currentStep.dotColor = 'var(--warning)';
                           }
                       }
                   }
               }
           });
          
          let taskLogRows = '';
          steps.forEach((step, idx) => {
              // Determine level
              let level = 'INFO';
              if (step.dotColor === 'var(--danger)' || step.dotColor === 'red' || step.dotColor === '#ef4444') {
                  level = 'ERROR';
              } else if (step.dotColor === 'var(--warning)' || step.dotColor === 'orange' || step.dotColor === '#f59e0b') {
                  level = 'WARN';
              }
              
              let outContent = step.outLines.join('\n').trim();
              
              // Show terminal box for subprocess initialization OR if traceback is detected
              const isCommandStep = step.title === 'Inicializar Subprocesso';
              const hasTraceback = outContent.includes('Traceback') || outContent.includes('Exception') || outContent.toLowerCase().includes('erro:');
              const showCard = isCommandStep || hasTraceback;
              
              let stepCardHtml = '';
              let subTextHtml = '';
              
              if (showCard) {
                  let outSectionHtml = outContent 
                      ? `<div style="display: flex; gap: 8px; margin-top: 6px; padding-top: 6px; border-top: 1px dashed var(--border-color, #e5e7eb);"><span style="color: var(--text-secondary); font-weight: bold; flex-shrink: 0; width: 32px; user-select: none; font-size: 11px;">OUT</span><span style="color: var(--text-primary); white-space: pre-wrap; word-break: break-all; font-family: Consolas, Monaco, monospace; font-size: 11px;">${escapeHtml(outContent)}</span></div>`
                      : '';
                      
                  stepCardHtml = `
                      <div class="timeline-detail-card" style="margin-top: 8px; margin-bottom: 8px; background: #f6f8fa; border: 1px solid var(--border-color, #e5e7eb); border-radius: 6px; padding: 10px 14px; font-family: Consolas, Monaco, 'Courier New', monospace; font-size: 11px; line-height: 1.4; color: var(--text-primary); box-shadow: inset 0 1px 2px rgba(0,0,0,0.02);">
                          <div style="display: flex; gap: 8px;"><span style="color: var(--text-secondary); font-weight: bold; flex-shrink: 0; width: 32px; user-select: none; font-size: 11px;">IN</span><span style="white-space: pre-wrap; word-break: break-all; color: #032f62; font-size: 11px;">${escapeHtml(step.inContent)}</span></div>
                          ${outSectionHtml}
                      </div>
                  `;
              } else {
                  let linesToRender = step.outLines.filter(l => l.trim() !== '');
                  if (linesToRender.length > 0) {
                      let groupedLines = [];
                      let currentGroup = [];
                      const processStepGroup = (group) => {
                          if (group.length === 1) {
                              return group[0].replace(/^(\s*)\|-\s*/, '$1├─ ').replace(/^(\s*)\| - \s*/, '$1├─ ');
                          }
                          const firstLine = group[0];
                          const prefixMatch = firstLine.match(/^([|\-├─\s]+)/);
                          const prefix = prefixMatch ? prefixMatch[0].replace('|-', '├─').replace('| -', '├─') : '├─ ';
                          const cleanedParts = group.map(line => {
                              return line.replace(/^([|\-├─\s]+)/, '')
                                         .replace(/\.\.\.$/, '')
                                         .trim();
                          });
                          return prefix + cleanedParts.join(' -> ');
                      };
                      
                      linesToRender.forEach(l => {
                          const trimmed = l.trim();
                          const isGroupable = trimmed.startsWith('|-') || trimmed.startsWith('├─') || trimmed.startsWith('| -');
                          if (isGroupable) {
                              currentGroup.push(l);
                          } else {
                              if (currentGroup.length > 0) {
                                  groupedLines.push(processStepGroup(currentGroup));
                                  currentGroup = [];
                              }
                              groupedLines.push(l);
                          }
                      });
                      if (currentGroup.length > 0) {
                          groupedLines.push(processStepGroup(currentGroup));
                      }
                      linesToRender = groupedLines;

                      subTextHtml = `
                          <div style="display: flex; flex-direction: column; gap: 4px; margin-top: 4px; padding-left: 2px;">
                              ${linesToRender.map(l => {
                                  const isThought = l.toLowerCase().includes('thought') || l.toLowerCase().startsWith('vou ') || l.toLowerCase().includes('analisando') || l.toLowerCase().includes('validar existencia');
                                  const textColor = isThought ? 'var(--text-secondary)' : 'var(--text-primary)';
                                  const fontSize = isThought ? '11px' : '12px';
                                  const icon = isThought ? '💡' : '·';
                                  return `<div style="font-size: ${fontSize}; color: ${textColor}; display: flex; gap: 6px; align-items: flex-start; line-height: 1.4;">
                                      <span style="opacity: 0.6; flex-shrink: 0;">${icon}</span>
                                      <span>${escapeHtml(l)}</span>
                                  </div>`;
                              }).join('')}
                          </div>
                      `;
                  }
              }
              
              const startDate = getRawDateForLine(job, step.lineIndex, lines.length);
              let endDate;
              if (idx < steps.length - 1) {
                  endDate = getRawDateForLine(job, steps[idx + 1].lineIndex, lines.length);
              } else {
                  endDate = (job.state === 'running' || job.state === 'pending') 
                      ? new Date() 
                      : new Date(job.updated_at);
              }
              
              const startTimeStr = step.timestamp;
              const endTimeStr = formatTimeOnly(endDate);
              
              const diffMs = Math.max(0, endDate - startDate);
              const diffSecs = Math.round(diffMs / 1000);
              let durationStr = '';
              if (diffSecs < 60) {
                  durationStr = `${diffSecs}s`;
              } else {
                  const mins = Math.floor(diffSecs / 60);
                  const secs = diffSecs % 60;
                  durationStr = `${mins}m ${secs}s`;
              }
              
              const timeRangeStr = `${startTimeStr} · ${durationStr} ago`;
              
              let dotColor = '#2da44e'; // green
              if (level === 'ERROR') {
                  dotColor = '#cf222e'; // red
              } else if (level === 'WARN') {
                  dotColor = '#9a6700'; // orange
              } else if (step.title.toLowerCase().includes('thought') || step.title.toLowerCase().includes('info')) {
                  dotColor = '#8c959f'; // gray
              }
              
              const hasDetails = (subTextHtml.trim() !== '' || stepCardHtml.trim() !== '');
              let titleWithToggleHtml = '';
              let detailsWrapperHtml = '';
              
              if (hasDetails) {
                  titleWithToggleHtml = `
                      <div style="display: flex; align-items: center; gap: 4px;">
                          <span style="font-weight: 600; font-size: 13px; color: var(--text-primary); cursor: pointer;" onclick="const btn = document.getElementById('toggle-btn-${job.id}-${idx}'); if (btn) btn.click();">${escapeHtml(step.title)}</span>
                          <button id="toggle-btn-${job.id}-${idx}" onclick="event.stopPropagation(); toggleStepDetails(this, 'step-details-${job.id}-${idx}')" style="background: transparent; border: none; color: var(--primary); font-weight: bold; cursor: pointer; font-size: 13px; padding: 0 4px; display: inline-flex; align-items: center; outline: none;">(+)</button>
                      </div>
                  `;
                  detailsWrapperHtml = `
                      <div id="step-details-${job.id}-${idx}" style="display: none; width: 100%;">
                          ${subTextHtml}
                          ${stepCardHtml}
                      </div>
                  `;
              } else {
                  titleWithToggleHtml = `
                      <span style="font-weight: 600; font-size: 13px; color: var(--text-primary);">${escapeHtml(step.title)}</span>
                  `;
              }

              taskLogRows += `
                  <div class="timeline-item" data-level="${level}" style="position: relative; padding-left: 36px; margin-bottom: 16px; display: flex; flex-direction: column;">
                      <div class="timeline-dot-wrapper" style="position: absolute; left: 6px; top: 4px; z-index: 2; width: 14px; height: 14px; display: flex; align-items: center; justify-content: center;">
                          <span style="width: 10px; height: 10px; border-radius: 50%; background: ${dotColor}; display: inline-block; border: 2px solid #ffffff; box-shadow: 0 0 0 1px rgba(27,31,35,0.15);"></span>
                      </div>
                      
                      <div class="timeline-content" style="display: flex; flex-direction: column;">
                          <div style="display: flex; align-items: center; gap: 8px; justify-content: space-between; width: 100%;">
                              ${titleWithToggleHtml}
                              <span style="color: var(--text-secondary); font-family: monospace; font-size: 11px; margin-left: auto; white-space: nowrap;">${timeRangeStr}</span>
                          </div>
                          ${detailsWrapperHtml}
                      </div>
                  </div>
              `;
              taskLogCount++;
          });
          
          let lineHtml = (taskLogRows) ? `<div style="position: absolute; left: 33px; top: 32px; bottom: 32px; width: 2px; background: #e1e4e8; z-index: 1;"></div>` : '';
          
          let taskLogsHtml = `
             <div style="flex: 1; min-width: 320px; display: flex; flex-direction: column;">
                 <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-wrap: wrap; gap: 12px;">
                    <span style="font-weight: 700; color: var(--text-primary); font-size: 15px;">
                       Registro de Atividades (Logs)
                    </span>
                    <div style="display: flex; align-items: center; gap: 8px;">
                       <input type="text" id="log-search-input" placeholder="Search logs" oninput="filterLogs()" style="padding: 5px 10px; font-size: 12px; border: 1px solid var(--border-color); border-radius: 6px; background: #ffffff; color: var(--text-primary); width: 140px; outline: none; height: 28px;">
                       <select id="log-level-select" onchange="filterLogs()" style="padding: 5px 10px; font-size: 12px; border: 1px solid var(--border-color); border-radius: 6px; background: #ffffff; color: var(--text-secondary); cursor: pointer; outline: none; height: 28px;">
                          <option value="ALL">Log Level</option>
                          <option value="INFO">Info</option>
                          <option value="WARN">Warning</option>
                          <option value="ERROR">Error</option>
                       </select>
                       <button class="btn" onclick="copyJobLog('${job.id}')" style="padding: 5px 10px; font-size: 12px; border: 1px solid var(--border-color); border-radius: 6px; background: #ffffff; color: var(--text-secondary); display: inline-flex; align-items: center; gap: 4px; cursor: pointer; font-weight: 500; height: 28px;">
                          📋 Copiar Logs
                       </button>
                    </div>
                 </div>
                 <div class="activity-log-scroll terminal-scroll" id="activity-log-scroll-${job.id}" style="height: 380px; overflow-y: auto; border: 1px solid var(--border-color); border-radius: 8px; background: #ffffff; display: flex; flex-direction: column; position: relative; padding: 24px 20px;">
                    ${lineHtml}
                    <div style="position: relative; z-index: 2; display: flex; flex-direction: column; gap: 16px;">
                        ${taskLogRows || '<div style="text-align: center; color: var(--text-secondary); font-size: 12px; font-style: italic; padding: 20px;">Nenhuma atividade registrada...</div>'}
                    </div>
                 </div>
             </div>
          `;
          // Build single-row horizontal stepper: step -> connector -> step -> connector
          let stepperRowHtml = '';
          for (let s = 0; s < numStages; s++) {
              const stageData = parseStageLabel(stageLabels[s], s);
              const cls = stepClasses[s];
              let iconContent = '';
              if (cls === 'completed') {
                  iconContent = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>`;
              } else if (cls === 'active') {
                  iconContent = `<span class="wizard-dot"></span>`;
              } else if (cls === 'failed') {
                  iconContent = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>`;
              }
              
              // Circle node
              let circleColorStyle = '';
              if (cls === 'completed') circleColorStyle = 'border-color: var(--success); color: var(--success); background: rgba(16, 185, 129, 0.08);';
              else if (cls === 'active') circleColorStyle = 'border-color: var(--primary); color: var(--primary); background: rgba(59, 130, 246, 0.08); box-shadow: 0 0 0 4px rgba(59, 130, 246, 0.12);';
              else if (cls === 'failed') circleColorStyle = 'border-color: var(--danger); color: var(--danger); background: rgba(239, 68, 68, 0.08);';
              else circleColorStyle = 'border-color: var(--border-color); color: var(--text-secondary); background: transparent;';

              // Connector line (after each step except last)
              let connectorHtml = '';
              if (s < numStages - 1) {
                  let connectorColor = 'var(--border-color)';
                  if (stepClasses[s] === 'completed') connectorColor = 'var(--success)';
                  else if (stepClasses[s] === 'active') connectorColor = 'var(--primary)';
                  connectorHtml = `<div style="flex: 1; height: 2px; background: ${connectorColor}; margin: 0 12px; transition: background 0.3s ease;"></div>`;
              }

              // Step markup (Circle + Labels side-by-side)
              stepperRowHtml += `
                  <div style="display: flex; align-items: center; gap: 10px; flex-shrink: 0;">
                      <div style="width: 24px; height: 24px; border-radius: 50%; border: 2px solid; display: flex; align-items: center; justify-content: center; flex-shrink: 0; transition: all 0.3s ease; ${circleColorStyle}">${iconContent}</div>
                      <div style="display: flex; flex-direction: column;">
                          <div style="font-size: 11px; font-weight: 700; color: var(--text-primary); white-space: nowrap;">${s+1}. ${escapeHtml(stageData.title)}</div>
                          <div style="font-size: 9px; color: var(--text-secondary); white-space: nowrap; margin-top: 1px;">${escapeHtml(stageData.sub)}</div>
                      </div>
                  </div>
                  ${connectorHtml}
              `;
          }
          let timelineHtml = `
             <!-- Active Job KPIs Row -->
             <div class="active-job-kpis" style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-top: 24px;">
                <!-- Card 1: Roles Concluídas -->
                <div class="job-kpi-card">
                   <div class="job-kpi-icon green">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                         <polyline points="20 6 9 17 4 12"></polyline>
                      </svg>
                   </div>
                   <div class="job-kpi-info">
                      <span class="job-kpi-value">${concludedRoles}/${totalRoles}</span>
                      <span class="job-kpi-title">${isCadeiaProcess ? 'Cadeias verificadas' : 'Roles concluídas'}</span>
                      <span class="job-kpi-sub">${totalRoles > 0 ? Math.round((concludedRoles / totalRoles) * 100) + '%' : '100%'}</span>
                   </div>
                </div>
                
                <!-- Card 2: Ações -->
                <div class="job-kpi-card">
                   <div class="job-kpi-icon purple">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                         <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon>
                      </svg>
                   </div>
                   <div class="job-kpi-info">
                      <span class="job-kpi-value">${totalActions}</span>
                      <span class="job-kpi-title">Ações</span>
                      <span class="job-kpi-sub">Total executadas</span>
                   </div>
                </div>
                
                <!-- Card 3: Tempo total -->
                <div class="job-kpi-card">
                   <div class="job-kpi-icon blue">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                         <circle cx="12" cy="12" r="10"></circle>
                         <polyline points="12 6 12 12 16 14"></polyline>
                      </svg>
                   </div>
                   <div class="job-kpi-info">
                      <span class="job-kpi-value">${formatDuration(job.created_at, (job.state === 'running' || job.state === 'pending') ? null : job.updated_at)}</span>
                      <span class="job-kpi-title">Tempo total</span>
                      <span class="job-kpi-sub">Duração da execução</span>
                   </div>
                </div>
                
                <!-- Card 4: Erros -->
                <div class="job-kpi-card">
                   <div class="job-kpi-icon red">
                      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                         <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path>
                         <line x1="12" y1="9" x2="12" y2="13"></line>
                         <line x1="12" y1="17" x2="12.01" y2="17"></line>
                      </svg>
                   </div>
                   <div class="job-kpi-info">
                      <span class="job-kpi-value">${errorCount}</span>
                      <span class="job-kpi-title">Erros</span>
                      <span class="job-kpi-sub">${errorCount > 0 ? 'Requer atenção' : 'Sem erros'}</span>
                   </div>
                </div>
             </div>
             
             <!-- Split Bottom Columns -->
             <div style="display: flex; gap: 24px; margin-top: 24px; flex-wrap: wrap;">
                ${roleHistoryHtml}
                ${taskLogsHtml}
             </div>
          `;
          let progressLabelsHtml = '';
          let progressWidth = '0%';
          if (totalRoles > 0) {
              const pct = Math.round((concludedRoles / totalRoles) * 100);
              progressLabelsHtml = `<span>Progresso (${concludedRoles}/${totalRoles} ${isCadeiaProcess ? 'Cadeias' : 'Roles'})</span><span>${pct}%</span>`;
              progressWidth = `${pct}%`;
          } else {
              if (job.state === 'running') {
                  progressLabelsHtml = `<span>Progresso (A iniciar...)</span><span>8%</span>`;
                  progressWidth = '8%';
              } else if (job.state === 'pending') {
                  progressLabelsHtml = `<span>Progresso (Pendente na fila)</span><span>0%</span>`;
                  progressWidth = '0%';
              } else if (job.state === 'failed') {
                  progressLabelsHtml = `<span>Progresso (${concludedRoles}/${totalRoles} ${isCadeiaProcess ? 'Cadeias' : 'Roles'})</span><span>100%</span>`;
                  progressWidth = '100%';
              } else if (job.state === 'succeeded') {
                  progressLabelsHtml = `<span>Progresso (Concluído)</span><span>100%</span>`;
                  progressWidth = '100%';
              }
          }

          let actionBtns = '';
          if (job.state === 'running' || job.state === 'pending') {
              actionBtns = `<button class="btn" style="padding: 4px 12px; font-size: 11px; font-weight: bold; background: rgba(239, 68, 68, 0.08); border: 1px solid rgba(239, 68, 68, 0.25); color: #EF4444; border-radius: 20px; cursor: pointer; transition: all 0.3s ease; display: inline-flex; align-items: center; gap: 4px; height: 26px;" onmouseover="this.style.background='rgba(239, 68, 68, 0.18)'; this.style.borderColor='rgba(239, 68, 68, 0.4)';" onmouseout="this.style.background='rgba(239, 68, 68, 0.08)'; this.style.borderColor='rgba(239, 68, 68, 0.25)';" onclick="event.stopPropagation(); cancelJob('${job.id}')">🛑 Cancelar</button>`;
          }

          let dropdownOptionsHtml = '';
          if (job.state === 'running' || job.state === 'pending') {
              dropdownOptionsHtml += `
                 <button class="dropdown-item" style="display: flex; align-items: center; gap: 8px; width: 100%; border: none; background: transparent; padding: 10px 16px; font-size: 13px; color: var(--danger); font-weight: 600; cursor: pointer; text-align: left; transition: background 0.2s; border-bottom: 1px solid rgba(0,0,0,0.05);" onmouseover="this.style.background='rgba(239, 68, 68, 0.08)'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); cancelJob('${job.id}')">
                    🛑 Cancelar Processamento
                 </button>
              `;
          } else {
              if (job.is_archived) {
                  dropdownOptionsHtml += `
                     <button class="dropdown-item" style="display: flex; align-items: center; gap: 8px; width: 100%; border: none; background: transparent; padding: 10px 16px; font-size: 13px; color: var(--text-secondary); cursor: pointer; text-align: left; transition: background 0.2s; border-bottom: 1px solid rgba(0,0,0,0.05);" onmouseover="this.style.background='rgba(0, 0, 0, 0.04)'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); unarchiveJob('${job.id}')">
                        📦 Desarquivar Job
                     </button>
                  `;
              } else {
                  dropdownOptionsHtml += `
                     <button class="dropdown-item" style="display: flex; align-items: center; gap: 8px; width: 100%; border: none; background: transparent; padding: 10px 16px; font-size: 13px; color: var(--text-secondary); cursor: pointer; text-align: left; transition: background 0.2s; border-bottom: 1px solid rgba(0,0,0,0.05);" onmouseover="this.style.background='rgba(0, 0, 0, 0.04)'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); archiveJob('${job.id}')">
                        📦 Arquivar Job
                     </button>
                  `;
              }
          }
          dropdownOptionsHtml += `
             <button class="dropdown-item" style="display: flex; align-items: center; gap: 8px; width: 100%; border: none; background: transparent; padding: 10px 16px; font-size: 13px; color: var(--text-secondary); cursor: pointer; text-align: left; transition: background 0.2s;" onmouseover="this.style.background='rgba(0, 0, 0, 0.04)'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); copyJobLog('${job.id}')">
                📋 Copiar Logs
             </button>
             <button class="dropdown-item" style="display: flex; align-items: center; gap: 8px; width: 100%; border: none; background: transparent; padding: 10px 16px; font-size: 13px; color: var(--danger); font-weight: 600; cursor: pointer; text-align: left; transition: background 0.2s;" onmouseover="this.style.background='rgba(239, 68, 68, 0.07)'" onmouseout="this.style.background='transparent'" onclick="event.stopPropagation(); deleteJob('${job.id}')">
               🗑️ Eliminar Job
             </button>
          `;

          const shortId = job.id.substring(0, 8);
          const isFocused = job.id === activeJobId;

          const borderStyle = isFocused 
              ? 'border: 2px solid var(--primary); box-shadow: 0 0 16px rgba(59, 130, 246, 0.2);' 
              : 'border: 1px solid var(--border-color);';

          const card = existingCard || document.createElement('div');
          card.className = 'card';
          card.id = `card-job-${job.id}`;
          card.style = `${borderStyle} transition: all 0.3s ease; cursor: pointer; margin-bottom: 24px; padding: 24px; position: relative;`;
          card.onclick = () => focusJob(job.id);
          card.setAttribute('data-updated-at', job.updated_at);
          card.setAttribute('data-state', job.state);
          // Focus indicator badge
          const focusIndicator = isFocused 
              ? `<span style="font-size: 9px; font-weight: 800; background: var(--primary); color: white; padding: 2px 6px; border-radius: 4px; text-transform: uppercase; letter-spacing: 0.05em; display: inline-flex; align-items: center; gap: 2px;">⚡ Focado</span>` 
              : `<span style="font-size: 9px; font-weight: 600; background: rgba(255,255,255,0.03); color: var(--text-secondary); padding: 2px 6px; border-radius: 4px; text-transform: uppercase; letter-spacing: 0.05em; display: inline-flex; align-items: center; gap: 2px;">Visualizar log</span>`;
          card.innerHTML = `
             <div style="display: flex; gap: 16px; align-items: center; margin-bottom: 20px; border-bottom: 1px solid var(--border-color); padding-bottom: 16px; position: relative;">
                <div style="width: 42px; height: 42px; border-radius: 50%; background: rgba(59, 130, 246, 0.08); display: flex; align-items: center; justify-content: center; color: var(--primary); flex-shrink: 0;">
                   <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                </div>
                <div style="flex: 1; min-width: 0;">
                   <div style="font-size: 11px; color: var(--text-secondary); display: flex; gap: 8px; flex-wrap: wrap; align-items: center;">
                      <span>📅 <strong>Data:</strong> ${data}</span>
                      <span style="opacity: 0.4">·</span>
                      <span>💻 <strong>Executado por:</strong> ${escapeHtml(job.worker_name || 'Sistema')}</span>
                      <span style="opacity: 0.4">·</span>
                      <span><strong>Sistema:</strong> ${escapeHtml(p.sap_system || 'S4D')} | <strong>Cliente:</strong> ${escapeHtml(p.sap_client || '100')}${p.sap_user ? ` | <strong>Utilizador SAP:</strong> ${escapeHtml(p.sap_user)}` : ''} | <strong>Processo:</strong> ${escapeHtml(sub || proc || job.task)} <span style="font-family: monospace; font-size: 11px;">#${shortId}</span></span>
                   </div>
                </div>
                
                <div style="display: flex; align-items: center; gap: 12px; flex-shrink: 0;">
                   ${focusIndicator}
                   <span class="badge-outline ${job.state}">${job.state.toUpperCase()}</span>
                   <div style="position: relative; display: inline-block;">
                      <button type="button" class="btn-icon" style="padding: 4px; color: var(--text-secondary);" onclick="event.stopPropagation(); toggleJobOptionsDropdown(event, '${job.id}')"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="1"></circle><circle cx="12" cy="5" r="1"></circle><circle cx="12" cy="19" r="1"></circle></svg></button>
                      <div id="job-options-menu-${job.id}" class="job-options-menu" style="display: none; position: absolute; right: 0; top: 30px; background: white; border: 1px solid var(--border-color); border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 100; min-width: 180px; padding: 6px 0;">
                         ${dropdownOptionsHtml}
                      </div>
                   </div>
                </div>
             </div>
             
             <!-- Connected Stages Timeline -->
             <div style="margin-bottom: 24px; background: rgba(0,0,0,0.01); border: 1px solid var(--border-color); border-radius: 12px; padding: 12px 20px;">
                 <div style="display: flex; align-items: center; width: 100%; justify-content: space-between;">
                     ${stepperRowHtml}
                 </div>
             </div>

             <div style="display: flex; gap: 20px; align-items: stretch; margin-bottom: 24px; flex-wrap: wrap;">
                <!-- Progress Bar -->
                <div class="card-inner" style="flex: 1; min-width: 250px; background: rgba(255,255,255,0.02); border: 1px solid rgba(255,255,255,0.05); padding: 16px; border-radius: 12px;">
                   <div style="display: flex; justify-content: space-between; font-size: 12px; color: var(--text-secondary); margin-bottom: 8px; font-weight: 600;">
                      ${progressLabelsHtml}
                   </div>
                   <div class="progress-bar-bg" style="height: 6px; background: #e5e7eb; border-radius: 3px; overflow: hidden; margin-bottom: 8px;">
                      <div class="progress-bar-fill" style="width: ${progressWidth}; transition: width 0.5s ease; height: 100%; background: var(--primary);"></div>
                   </div>
                   ${currentRoleName && job.state === 'running' ? `<div style="font-size:11px; color:var(--primary); font-weight:bold; animation: pulse-opacity 1.5s infinite;">A processar: ${escapeHtml(currentRoleName)}</div>` : ''}
                </div>
             </div>
             <div id="details-${job.id}">
                ${timelineHtml}
             </div>
          `;
          if (!existingCard) {
              if (container.children[idx]) {
                  container.insertBefore(card, container.children[idx]);
              } else {
                  container.appendChild(card);
              }
          }
          // Auto-scroll the role history lists to keep the active or last completed role visible inside this card
          const scrollContainer = card.querySelector('.role-history-scroll');
          if (scrollContainer) {
              const activeRow = scrollContainer.querySelector('.role-history-row.active-role');
              if (activeRow) {
                  const containerHeight = scrollContainer.clientHeight;
                  const rowTop = activeRow.offsetTop;
                  const rowHeight = activeRow.offsetHeight;
                  scrollContainer.scrollTop = rowTop - (containerHeight / 2) + (rowHeight / 2);
              } else {
                  const completedRows = scrollContainer.querySelectorAll('.role-history-row');
                  let lastCompletedRow = null;
                  for (let i = completedRows.length - 1; i >= 0; i--) {
                      if (completedRows[i].innerHTML.includes('Concluída')) {
                          lastCompletedRow = completedRows[i];
                          break;
                      }
                  }
                  if (lastCompletedRow) {
                      const containerHeight = scrollContainer.clientHeight;
                      const rowTop = lastCompletedRow.offsetTop;
                      const rowHeight = lastCompletedRow.offsetHeight;
                      scrollContainer.scrollTop = rowTop - (containerHeight / 2) + (rowHeight / 2);
                  }
              }
          }
          // Restaurar ou atualizar o scroll do registo de atividades
          const actScrollContainer = card.querySelector('.activity-log-scroll');
          if (actScrollContainer) {
                      const saved = savedScrolls[actScrollContainer.id];
              const isRunning = job.state === 'running' || job.state === 'pending';
              if (saved) {
                  if (isRunning) {
                      if (saved.isAtBottom) {
                          actScrollContainer.scrollTop = actScrollContainer.scrollHeight;
                      } else {
                          actScrollContainer.scrollTop = saved.scrollTop;
                      }
                  } else {
                      actScrollContainer.scrollTop = saved.scrollTop;
                  }
              } else {
                  actScrollContainer.scrollTop = actScrollContainer.scrollHeight;
              }
          }
       });

       // Restore input value and trigger filter for all cards
       const queryInput = document.getElementById('log-search-input');
       const stateSelect = document.getElementById('log-state-select');
       const timeSelect = document.getElementById('log-time-select');
       if (queryInput) {
           queryInput.value = activeLogQuery;
       }
       if (stateSelect) {
           stateSelect.value = activeLogState;
       }
       if (timeSelect) {
           timeSelect.value = activeLogTime;
       }
       filterLogs();

       // Render Focused Job Logs in the sidebar
       const focusedJob = allJobs.find(j => j.id === activeJobId);
       if (focusedJob && logContainer) {
           const lastJobId = logContainer.getAttribute('data-job-id');
           const lastUpdated = logContainer.getAttribute('data-updated-at');
           const lastState = logContainer.getAttribute('data-state');
           
           if (lastJobId !== focusedJob.id || lastUpdated !== focusedJob.updated_at || lastState !== focusedJob.state) {
               logContainer.setAttribute('data-job-id', focusedJob.id);
               logContainer.setAttribute('data-updated-at', focusedJob.updated_at);
               logContainer.setAttribute('data-state', focusedJob.state);
               const rawLines = (focusedJob.log || '')
                  .replace(/\\r\\n/g, '\n')
                  .replace(/\\n/g, '\n')
                  .split('\n')
                  .filter(l => l.trim() !== '');
               let rawLog = rawLines.map((line, idx) => {
                   let tr = line.trim();
                   let cls = 'info';
                   if (tr.startsWith('[OK]')) cls = 'ok';
                   if (tr.startsWith('[ERRO]') || tr.includes('Erro') || tr.includes('Falha') || tr.includes('Exception') || tr.includes('Traceback')) cls = 'erro';
                   if (tr.startsWith('[WARN]')) cls = 'warn';
                   
                   const ts = getTimestampForLine(focusedJob, idx, rawLines.length);
                   
                   return `
                      <div class="log-row" style="display: flex; gap: 12px; margin-bottom: 4px; font-family: Consolas, monospace; font-size: 12px; line-height: 1.4;">
                         <span style="color: #58a6ff; flex-shrink: 0; user-select: none; width: 60px;">${ts}</span>
                         <span style="color: rgba(255,255,255,0.15); user-select: none;">|</span>
                         <span class="log-line ${cls}" style="white-space: pre-wrap; flex: 1;">${escapeHtml(tr)}</span>
                      </div>
                    `;
               }).join('');
               if (focusedJob.state === 'running' || focusedJob.state === 'pending') {
                   const dummyTime = formatTimeOnly(new Date());
                   rawLog += `
                      <div class="log-row" style="display: flex; gap: 12px; margin-bottom: 4px; font-family: Consolas, monospace; font-size: 12px; line-height: 1.4;">
                         <span style="color: #58a6ff; flex-shrink: 0; user-select: none; width: 60px;">${dummyTime}</span>
                         <span style="color: rgba(255,255,255,0.15); user-select: none;">|</span>
                         <span class="log-line info" style="white-space: pre-wrap; flex: 1;">> a processar<span class="cursor-blink" style="color:white; display:inline-block; width:8px; height:12px; background:white; margin-left:4px; animation: blink 1s step-end infinite;"></span></span>
                      </div>
                    `;
               }
               logContainer.innerHTML = rawLog || 'Nenhum log disponível para o job focado.';
               const statusEl = document.getElementById('log-connection-status');
               const timeEl = document.getElementById('log-updated-time');
               if (statusEl) {
                   if (focusedJob.state === 'running') {
                       statusEl.textContent = 'A processar...';
                   } else if (focusedJob.state === 'pending') {
                       statusEl.textContent = 'Pendente...';
                   } else {
                       statusEl.textContent = 'Conectado';
                   }
               }
               if (timeEl) {
                   timeEl.textContent = `Atualizado: ${formatTimeOnly(new Date())}`;
               }
               // Restaurar ou atualizar o scroll do log em tempo real (sidebar)
               const isRunning = focusedJob.state === 'running' || focusedJob.state === 'pending';
               if (savedRealtimeScroll !== null) {
                   if (isRunning) {
                       if (realtimeIsAtBottom) {
                           logContainer.scrollTop = logContainer.scrollHeight;
                       } else {
                           logContainer.scrollTop = savedRealtimeScroll;
                       }
                   } else {
                       logContainer.scrollTop = savedRealtimeScroll;
                   }
               } else {
                   logContainer.scrollTop = logContainer.scrollHeight;
               }
           }
           // Copy log action
           document.getElementById('btn-copy-log').onclick = () => {
              navigator.clipboard.writeText(focusedJob.log || '');
              const old = document.getElementById('btn-copy-log').textContent;
              document.getElementById('btn-copy-log').textContent = 'Copiado!';
              setTimeout(() => document.getElementById('btn-copy-log').textContent = old, 2000);
           };
       } else if (logContainer) {
          logContainer.innerHTML = 'Nenhum job em foco para visualização de logs.';
       }
    }

    function filterLogs() {
        const queryInput = document.getElementById('log-search-input');
        const stateSelect = document.getElementById('log-state-select');
        const timeSelect = document.getElementById('log-time-select');
        
        activeLogQuery = queryInput ? queryInput.value.toLowerCase() : '';
        activeLogState = stateSelect ? stateSelect.value : 'ALL';
        activeLogTime = timeSelect ? timeSelect.value : 'ALL';
        
        const now = Date.now();
        const oneDayMs = 24 * 60 * 60 * 1000;
        
        const items = document.querySelectorAll('.timeline-item');
        items.forEach(item => {
            const text = item.textContent.toLowerCase();
            const state = item.getAttribute('data-state') || 'completed';
            const timestamp = parseInt(item.getAttribute('data-time') || '0');
            
            const matchQuery = text.includes(activeLogQuery);
            const matchState = (activeLogState === 'ALL') || (state === activeLogState);
            
            let matchTime = true;
            if (activeLogTime === 'TODAY') {
                const todayStart = new Date().setHours(0,0,0,0);
                matchTime = timestamp >= todayStart;
            } else if (activeLogTime === '24H') {
                matchTime = (now - timestamp) <= oneDayMs;
            } else if (activeLogTime === '7D') {
                matchTime = (now - timestamp) <= (7 * oneDayMs);
            } else if (activeLogTime === '30D') {
                matchTime = (now - timestamp) <= (30 * oneDayMs);
            }
            
            if (matchQuery && matchState && matchTime) {
                item.style.display = 'flex';
            } else {
                item.style.display = 'none';
            }
        });
    }

    function toggleLogItemDetails(itemId) {
        expandedLogItems[itemId] = !expandedLogItems[itemId];
        const container = document.getElementById(`log-item-details-${itemId}`);
        const icon = document.getElementById(`chevron-icon-${itemId}`);
        if (container) {
            container.style.display = expandedLogItems[itemId] ? 'block' : 'none';
        }
        if (icon) {
            icon.style.transform = expandedLogItems[itemId] ? 'rotate(180deg)' : '';
        }
    }

    function handleLogExport(el, jobId) {
        const val = el.value;
        if (!val) return;
        
        el.value = ''; // Reset select
        
        const job = allJobs.find(j => j.id === jobId);
        if (!job) return;
        
        if (val === 'COPY') {
            navigator.clipboard.writeText(job.log || '');
            alert('Logs copiados para a área de transferência!');
        } else if (val === 'JSON') {
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(job, null, 2));
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href",     dataStr);
            downloadAnchor.setAttribute("download", `job_${job.id}_logs.json`);
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
        } else if (val === 'TXT') {
            const dataStr = "data:text/plain;charset=utf-8," + encodeURIComponent(job.log || '');
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href",     dataStr);
            downloadAnchor.setAttribute("download", `job_${job.id}_logs.txt`);
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
        }
    }

    function toggleStepDetails(btn, detailsId) {
        const detailsDiv = document.getElementById(detailsId);
        if (!detailsDiv) return;
        if (detailsDiv.style.display === 'none') {
            detailsDiv.style.display = 'block';
            btn.textContent = '(-)';
        } else {
            detailsDiv.style.display = 'none';
            btn.textContent = '(+)';
        }
    }

    function focusJob(jobId) {
        activeJobId = jobId;
        renderQueue();
    }



    async function cancelJob(jobId) {
        try {
            const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/cancel`, { method: 'POST' });
            if (res.ok) {
                await loadJobs();
            }
        } catch(e) {
            console.error('Erro de rede ao cancelar:', e);
        }
    }

    async function archiveJob(jobId) {
        try {
            const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/archive`, { method: 'POST' });
            if (res.ok) {
                await loadJobs();
            }
        } catch(e) {
            console.error('Erro de rede ao arquivar:', e);
        }
    }

    async function unarchiveJob(jobId) {
        try {
            const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}/unarchive`, { method: 'POST' });
            if (res.ok) {
                // Remove from temporary allJobs if it was temporary, but let loadJobs handle it
                await loadJobs();
            }
        } catch(e) {
            console.error('Erro de rede ao desarquivar:', e);
        }
    }

    async function deleteJob(jobId) {
        if (!confirm('Tem a certeza que quer eliminar este job permanentemente? Esta ação não pode ser desfeita.')) {
            return;
        }
        try {
            const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`, { method: 'DELETE' });
            if (res.ok) {
                const card = document.getElementById(`card-job-${jobId}`);
                if (card) card.remove();
                if (typeof activeJobId !== 'undefined' && activeJobId === jobId) {
                    activeJobId = null;
                }
                await loadJobs();
            } else {
                const data = await res.json().catch(() => ({}));
                alert('Erro ao eliminar o job: ' + (data.detail || res.status));
            }
        } catch(e) {
            console.error('Erro de rede ao eliminar:', e);
            alert('Erro de rede ao eliminar o job.');
        }
    }

    function copyJobLog(jobId) {
        const job = allJobs.find(j => j.id === jobId);
        if (job) {
            navigator.clipboard.writeText(job.log || '');
            alert('Logs copiados com sucesso!');
        }
    }

    function toggleJobOptionsDropdown(event, jobId) {
        event.stopPropagation();
        
        // Close any other open menus
        const allMenus = document.querySelectorAll('.job-options-menu');
        allMenus.forEach(menu => {
            if (menu.id !== `job-options-menu-${jobId}`) {
                menu.style.display = 'none';
            }
        });

        const menu = document.getElementById(`job-options-menu-${jobId}`);
        if (menu) {
            const isHidden = menu.style.display === 'none' || menu.style.display === '';
            menu.style.display = isHidden ? 'block' : 'none';
        }
    }


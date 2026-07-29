const AUTH_CHAT_STATES = {
      LOADING: 'loading',
      WAITING_USER: 'waiting_user',
      WAITING_SYSTEM: 'waiting_system',
      WAITING_ANALYSIS_TYPE: 'waiting_analysis_type',
      WAITING_INDIVIDUAL_USER: 'waiting_individual_user',
      WAITING_INDIVIDUAL_SYSTEM: 'waiting_individual_system',
      WAITING_INDIVIDUAL_PARAMS: 'waiting_individual_params',
      READY: 'ready',
      OPENING_SAP: 'opening_sap',
      SAP_READY: 'sap_ready',
      VALIDATING: 'validating',
      ANALYZING: 'analyzing',
      ANALYSIS_COMPLETE: 'analysis_complete',
      ERROR: 'error'
    };

    let authorizationChatState = AUTH_CHAT_STATES.LOADING;
    let authorizationTargetUser = '';
    let authorizationSelectedSystem = null;
    let authorizationAvailableSystems = [];
    let authorizationTechnicalUser = '';
    let authChatFetchController = null;
    let authorizationChatRequestId = 0;
    let authorizationSelectedAnalysisType = null;
    let authorizationLastStatusData = null;
    let authorizationLastDisplayedRoles = [];
    let authorizationPendingRemoval = null;
    let authorizationRemovalLastContext = null;
    let authorizationIndividualContext = null;
    let authorizationAnalysisTypes = [];
    let authorizationActiveJobId = null;
    let authorizationJobRequestId = 0;
    let authorizationRemovalJobRequestId = 0;

    let authorizationLoadRequestId = 0;
    let authorizationChatLoading = false;
    let authorizationChatInitialized = false;
    let authorizationLoadPromise = null;
    let authorizationLoadingWatchdog = null;
    const AUTHORIZATION_CONFIG_ENDPOINT = '/api/authorizations/config';

    function getAuthorizationExecutionMode() {
      const systemKey = String(authorizationSelectedSystem?.key || '').trim().toUpperCase();
      const explicitMode = String(authorizationSelectedSystem?.execution_mode || '').trim().toUpperCase();
      if (systemKey.startsWith('S4')) {
        return 'RFC';
      }
      if (systemKey.startsWith('SPA')) {
        return 'CUA';
      }
      if (explicitMode === 'RFC' || explicitMode === 'CUA') {
        return explicitMode;
      }
      return '';
    }

    function isAuthorizationDevFlow() {
      const systemValue = String(authorizationSelectedSystem?.system || '').trim().toUpperCase();
      const keyValue = String(authorizationSelectedSystem?.key || '').trim().toUpperCase();
      return systemValue === 'DEV' || keyValue.startsWith('S4');
    }

    async function fetchWithTimeout(url, options = {}, timeoutMs = 10000) {
      const controller = new AbortController();
      const timeoutId = window.setTimeout(() => controller.abort(), timeoutMs);
      try {
        return await fetch(url, {
          ...options,
          signal: controller.signal,
          cache: 'no-store'
        });
      } finally {
        window.clearTimeout(timeoutId);
      }
    }

    function escapeAuthorizationText(text) {
      if (!text) return '';
      return text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }

    function updateAuthorizationStatus(stateKey, customText) {
      const badge = document.getElementById('auth-chat-status') || document.querySelector('.auth-chat-status');
      if (!badge) return;

      badge.className = 'auth-chat-status';

      if (stateKey === 'ready' || stateKey === 'connected') {
        badge.classList.add('connected');
        badge.innerHTML = `<span class="auth-chat-status-dot"></span> ${customText || ('Sessão técnica: ' + (authorizationTechnicalUser || 'SAP'))}`;
      } else if (stateKey === 'loading' || stateKey === 'connecting') {
        badge.classList.add('connecting');
        badge.innerHTML = `<span class="auth-chat-status-dot"></span> ${customText || 'A carregar...'}`;
      } else if (stateKey === 'error') {
        badge.classList.add('error');
        badge.innerHTML = `<span class="auth-chat-status-dot"></span> ${customText || 'Erro na ligação'}`;
      } else {
        badge.innerHTML = `<span class="auth-chat-status-dot"></span> ${customText || stateKey}`;
      }
    }

    function updateAuthorizationStatusBadge() {
      if (authorizationChatState === AUTH_CHAT_STATES.LOADING) {
        updateAuthorizationStatus('loading');
      } else if (authorizationChatState === AUTH_CHAT_STATES.ERROR) {
        updateAuthorizationStatus('error');
      } else {
        updateAuthorizationStatus('ready');
      }
    }

    function updateAuthorizationComposer() {
      const input = document.getElementById('authorization-chat-input');
      const button = document.getElementById('authorization-chat-send');

      if (!input || !button) {
        return;
      }

      const waitingForUser =
        authorizationChatState === AUTH_CHAT_STATES.WAITING_USER ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS;
      const followUpReady =
        authorizationChatState === AUTH_CHAT_STATES.READY ||
        authorizationChatState === AUTH_CHAT_STATES.ANALYSIS_COMPLETE;
      const executionMode = getAuthorizationExecutionMode();

      input.disabled = !(waitingForUser || followUpReady);

      if (authorizationChatState === AUTH_CHAT_STATES.LOADING) {
        input.placeholder = 'A carregar configuração SAP...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_USER) {
        input.placeholder = 'Escreva a sua mensagem ou utilizador SAP...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER) {
        input.placeholder = 'Escreva o utilizador SAP alvo (ex: CSILVA)...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_SYSTEM || authorizationChatState === AUTH_CHAT_STATES.WAITING_SYSTEM) {
        input.placeholder = 'Selecione um sistema/ambiente acima...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS) {
        input.placeholder = 'Escreva os detalhes/parâmetros do pedido...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_ANALYSIS_TYPE) {
        input.placeholder = 'Selecione o tipo de análise...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.OPENING_SAP) {
        input.placeholder = executionMode === 'RFC' ? 'A abrir ligação RFC...' : 'A abrir sessão SAP...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.SAP_READY) {
        input.placeholder = executionMode === 'RFC' ? 'Ligação RFC pronta.' : 'Sessão SAP pronta.';
      } else if (authorizationChatState === AUTH_CHAT_STATES.ANALYZING) {
        input.placeholder = executionMode === 'RFC' ? 'A analisar autorizações via RFC...' : 'A analisar autorizações...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.ANALYSIS_COMPLETE) {
        input.placeholder = 'Pergunte sobre a lista ou diga "remova estas funções"...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.READY) {
        input.placeholder = 'Pergunte sobre a lista ou diga "remova estas funções"...';
      } else {
        input.placeholder = 'Aguarde...';
      }


      button.disabled =
        !(waitingForUser || followUpReady) ||
        input.value.trim() === '';
    }

    function setAuthorizationChatState(newState) {
      authorizationChatState = newState;
      updateAuthorizationStatusBadge();
      updateAuthorizationComposer();
    }

    function renderAuthorizationLoadingState() {
      updateAuthorizationStatus('loading');
      updateAuthorizationComposer();
    }

    function renderAuthorizationFollowUpResult(title, items, emptyMessage) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      authorizationLastDisplayedRoles = Array.isArray(items) ? items.slice() : [];

      const wrapper = document.createElement('div');
      wrapper.className = 'auth-chat-summary';

      const header = document.createElement('div');
      header.className = 'auth-chat-summary-row';
      header.style.fontWeight = '700';
      header.style.paddingBottom = '10px';
      header.innerHTML = `<span class="auth-chat-summary-label">${escapeAuthorizationText(title)}</span><span class="auth-chat-summary-value">${String(items.length)}</span>`;
      wrapper.appendChild(header);

      if (!items.length) {
        const empty = document.createElement('div');
        empty.style.fontSize = '0.85rem';
        empty.style.color = 'var(--text-secondary)';
        empty.textContent = emptyMessage;
        wrapper.appendChild(empty);
      } else {
        const list = document.createElement('div');
        list.style.display = 'grid';
        list.style.gap = '8px';

        items.forEach(item => {
          const row = document.createElement('div');
          row.style.display = 'flex';
          row.style.justifyContent = 'space-between';
          row.style.gap = '12px';
          row.style.padding = '8px 10px';
          row.style.border = '1px solid var(--border-color)';
          row.style.borderRadius = '10px';
          row.style.background = 'rgba(0,0,0,0.02)';

          const left = document.createElement('span');
          left.style.fontWeight = '700';
          left.textContent = item.role || item.function || '';

          const right = document.createElement('span');
          right.style.color = 'var(--text-secondary)';
          right.textContent = [item.valid_from, item.valid_to, item.assignment_origin_label || item.assignment_origin].filter(Boolean).join(' · ');

          row.appendChild(left);
          row.appendChild(right);
          list.appendChild(row);
        });

        wrapper.appendChild(list);
      }

      container.appendChild(wrapper);
      container.scrollTop = container.scrollHeight;
    }

    function renderAuthorizationRemovalPrompt(roles, label) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const list = Array.isArray(roles) ? roles.slice() : [];
      authorizationPendingRemoval = {
        roles: list,
        label: label || 'as funções selecionadas',
        targetUser: authorizationTargetUser,
        targetSystemKey: authorizationSelectedSystem?.key || '',
        systemShort: authorizationSelectedSystem?.system || ''
      };

      const wrapper = document.createElement('div');
      wrapper.className = 'auth-chat-summary';

      const header = document.createElement('div');
      header.className = 'auth-chat-summary-row';
      header.style.fontWeight = '700';
      header.style.paddingBottom = '10px';
      header.innerHTML = `<span class="auth-chat-summary-label">Remover funções</span><span class="auth-chat-summary-value">${String(list.length)}</span>`;
      wrapper.appendChild(header);

      const text = document.createElement('div');
      text.style.marginBottom = '10px';
      text.textContent = `Queres que eu crie a remoção no CUA para ${label || 'as funções selecionadas'} do utilizador ${authorizationTargetUser}?`;
      wrapper.appendChild(text);

      const items = document.createElement('div');
      items.style.display = 'grid';
      items.style.gap = '8px';

      list.forEach(item => {
        const row = document.createElement('div');
        row.style.display = 'flex';
        row.style.justifyContent = 'space-between';
        row.style.gap = '12px';
        row.style.padding = '8px 10px';
        row.style.border = '1px solid var(--border-color)';
        row.style.borderRadius = '10px';
        row.style.background = 'rgba(0,0,0,0.02)';

        const left = document.createElement('span');
        left.style.fontWeight = '700';
        left.textContent = item.role || item.function || '';

        const right = document.createElement('span');
        right.style.color = 'var(--text-secondary)';
        right.textContent = [item.valid_from, item.valid_to, item.assignment_origin_label || item.assignment_origin].filter(Boolean).join(' · ');

        row.appendChild(left);
        row.appendChild(right);
        items.appendChild(row);
      });

      wrapper.appendChild(items);

      const actions = document.createElement('div');
      actions.className = 'auth-chat-summary-actions';
      actions.style.marginTop = '10px';

      actions.innerHTML = `
        <button type="button" class="btn btn-secondary btn-sm" onclick="cancelAuthorizationRemoval()">Cancelar</button>
        <button type="button" class="btn btn-primary btn-sm" onclick="confirmAuthorizationRemoval()">Criar remoção</button>
      `;
      wrapper.appendChild(actions);

      container.appendChild(wrapper);
      container.scrollTop = container.scrollHeight;
    }

    function isAuthorizationTerminalJobState(state) {
      const normalized = String(state || '').trim().toLowerCase();
      return ['succeeded', 'succeeded_with_warnings', 'failed', 'cancelled', 'canceled', 'stopped', 'error'].includes(normalized);
    }

    function parseAuthorizationRemovalSummary(job) {
      const logText = String(job?.log || '');
      const statusText = String(job?.status || '');
      const sourceText = `${logText}\n${statusText}`;
      const getMatchNumber = (pattern) => {
        const match = sourceText.match(pattern);
        return match ? Number.parseInt(match[1], 10) : null;
      };

      const systemMatch = sourceText.match(/Sistema principal do pedido:\s*([^\n\r]+)/i)
        || sourceText.match(/SISTEMA='([^']+)'/i)
        || sourceText.match(/SISTEMA=([A-Z0-9_]+)/i);
      const userMatch = sourceText.match(/UTILIZADOR='([^']+)'/i)
        || sourceText.match(/UTILIZADOR=([A-Z0-9.\-_]+)/i);

      return {
        user: userMatch ? String(userMatch[1] || '').trim() : '',
        system: systemMatch ? String(systemMatch[1] || '').trim() : '',
        processed: getMatchNumber(/Linhas processadas:\s*(\d+)/i),
        concluded: getMatchNumber(/Conclu[ií]das:\s*(\d+)/i),
        warnings: getMatchNumber(/Avisos:\s*(\d+)/i),
        errors: getMatchNumber(/Erros:\s*(\d+)/i),
        removed: getMatchNumber(/Fun[cç][oõ]es eliminadas:\s*(\d+)/i),
        noMatches: /Nenhuma fun[cç][aã]o encontrada/i.test(sourceText)
      };
    }

    function buildAuthorizationRemovalFeedback(job) {
      const summary = parseAuthorizationRemovalSummary(job);
      const fallbackContext = authorizationRemovalLastContext || {};
      const state = String(job?.state || '').trim().toLowerCase();
      const stateLabel = state === 'succeeded_with_warnings' ? 'concluida com avisos' : state === 'succeeded' ? 'concluida com sucesso' : state || 'terminada';
      const jobIdShort = String(job?.id || '').slice(0, 8);
      const pluralize = (value, singular, plural) => `${value} ${value === 1 ? singular : plural}`;
      const counts = [];

      if (summary.processed !== null) counts.push(pluralize(summary.processed, 'linha processada', 'linhas processadas'));
      if (summary.concluded !== null) counts.push(pluralize(summary.concluded, 'concluida', 'concluidas'));
      if (summary.warnings !== null) counts.push(pluralize(summary.warnings, 'aviso', 'avisos'));
      if (summary.errors !== null) counts.push(pluralize(summary.errors, 'erro', 'erros'));
      if (summary.removed !== null) counts.push(pluralize(summary.removed, 'funcao eliminada', 'funcoes eliminadas'));

      const subjectUser = summary.user || String(fallbackContext.user || '').trim();
      const subjectSystem = summary.system || String(fallbackContext.system || '').trim();
      const subject = subjectUser && subjectSystem
        ? `${subjectUser} no sistema ${subjectSystem}`
        : subjectUser
          ? `o utilizador ${subjectUser}`
          : 'a remocao';

      let message = `Remocao ${stateLabel} para ${subject}`;
      if (jobIdShort) {
        message += ` (job #${jobIdShort})`;
      }
      if (counts.length > 0) {
        message += `: ${counts.join(', ')}.`;
      }
      if (summary.noMatches) {
        message += ' O log indica que nao foram encontradas funcoes no sistema alvo.';
      } else if (summary.removed === 0) {
        message += ' Sem funcoes eliminadas.';
      }
      return message;
    }

    function normalizeAuthorizationToken(value) {
      return String(value || '').trim().toUpperCase();
    }

    function normalizeAuthorizationSearchText(value) {
      return String(value || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase();
    }

    function extractAuthorizationTransactionCode(query) {
      const upperQuery = normalizeAuthorizationSearchText(query);
      const candidates = upperQuery.match(/\b[A-Z0-9]{3,5}\b/g) || [];
      const stopWords = new Set([
        'UM', 'UMA', 'UNS', 'UMAS', 'ESTE', 'ESTA', 'ESSA', 'ESSE', 'ESTES', 'ESTAS', 'ESSES', 'ESSAS',
        'QUERO', 'PESQUISAR', 'PESQUISA', 'SOBRE', 'PROCESSO', 'PROCESSOS',
        'SKILL', 'SKILLS', 'PODES', 'FAZER', 'COMO', 'ONDE', 'QUAL', 'QUAIS',
        'NOVO', 'NOVA', 'NOME', 'OBTER', 'MOSTRAR', 'DAR', 'DAME', 'DÁ-ME',
        'VER', 'SABER', 'QUERIA', 'GOSTARIA', 'AJUDA', 'AJUDAR', 'EXCLUSIVAS',
        'USER', 'USERS', 'UTILIZADOR', 'UTILIZADORES', 'CONTA', 'ID', 'SISTEMA', 'SISTEMAS',
        'ANALISAR', 'ATRIBUIDO', 'ATRIBUIDA', 'ATRIBUIDOS', 'ATRIBUIDAS',
        'ROLE', 'ROLES', 'FUNCAO', 'FUNCOES', 'FUNÇÃO', 'FUNÇÕES', 'PERFIL', 'PERFIS',
        'TRANSA', 'TRANSAO', 'TRANSAÇÕES', 'TRANSACOES', 'TRANACOES', 'LISTA', 'LISTAR',
        'QUAISAS', 'MINHAS', 'MEUS', 'MINHA', 'TEM', 'COM', 'PARA',
        'EXPIRADA', 'EXPIRADAS', 'EXPIRADO', 'EXPIRADOS',
        'ATIVA', 'ATIVAS', 'ATIVO', 'ATIVOS',
        'DIRETA', 'DIRETAS', 'DIRETO', 'DIRETOS',
        'INDIRETA', 'INDIRETAS', 'INDIRETO', 'INDIRETOS'
      ]);
      return candidates.find((code) => !stopWords.has(code)) || '';
    }

    function extractAuthorizationRoleFromQuery(query, roles) {
      const queryUpper = normalizeAuthorizationSearchText(query);
      const roleList = Array.isArray(roles) ? roles : [];

      const directMatch = roleList.find((role) => {
        const roleName = normalizeAuthorizationToken(role?.role || role?.name || role?.function || '');
        return roleName && queryUpper.includes(roleName);
      });
      if (directMatch) {
        return String(directMatch.role || directMatch.name || directMatch.function || '').trim();
      }

      const roleMatch = normalizeAuthorizationSearchText(query).match(/(?:ROLE|FUNCAO|FUNCOES)\s+([A-Z0-9._-]+)/i);
      if (roleMatch && roleMatch[1]) {
        return String(roleMatch[1]).trim();
      }

      return '';
    }

    function getAuthorizationRoleFunctions(data, roleName) {
      const targetRole = normalizeAuthorizationToken(roleName);
      if (!targetRole) {
        return [];
      }

      const roleFunctionsMap = data && typeof data.role_functions === 'object' && data.role_functions ? data.role_functions : {};
      const mapEntry = roleFunctionsMap[targetRole] || roleFunctionsMap[String(roleName).trim()] || [];
      const roleEntry = Array.isArray(data?.roles)
        ? data.roles.find((role) => normalizeAuthorizationToken(role?.role || role?.name || role?.function || '') === targetRole)
        : null;
      const rawList = Array.isArray(mapEntry) && mapEntry.length > 0
        ? mapEntry
        : Array.isArray(roleEntry?.functions)
          ? roleEntry.functions
          : [];

      const seen = new Set();
      const result = [];
      rawList.forEach((item) => {
        const tcode = String(item || '').trim().toUpperCase();
        if (!tcode || seen.has(tcode)) {
          return;
        }
        seen.add(tcode);
        result.push(tcode);
      });
      return result;
    }

    function getAuthorizationRolesWithTransaction(data, tcode) {
      const targetTcode = normalizeAuthorizationToken(tcode);
      if (!targetTcode) {
        return [];
      }

      const roleList = Array.isArray(data?.roles) ? data.roles : [];
      const roleFunctionsMap = data && typeof data.role_functions === 'object' && data.role_functions ? data.role_functions : {};
      const seen = new Set();
      const result = [];

      roleList.forEach((role) => {
        const roleName = String(role?.role || role?.name || role?.function || '').trim();
        if (!roleName) {
          return;
        }

        const roleKey = normalizeAuthorizationToken(roleName);
        const roleFunctions = Array.isArray(role?.functions) && role.functions.length > 0
          ? role.functions
          : Array.isArray(roleFunctionsMap[roleKey]) ? roleFunctionsMap[roleKey] : [];

        const hasTransaction = roleFunctions.some((item) => normalizeAuthorizationToken(item) === targetTcode);
        if (!hasTransaction || seen.has(roleKey)) {
          return;
        }

        seen.add(roleKey);
        result.push(roleName);
      });

      return result;
    }

    function getAuthorizationAllFunctions(data) {
      const seen = new Set();
      const result = [];
      const functionsList = Array.isArray(data?.functions) ? data.functions : [];
      functionsList.forEach((item) => {
        const tcode = String(item?.tcode || item?.function || item || '').trim().toUpperCase();
        if (!tcode || seen.has(tcode)) {
          return;
        }
        seen.add(tcode);
        result.push(tcode);
      });
      return result;
    }

    function renderAuthorizationPlainList(title, items, emptyMessage) {
      if (!Array.isArray(items) || items.length === 0) {
        appendAuthorizationMessage('assistant', emptyMessage);
        return;
      }

      const html = `
        <div class="auth-chat-summary" style="width:100%;">
          <div class="auth-chat-summary-row" style="font-weight:700; margin-bottom:12px; border-bottom:1px solid var(--border-color); padding-bottom:8px;">
            <span class="auth-chat-summary-label" style="font-size:0.9rem; color:var(--text-primary);">${escapeAuthorizationText(title)}</span>
            <span class="auth-chat-summary-value" style="background:#2563eb; color:white; padding:2px 10px; border-radius:12px; font-size:0.78rem;">${String(items.length)}</span>
          </div>
          <div style="display:flex; flex-wrap:wrap; gap:8px; width:100%; margin-top:4px;">
            ${items.map((item) => `
              <div style="display:inline-flex; align-items:center; justify-content:center; padding:6px 14px; border:1px solid rgba(37,99,235,0.25); border-radius:8px; background:rgba(37,99,235,0.06); font-family:'JetBrains Mono', 'Fira Code', 'Segoe UI Mono', monospace; font-size:0.84rem; font-weight:700; color:#1e40af; box-shadow:0 1px 2px rgba(0,0,0,0.03);">
                ${escapeAuthorizationText(item)}
              </div>
            `).join('')}
          </div>
        </div>
      `;
      appendAuthorizationMessage('assistant', html, true);
    }

    function renderAuthorizationTable(title, columns, rows, emptyMessage) {
      if (!Array.isArray(rows) || rows.length === 0) {
        appendAuthorizationMessage('assistant', emptyMessage);
        return;
      }

      const thead = `<tr>${columns.map((column) => `<th>${escapeAuthorizationText(column)}</th>`).join('')}</tr>`;
      const tbody = rows.map((row) => {
        const cells = Array.isArray(row) ? row : [];
        return `<tr>${cells.map((cell) => `<td style="vertical-align:top;">${escapeAuthorizationText(cell)}</td>`).join('')}</tr>`;
      }).join('');

      const html = `
        <div class="auth-chat-summary">
          <div class="auth-chat-summary-row" style="font-weight:700; margin-bottom:10px;">
            <span class="auth-chat-summary-label">${escapeAuthorizationText(title)}</span>
            <span class="auth-chat-summary-value">${String(rows.length)}</span>
          </div>
          <div class="auth-table-wrapper">
            <table class="auth-result-table">
              <thead>${thead}</thead>
              <tbody>${tbody}</tbody>
            </table>
          </div>
        </div>
      `;
      appendAuthorizationMessage('assistant', html, true);
    }
    async function pollAuthorizationRemovalJob(jobId, requestId) {
      const startTime = Date.now();
      const timeoutMs = 180000;

      async function check() {
        if (requestId !== authorizationRemovalJobRequestId) {
          return;
        }

        if (Date.now() - startTime > timeoutMs) {
          if (requestId === authorizationRemovalJobRequestId) {
            hideAuthorizationTypingIndicator();
            appendAuthorizationMessage(
              'assistant',
              `O job de remocao #${String(jobId || '').slice(0, 8)} demorou demasiado a concluir. Vai a lista de jobs para confirmar o estado final.`
            );
          }
          return;
        }

        try {
          const response = await fetchWithTimeout(`/api/jobs/${jobId}`, {}, 10000);
          if (!response.ok) {
            throw new Error(`Erro HTTP ${response.status}`);
          }

          const job = await response.json();
          if (requestId !== authorizationRemovalJobRequestId) {
            return;
          }

          if (!isAuthorizationTerminalJobState(job.state)) {
            window.setTimeout(check, 2000);
            return;
          }

          hideAuthorizationTypingIndicator();

          if (String(job.state || '').toLowerCase() === 'failed' || String(job.state || '').toLowerCase() === 'error') {
            appendAuthorizationMessage(
              'assistant',
              `A remocao terminou com erro no job #${String(jobId || '').slice(0, 8)}. ${String(job.status || 'Consulta o log do job para ver o detalhe.')}`
            );
            return;
          }

          appendAuthorizationMessage('assistant', buildAuthorizationRemovalFeedback(job));
        } catch (error) {
          if (requestId !== authorizationRemovalJobRequestId) {
            return;
          }

          window.setTimeout(check, 3000);
        }
      }

      window.setTimeout(check, 1200);
    }

    function handleAuthorizationFollowUp(rawQuery) {
      const query = String(rawQuery || '').trim().toLowerCase();
      const queryNorm = normalizeAuthorizationSearchText(query);

      // 0. Se o utilizador pedir para iniciar uma nova análise ou limpar a conversa
      if (
        queryNorm.includes('NOVA ANALISE') ||
        queryNorm.includes('REINICIAR') ||
        queryNorm.includes('LIMPAR') ||
        queryNorm.includes('NOVO UTILIZADOR') ||
        queryNorm.includes('NOVA PESQUISA')
      ) {
        resetAuthorizationChat();
        return;
      }

      // 0.1 Se o utilizador pedir para executar um processo SAP
      if (
        queryNorm.includes('EXECUTAR PROCESSO') ||
        queryNorm.includes('EXECUTAR') ||
        queryNorm.includes('CRIAR ROLE') ||
        queryNorm.includes('ELIMINAR ROLE')
      ) {
        showPfcgProcessExecutionOptions();
        return;
      }

      // 1. Verificar se o utilizador está a perguntar sobre um Processo ou Skill
      const isProcessQuery = /processo|processos|skill|skills|modulo|modulos/i.test(queryNorm);
      if (isProcessQuery || queryNorm.includes('AUTORIZA') || queryNorm.includes('PFCG') || queryNorm.includes('CODIGOS IVA') || queryNorm.includes('CADEIAS')) {
        if (queryNorm.includes('AUTORIZA') || queryNorm.includes('PERFIL') || queryNorm.includes('PFCG') || queryNorm.includes('CUA')) {
          promptInitialSystemSelection('Perfil de autorização');
          return;
        }

        if (queryNorm.includes('IVA') || queryNorm.includes('IMPOSTO')) {
          appendAuthorizationMessage('assistant', `
            <div class="auth-chat-summary">
              <div style="font-weight:700; margin-bottom:6px; color:#2563eb;">📋 Processo: Códigos IVA</div>
              <div style="font-size:0.83rem;">Automatização da transação <b>FTXP</b> para criação e manutenção de códigos e taxas de IVA no SAP.</div>
            </div>
          `, true);
          return;
        }

        if (queryNorm.includes('CADEIA') || queryNorm.includes('EXTRATO')) {
          appendAuthorizationMessage('assistant', `
            <div class="auth-chat-summary">
              <div style="font-weight:700; margin-bottom:6px; color:#2563eb;">📋 Processo: Cadeias de Pesquisa</div>
              <div style="font-size:0.83rem;">Configuração e atribuição de cadeias de pesquisa de extratos bancários em FI.</div>
            </div>
          `, true);
          return;
        }

        if (queryNorm.includes('BANCO') || queryNorm.includes('CHAVE')) {
          appendAuthorizationMessage('assistant', `
            <div class="auth-chat-summary">
              <div style="font-weight:700; margin-bottom:6px; color:#2563eb;">📋 Processo: Chave de Banco</div>
              <div style="font-size:0.83rem;">Criação e manutenção automatizada de chaves de banco (dados mestres bancários FI).</div>
            </div>
          `, true);
          return;
        }

        if (queryNorm.includes('REVERTER') || queryNorm.includes('ESTORNO') || queryNorm.includes('DOCUMENTO')) {
          appendAuthorizationMessage('assistant', `
            <div class="auth-chat-summary">
              <div style="font-weight:700; margin-bottom:6px; color:#2563eb;">📋 Processo: Reverter Documento</div>
              <div style="font-size:0.83rem;">Anulação e reversão automatizada de documentos contabilísticos no SAP.</div>
            </div>
          `, true);
          return;
        }

        if (isProcessQuery) {
          appendAuthorizationMessage('assistant', `
            <div class="auth-chat-summary">
              <div style="font-weight:700; margin-bottom:6px; color:#2563eb;">📋 Processos SAP Suportados pelo Cockpit</div>
              <div style="font-size:0.8rem; display:grid; gap:4px;">
                <div>• <b>Funções PFCG & Autorizações:</b> Perfis simples, compostos e utilizadores CUA.</div>
                <div>• <b>Códigos IVA:</b> Manutenção de taxas de IVA (FTXP).</div>
                <div>• <b>Cadeias de Pesquisa:</b> Regras de extrato bancário FI.</div>
                <div>• <b>Chave de Banco:</b> Dados mestres bancários.</div>
                <div>• <b>CUA Login:</b> Reset de logins CUA.</div>
                <div>• <b>Reverter Documento:</b> Anulação de lançamentos contabilísticos.</div>
              </div>
            </div>
          `, true);
          return;
        }
      }

      const data = authorizationLastStatusData;
      const roles = Array.isArray(data?.roles) ? data.roles : [];
      const requestedRole = extractAuthorizationRoleFromQuery(query, roles);
      const requestedTransaction = extractAuthorizationTransactionCode(query);
      const wantsTransactions = /(?:quais|quaisas|lista|list|mostra|ver|tabela|pesquisar|procurar)?\s*(?:transa|tcode|tcodes|codigos|t-code)/i.test(queryNorm) || queryNorm.includes('TCODE') || queryNorm.includes('TRANSACOES') || queryNorm.includes('TRANSAÇÃO');
      const wantsRemoval = /elimin|eliin|elimn|remov|remoç|remoc|apag|exclui|tirar|retir|delet|borrar/i.test(queryNorm);
      const wantsAllRoles = query.includes('todas') || query.includes('todos') || query.includes('tudo');
      const isExpired = query.includes('expir');
      const isActive = query.includes('ativa') || query.includes('activo') || query.includes('ativo') || query.includes('active');
      const isDirect = query.includes('diret') || query.includes('direct');
      const isIndirect = query.includes('indiret') || query.includes('indirect');

      if (!data || !Array.isArray(data.roles)) {
        if (requestedTransaction) {
          appendAuthorizationMessage(
            'assistant',
            `Ainda não fez a análise do utilizador. Pode introduzir o utilizador SAP para analisar as suas funções ou perguntar por qualquer processo SAP.`
          );
          return;
        }

        appendAuthorizationMessage(
          'assistant',
          'Pode introduzir o nome de um utilizador SAP para iniciar a análise de autorizações, ou perguntar sobre os processos e rotinas SAP disponíveis.'
        );
        return;
      }

      if (wantsRemoval) {
        const candidateRoles = authorizationLastDisplayedRoles.length > 0
          ? authorizationLastDisplayedRoles
          : (wantsAllRoles ? roles : []);

        if (!candidateRoles.length) {
          appendAuthorizationMessage(
            'assistant',
            'Para remover funções, primeiro pede a lista filtrada ou indica "remova as funções mostradas" / "remova todas as funções do utilizador".'
          );
          return;
        }

        renderAuthorizationRemovalPrompt(
          candidateRoles,
          wantsAllRoles || authorizationLastDisplayedRoles.length === 0
            ? 'as funções do utilizador'
            : 'as funções mostradas na lista'
        );
        return;
      }

      let filtered = roles.slice();
      if (isExpired) {
        filtered = filtered.filter(r => String(r.validity_status || '').toLowerCase() === 'expired');
      } else if (isActive) {
        filtered = filtered.filter(r => String(r.validity_status || '').toLowerCase() === 'active');
      }

      if (isDirect) {
        filtered = filtered.filter(r => {
          const origin = String(r.assignment_origin_label || r.assignment_origin || '').toLowerCase();
          return origin.includes('direta') || origin.includes('direct');
        });
      } else if (isIndirect) {
        filtered = filtered.filter(r => {
          const origin = String(r.assignment_origin_label || r.assignment_origin || '').toLowerCase();
          return origin.includes('indireta') || origin.includes('indirect');
        });
      }

      if (isExpired || isActive || isDirect || isIndirect) {
        renderAuthorizationFollowUpResult(
          isExpired ? 'Funções expiradas' : isActive ? 'Funções ativas' : isDirect ? 'Funções diretas' : 'Funções indiretas',
          filtered,
          'Nenhuma função corresponde ao filtro pedido.'
        );
        return;
      }

      if (wantsTransactions || requestedRole || requestedTransaction) {
        const allFunctions = getAuthorizationAllFunctions(data);

        if (requestedRole) {
          const roleFunctions = getAuthorizationRoleFunctions(data, requestedRole);
          if (roleFunctions.length > 0) {
            if (roleFunctions.length <= 8) {
              renderAuthorizationTable(
                `Transações da role ${requestedRole}`,
                ['Transação'],
                roleFunctions.map((item) => [item]),
                `Não encontrei transações associadas à role ${requestedRole} na análise atual.`
              );
            } else {
              renderAuthorizationPlainList(
                `Transações da role ${requestedRole}`,
                roleFunctions,
                `Não encontrei transações associadas à role ${requestedRole} na análise atual.`
              );
            }
            return;
          }

          const requestedToken = normalizeAuthorizationToken(requestedRole).replace(/^ROLE\s+/, '');
          const similarRoles = roles
            .map((role) => String(role?.role || role?.name || role?.function || '').trim())
            .filter(Boolean)
            .filter((roleName) => normalizeAuthorizationToken(roleName).includes(requestedToken))
            .slice(0, 8);

          appendAuthorizationMessage(
            'assistant',
            similarRoles.length > 0
              ? `Não encontrei transações associadas à role ${requestedRole} na análise atual. Roles parecidas nesta lista: ${similarRoles.join(', ')}.`
              : `Não encontrei transações associadas à role ${requestedRole} na análise atual. Se esta análise foi feita em CUA, pode não ter detalhe de AGR_TCODES.`
          );
          return;
        }

        if (requestedTransaction) {
          const matchingRoles = getAuthorizationRolesWithTransaction(data, requestedTransaction);
          if (matchingRoles.length > 0) {
            renderAuthorizationTable(
              `Funções com a transação ${requestedTransaction}`,
              ['Função'],
              matchingRoles.map((item) => [item]),
              `Não encontrei funções com a transação ${requestedTransaction} na análise atual.`
            );
            return;
          }

          appendAuthorizationMessage(
            'assistant',
            `Não encontrei funções com a transação ${requestedTransaction} na análise atual. Se a análise foi feita em CUA, pode não existir detalhe de AGR_TCODES por role.`
          );
          return;
        }

        if (allFunctions.length > 0) {
          renderAuthorizationPlainList(
            'Lista de funções/TCODEs',
            allFunctions,
            'Não tenho funções disponíveis na análise atual.'
          );
          return;
        }

        appendAuthorizationMessage('assistant', 'Não tenho funções disponíveis na análise atual.');
        return;
      }

      appendAuthorizationMessage(
        'assistant',
        `
          <div style="font-weight:700; margin-bottom:8px;">Deseja continuar a analisar ou pretende executar um processo para esta análise?</div>
          <div style="margin-top:10px; display:flex; gap:10px; flex-wrap:wrap;">
            <button type="button" class="btn btn-secondary btn-sm" onclick="resetAuthorizationChat()" style="display:inline-flex; align-items:center; gap:6px; cursor:pointer; padding:6px 14px; border-radius:8px; font-weight:600;">🔄 Nova análise</button>
            <button type="button" class="btn btn-primary btn-sm" onclick="showPfcgProcessExecutionOptions()" style="display:inline-flex; align-items:center; gap:6px; cursor:pointer; padding:6px 14px; border-radius:8px; font-weight:600;">⚙️ Executar processo SAP</button>
          </div>
        `,
        true
      );
    }

    let authorizationAnalysisType = 'authorizations';

    function showAnalysisTypeSelection() {
      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage('assistant', 'Selecione qual o tipo de análise que pretende realizar para a **Análise de Autorizações SAP**:');

      const grid = document.createElement('div');
      grid.style.display = 'grid';
      grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(240px, 1fr))';
      grid.style.gap = '10px';

      // Opção 1: Autorizações
      const btnAuth = document.createElement('button');
      btnAuth.type = 'button';
      btnAuth.className = 'auth-chat-analysis-card';
      btnAuth.style.padding = '12px 14px';
      btnAuth.onclick = () => {
        if (btnAuth.parentElement) {
          btnAuth.parentElement.querySelectorAll('button').forEach(b => {
            b.classList.remove('selected');
            b.setAttribute('aria-pressed', 'false');
          });
        }
        btnAuth.classList.add('selected');
        btnAuth.setAttribute('aria-pressed', 'true');
        appendAuthorizationMessage('user', '🔐 Autorizações');
        authorizationAnalysisType = 'authorizations';
        askTargetUserForAnalysis();
      };
      btnAuth.innerHTML = `
        <span class="analysis-title">🔐 Autorizações</span>
        <span class="analysis-desc">Análise detalhada de funções PFCG, perfis de autorização e acessos à USLA04.</span>
      `;
      grid.appendChild(btnAuth);

      // Opção 2: Dados Mestre
      const btnMaster = document.createElement('button');
      btnMaster.type = 'button';
      btnMaster.className = 'auth-chat-analysis-card';
      btnMaster.style.padding = '12px 14px';
      btnMaster.onclick = () => {
        if (btnMaster.parentElement) {
          btnMaster.parentElement.querySelectorAll('button').forEach(b => {
            b.classList.remove('selected');
            b.setAttribute('aria-pressed', 'false');
          });
        }
        btnMaster.classList.add('selected');
        btnMaster.setAttribute('aria-pressed', 'true');
        appendAuthorizationMessage('user', '👤 Dados Mestre');
        authorizationAnalysisType = 'master_data';
        askTargetUserForAnalysis();
      };
      btnMaster.innerHTML = `
        <span class="analysis-title">👤 Dados Mestre</span>
        <span class="analysis-desc">Análise de conta de utilizador, estado de bloqueio (USR02), dados pessoais e e-mail.</span>
      `;
      grid.appendChild(btnMaster);

      const container = document.getElementById('authorization-chat-messages');
      if (container) {
        container.appendChild(grid);
        container.scrollTop = container.scrollHeight;
      }
    }

    function askTargetUserForAnalysis() {
      authorizationTargetUser = '';
      authorizationSelectedSystem = null;
      authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage('assistant', 'Por favor, indique qual é o utilizador SAP que pretende analisar (ex: CSILVA ou U1234):');
      updateAuthorizationComposer();
      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    function promptProcessMode(processName, category, scriptName, description) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();

      appendAuthorizationAssistantMessage(`Deseja efetuar a alteração para **${escapeAuthorizationText(processName)}** em lote (massiva) ou para um utilizador individual?`);

      const grid = document.createElement('div');
      grid.style.display = 'grid';
      grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(240px, 1fr))';
      grid.style.gap = '10px';

      // Opção Massiva (Abre o formulário de Job via menu com ficheiro Excel)
      const btnMassive = document.createElement('button');
      btnMassive.type = 'button';
      btnMassive.className = 'auth-chat-analysis-card';
      btnMassive.style.padding = '12px 14px';
      btnMassive.onclick = () => {
        if (btnMassive.parentElement) {
          btnMassive.parentElement.querySelectorAll('button').forEach(b => {
            b.classList.remove('selected');
            b.setAttribute('aria-pressed', 'false');
          });
        }
        btnMassive.classList.add('selected');
        btnMassive.setAttribute('aria-pressed', 'true');
        appendAuthorizationMessage('user', '📊 Alteração Massiva (Ficheiro Excel)');
        abrirSubprocessoModal(category, scriptName);
        appendAuthorizationMessage('assistant', `A abrir formulário de **Alteração Massiva (Job em lote)** para ${escapeAuthorizationText(processName)}...`);
      };
      btnMassive.innerHTML = `
        <span class="analysis-title">📊 Alteração Massiva (Ficheiro Excel)</span>
        <span class="analysis-desc">Processar em lote através de ficheiro Excel (formulário de Job do menu).</span>
      `;
      grid.appendChild(btnMassive);

      // Opção Individual (Chat Direto: Pergunta utilizador e depois ambiente)
      const btnIndividual = document.createElement('button');
      btnIndividual.type = 'button';
      btnIndividual.className = 'auth-chat-analysis-card';
      btnIndividual.style.padding = '12px 14px';
      btnIndividual.onclick = () => {
        if (btnIndividual.parentElement) {
          btnIndividual.parentElement.querySelectorAll('button').forEach(b => {
            b.classList.remove('selected');
            b.setAttribute('aria-pressed', 'false');
          });
        }
        btnIndividual.classList.add('selected');
        btnIndividual.setAttribute('aria-pressed', 'true');
        appendAuthorizationMessage('user', '👤 Alteração Individual (Chat Direto)');
        startIndividualProcessFlow(processName, category, scriptName);
      };
      btnIndividual.innerHTML = `
        <span class="analysis-title">👤 Alteração Individual (Chat Direto)</span>
        <span class="analysis-desc">Efetuar alteração passo a passo diretamente aqui no assistente.</span>
      `;
      grid.appendChild(btnIndividual);

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function startIndividualProcessFlow(processName, category, scriptName) {
      authorizationIndividualContext = {
        processName,
        category,
        scriptName,
        targetUser: '',
        selectedSystem: authorizationSelectedSystem || null,
        parameters: {}
      };

      authorizationTargetUser = '';
      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER;

      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        `Para a **Alteração Individual** do processo **${escapeAuthorizationText(processName)}**, por favor indique qual é o utilizador SAP sobre o qual pretende efetuar a alteração (ex: CSILVA ou U1234):`
      );
      updateAuthorizationComposer();

      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    function showIndividualSystemOptions(onSelectSystem = null) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_SYSTEM;
      updateAuthorizationComposer();

      const grid = document.createElement('div');
      grid.id = 'individual-system-options';
      grid.className = 'auth-chat-system-grid';

      if (authorizationAvailableSystems.length === 0) {
        const noSystems = document.createElement('div');
        noSystems.style.color = '#ef4444';
        noSystems.style.fontWeight = '600';
        noSystems.textContent = 'Nenhum sistema SAP foi encontrado no .env.';
        container.appendChild(noSystems);
        container.scrollTop = container.scrollHeight;
        return;
      }

      authorizationAvailableSystems.forEach(sys => {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'auth-chat-system-card';
        card.setAttribute('aria-pressed', 'false');
        card.setAttribute('data-key', sys.key);

        const codeSpan = document.createElement('span');
        codeSpan.className = 'sys-code';
        codeSpan.textContent = sys.system;

        const clientSpan = document.createElement('span');
        clientSpan.className = 'sys-client';
        clientSpan.textContent = `Cliente ${sys.client}`;

        card.appendChild(codeSpan);
        card.appendChild(clientSpan);

        if (sys.connection_name) {
          const connSpan = document.createElement('span');
          connSpan.className = 'sys-conn';
          connSpan.textContent = sys.connection_name;
          card.appendChild(connSpan);
        }

        if (sys.execution_mode) {
          const modeSpan = document.createElement('span');
          modeSpan.className = 'sys-conn';
          modeSpan.textContent = sys.execution_mode;
          card.appendChild(modeSpan);
        }

        card.onclick = () => {
          if (card.parentElement) {
            card.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          card.classList.add('selected');
          card.setAttribute('aria-pressed', 'true');

          if (typeof onSelectSystem === 'function') {
            onSelectSystem(sys);
          } else {
            selectIndividualSystem(sys);
          }
        };

        grid.appendChild(card);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function promptInitialSystemSelection(categoryName = 'Perfil de autorização') {
      hideAuthorizationTypingIndicator();
      appendAuthorizationAssistantMessage(`Registado o interesse em **${escapeAuthorizationText(categoryName)}**. Em que sistema/ambiente pretende trabalhar?`);
      showIndividualSystemOptions((sys) => {
        authorizationSelectedSystem = sys;
        const sysLabel = sys.label || sys.system || sys.key;
        appendAuthorizationMessage('user', sysLabel);
        appendAuthorizationMessage(
          'assistant',
          `Ambiente **${escapeAuthorizationText(sysLabel)}** registado. Selecione qual das tarefas ou processos de **Funções PFCG & Autorizações** pretende realizar:`
        );
        showPfcgProcessExecutionOptions();
      });
    }

    async function selectIndividualSystem(sys) {
      if (!authorizationIndividualContext) return;
      authorizationIndividualContext.selectedSystem = sys;
      authorizationSelectedSystem = sys;

      const user = authorizationIndividualContext.targetUser;
      const proc = authorizationIndividualContext.processName;
      const sysLabel = sys.label || sys.system || sys.key;

      appendAuthorizationMessage('user', sysLabel);
      appendAuthorizationMessage(
        'assistant',
        `Perfeito. Registado o ambiente **${escapeAuthorizationText(sysLabel)}** para a alteração individual do utilizador **${escapeAuthorizationText(user)}** (${escapeAuthorizationText(proc)}).`
      );

      const script = (authorizationIndividualContext.scriptName || proc || '').toUpperCase();
      const isCuaProcess = script.includes('CUA') || proc.includes('CUA');
      const cuaSystemKey = window.authorizationCuaSystemKey || 'SPACLNT001';
      const targetSysKey = isCuaProcess ? cuaSystemKey : sys.key;

      // Tentar obter dinamicamente as funções/roles atribuídas a este utilizador no CUA (SE16 -> USLA04)
      showAuthorizationTypingIndicator(
        null,
        isCuaProcess
          ? `A consultar a tabela CUA USLA04 via GUI (SE16) para ${escapeAuthorizationText(user)} no subsistema ${escapeAuthorizationText(sys.system || sys.key)}...`
          : `A consultar funções de ${escapeAuthorizationText(user)} no sistema ${escapeAuthorizationText(sys.system || sys.key)}...`
      );

      try {
        const payload = {
          target_user: user,
          target_system_key: targetSysKey,
          subsystem_filter: sys.key,
          analysis_type: authorizationAnalysisType || 'authorizations'
        };

        const response = await fetch('/api/authorizations/start', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        if (data.job_id) {
          await pollIndividualUserRoles(data.job_id, user, sys);
          return;
        }
      } catch (err) {
        console.warn('[INDIVIDUAL FLOW] Não foi possível consultar roles dinamicamente:', err);
      }

      hideAuthorizationTypingIndicator();
      promptIndividualProcessParameters();
    }

    async function pollIndividualUserRoles(jobId, targetUser, sys) {
      const startTime = Date.now();
      const timeoutMs = 60000;

      async function check() {
        if (Date.now() - startTime > timeoutMs) {
          hideAuthorizationTypingIndicator();
          promptIndividualProcessParameters();
          return;
        }

        try {
          const response = await fetch(`/api/jobs/${jobId}`);
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const job = await response.json();

          if (job.state === 'succeeded' || job.state === 'succeeded_with_warnings') {
            hideAuthorizationTypingIndicator();
            let statusData = {};
            if (typeof job.status === 'string') {
              try {
                statusData = JSON.parse(job.status);
              } catch (e) {}
            } else if (job.status && typeof job.status === 'object') {
              statusData = job.status;
            } else if (job.result && typeof job.result === 'object') {
              statusData = job.result;
            }

            const roles = Array.isArray(statusData.roles) ? statusData.roles : (Array.isArray(job.result?.roles) ? job.result.roles : []);
            if (roles.length > 0) {
              renderIndividualUserRolesList(roles, targetUser, sys);
            } else {
              const sysName = sys.system || sys.label || sys.key;
              appendAuthorizationMessage(
                'assistant',
                `Para os parâmetros informados (Utilizador: **${escapeAuthorizationText(targetUser)}** | Sistema: **${escapeAuthorizationText(sysName)}**), não temos informações na tabela CUA **USLA04**.`
              );
              promptIndividualProcessParameters();
            }
            return;
          } else if (job.state === 'failed' || job.state === 'error' || job.state === 'cancelled') {
            hideAuthorizationTypingIndicator();
            promptIndividualProcessParameters();
            return;
          }

          window.setTimeout(check, 1500);
        } catch (e) {
          window.setTimeout(check, 2500);
        }
      }

      window.setTimeout(check, 1000);
    }

    async function fetchCuaGlobalUserSummary(userVal) {
      const procName = authorizationIndividualContext?.processName || 'CUA_REMOVE';
      hideAuthorizationTypingIndicator();
      showAuthorizationTypingIndicator(
        null,
        `A consultar a tabela CUA USLA04 para ${escapeAuthorizationText(userVal)} em todos os subsistemas...`
      );

      try {
        const defaultSys = (Array.isArray(authorizationAvailableSystems) ? authorizationAvailableSystems : []).find(s => s.system === 'SPA' || s.key === 'SPACLNT001') || authorizationAvailableSystems?.[0];
        const activeSysKey = authorizationSelectedSystem?.key || authorizationIndividualContext?.selectedSystem?.key || defaultSys?.key || 'SPACLNT001';
        const payload = {
          target_user: userVal,
          target_system_key: activeSysKey,
          subsystem_filter: '',
          analysis_type: 'authorizations'
        };

        const response = await fetch('/api/authorizations/start', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        if (data.job_id) {
          await pollCuaGlobalSummaryJob(data.job_id, userVal);
          return;
        }
      } catch (err) {
        console.warn('[CUA SUMMARY] Não foi possível consultar USLA04 globalmente:', err);
      }

      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        `Registado o utilizador **${escapeAuthorizationText(userVal)}** para a alteração em **${escapeAuthorizationText(procName)}**.\n\nEm que sistema/ambiente pretende efetuar a alteração?`
      );
      showIndividualSystemOptions();
    }

    async function pollCuaGlobalSummaryJob(jobId, targetUser) {
      const startTime = Date.now();
      const timeoutMs = 60000;

      async function check() {
        if (Date.now() - startTime > timeoutMs) {
          hideAuthorizationTypingIndicator();
          showIndividualSystemOptions();
          return;
        }

        try {
          const response = await fetch(`/api/jobs/${jobId}`);
          if (!response.ok) throw new Error(`HTTP ${response.status}`);
          const job = await response.json();

          if (job.state === 'succeeded') {
            hideAuthorizationTypingIndicator();
            let statusData = {};
            try {
              statusData = JSON.parse(job.status);
            } catch (e) {}

            const systemsSummary = Array.isArray(statusData.systems_summary) ? statusData.systems_summary : [];
            const allRoles = Array.isArray(statusData.roles) ? statusData.roles : [];

            if (systemsSummary.length > 0 || allRoles.length > 0) {
              renderCuaSystemsSummaryTable(systemsSummary, allRoles, targetUser);
            } else {
              appendAuthorizationMessage(
                'assistant',
                `Não foram encontradas funções ativas atribuídas ao utilizador **${escapeAuthorizationText(targetUser)}** na tabela CUA USLA04.`
              );
              showIndividualSystemOptions();
            }
            return;
          } else if (job.state === 'failed' || job.state === 'error' || job.state === 'cancelled') {
            hideAuthorizationTypingIndicator();
            showIndividualSystemOptions();
            return;
          }

          window.setTimeout(check, 1500);
        } catch (e) {
          window.setTimeout(check, 2500);
        }
      }

      window.setTimeout(check, 1000);
    }

    function renderCuaSystemsSummaryTable(systemsSummary, allRoles, targetUser) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const procName = authorizationIndividualContext?.processName || 'CUA_REMOVE (GUI)';

      if ((!systemsSummary || systemsSummary.length === 0) && allRoles.length > 0) {
        const map = {};
        allRoles.forEach(r => {
          const sys = r.subsystem || 'OUTROS';
          if (!map[sys]) {
            const sysName = sys.split('CLNT')[0] || sys;
            map[sys] = { subsystem: sys, system: sysName, roles_count: 0 };
          }
          map[sys].roles_count++;
        });
        systemsSummary = Object.values(map);
      }

      appendAuthorizationMessage(
        'assistant',
        `Leitura concluída da tabela CUA **USLA04** para o utilizador **${escapeAuthorizationText(targetUser)}**.\n\nAbaixo encontra-se a distribuição de funções por sistema/subsistema. Clique no sistema pretendido para selecionar a(s) função(ões) a alterar:`
      );

      const wrapper = document.createElement('div');
      wrapper.className = 'auth-chat-summary';
      wrapper.style.width = '100%';

      const header = document.createElement('div');
      header.style.fontWeight = '700';
      header.style.marginBottom = '12px';
      header.style.fontSize = '0.92rem';
      header.style.color = '#2563eb';
      header.textContent = `📊 Funções CUA do utilizador ${targetUser} (Agrupadas por Sistema)`;
      wrapper.appendChild(header);

      const tableWrapper = document.createElement('div');
      tableWrapper.className = 'auth-table-wrapper';

      const table = document.createElement('table');
      table.className = 'auth-result-table';
      table.innerHTML = `
        <thead>
          <tr>
            <th>Sistema / Subsistema</th>
            <th>Quantidade de Funções</th>
            <th style="text-align:center;">Ação</th>
          </tr>
        </thead>
        <tbody></tbody>
      `;

      const tbody = table.querySelector('tbody');
      systemsSummary.forEach(item => {
        const tr = document.createElement('tr');
        const sysLabel = item.system || item.subsystem;
        const count = item.roles_count || 0;

        const tdSys = document.createElement('td');
        tdSys.style.fontWeight = '600';
        tdSys.textContent = `${sysLabel} (${item.subsystem})`;

        const tdCount = document.createElement('td');
        tdCount.innerHTML = `<span style="background:#2563eb; color:white; padding:2px 10px; border-radius:12px; font-weight:700; font-size:0.8rem;">${count} ${count === 1 ? 'função' : 'funções'}</span>`;

        const tdAction = document.createElement('td');
        tdAction.style.textAlign = 'center';

        const btnSelect = document.createElement('button');
        btnSelect.type = 'button';
        btnSelect.className = 'btn btn-primary btn-sm';
        btnSelect.style.padding = '4px 12px';
        btnSelect.style.fontSize = '0.8rem';
        btnSelect.style.fontWeight = '600';
        btnSelect.style.borderRadius = '6px';
        btnSelect.textContent = `Selecionar ${sysLabel}`;
        btnSelect.onclick = () => {
          const matchingSys = authorizationAvailableSystems.find(s => 
            s.key === item.subsystem || s.system === item.system || s.key.startsWith(item.system)
          ) || { key: item.subsystem, system: item.system, label: item.system };

          authorizationIndividualContext.selectedSystem = matchingSys;
          authorizationSelectedSystem = matchingSys;

          const subRoles = allRoles.filter(r => 
            !r.subsystem || r.subsystem === item.subsystem || r.subsystem.startsWith(item.system)
          );

          appendAuthorizationMessage('user', `Selecionado o ambiente ${matchingSys.system || matchingSys.key}`);
          renderIndividualUserRolesList(subRoles.length > 0 ? subRoles : allRoles, targetUser, matchingSys);
        };

        tdAction.appendChild(btnSelect);
        tr.appendChild(tdSys);
        tr.appendChild(tdCount);
        tr.appendChild(tdAction);
        tbody.appendChild(tr);
      });

      tableWrapper.appendChild(table);
      wrapper.appendChild(tableWrapper);

      container.appendChild(wrapper);
      container.scrollTop = container.scrollHeight;
    }

    function renderIndividualUserRolesList(roles, targetUser, sys) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const ctx = authorizationIndividualContext || {};
      const procName = ctx.processName || 'Processo';
      const script = (ctx.scriptName || '').toUpperCase();

      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS;

      const isRemoval = script.includes('CUA_REMOVE') || script.includes('PFCG_DELETE');
      const isEndDate = script.includes('CUA_ENDDATE');

      const actionTitle = isRemoval
        ? 'Selecione a(s) função(ões) que pretende **remover**'
        : isEndDate
          ? 'Selecione a(s) função(ões) que pretende **alterar a validade**'
          : 'Seleção de funções do utilizador no sistema';

      appendAuthorizationAssistantMessage(`${actionTitle} para o utilizador **${escapeAuthorizationText(targetUser)}** no sistema **${escapeAuthorizationText(sys.system || sys.key)}**:`);

      const wrapper = document.createElement('div');
      wrapper.className = 'auth-chat-summary';

      // Header com contagem total
      const header = document.createElement('div');
      header.className = 'auth-chat-summary-row';
      header.style.fontWeight = '700';
      header.style.paddingBottom = '10px';
      header.innerHTML = `<span class="auth-chat-summary-label">Funções atribuídas a ${escapeAuthorizationText(targetUser)} em ${escapeAuthorizationText(sys.system || sys.key)}</span><span class="auth-chat-summary-value">${String(roles.length)}</span>`;
      wrapper.appendChild(header);

      // Barra de ferramentas: Checkbox "Marcar todas" e badge de seleção
      const toolbar = document.createElement('div');
      toolbar.style.display = 'flex';
      toolbar.style.alignItems = 'center';
      toolbar.style.justifyContent = 'space-between';
      toolbar.style.padding = '8px 12px';
      toolbar.style.marginBottom = '10px';
      toolbar.style.background = 'rgba(37, 99, 235, 0.05)';
      toolbar.style.border = '1px solid rgba(37, 99, 235, 0.2)';
      toolbar.style.borderRadius = '10px';

      const selectAllLabel = document.createElement('label');
      selectAllLabel.style.display = 'flex';
      selectAllLabel.style.alignItems = 'center';
      selectAllLabel.style.gap = '8px';
      selectAllLabel.style.cursor = 'pointer';
      selectAllLabel.style.fontWeight = '700';
      selectAllLabel.style.fontSize = '0.85rem';
      selectAllLabel.style.color = 'var(--text-primary)';

      const selectAllCheckbox = document.createElement('input');
      selectAllCheckbox.type = 'checkbox';
      selectAllCheckbox.style.width = '17px';
      selectAllCheckbox.style.height = '17px';
      selectAllCheckbox.style.cursor = 'pointer';
      selectAllCheckbox.style.accentColor = '#2563eb';

      const selectAllText = document.createElement('span');
      selectAllText.textContent = `☑️ Marcar todas (${roles.length})`;

      selectAllLabel.appendChild(selectAllCheckbox);
      selectAllLabel.appendChild(selectAllText);
      toolbar.appendChild(selectAllLabel);

      const countBadge = document.createElement('span');
      countBadge.style.fontSize = '0.8rem';
      countBadge.style.fontWeight = '700';
      countBadge.style.color = '#2563eb';
      countBadge.style.background = 'white';
      countBadge.style.padding = '3px 10px';
      countBadge.style.borderRadius = '12px';
      countBadge.style.border = '1px solid rgba(37, 99, 235, 0.3)';
      countBadge.textContent = '0 selecionadas';
      toolbar.appendChild(countBadge);

      wrapper.appendChild(toolbar);

      // Lista de roles com checkboxes
      const listDiv = document.createElement('div');
      listDiv.style.display = 'grid';
      listDiv.style.gap = '8px';
      listDiv.style.maxHeight = '360px';
      listDiv.style.overflowY = 'auto';
      listDiv.style.paddingRight = '4px';

      const checkboxes = [];

      roles.forEach(item => {
        const roleName = String(item.role || item.function || item.agr_name || item.AGR_NAME || '').trim();
        if (!roleName) return;

        const row = document.createElement('label');
        row.style.display = 'flex';
        row.style.alignItems = 'center';
        row.style.gap = '12px';
        row.style.padding = '8px 12px';
        row.style.border = '1px solid var(--border-color)';
        row.style.borderRadius = '10px';
        row.style.background = 'rgba(0,0,0,0.02)';
        row.style.cursor = 'pointer';

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.className = 'role-item-checkbox';
        cb.value = roleName;
        cb.style.width = '16px';
        cb.style.height = '16px';
        cb.style.cursor = 'pointer';
        cb.style.accentColor = '#2563eb';

        checkboxes.push(cb);

        const content = document.createElement('div');
        content.style.display = 'flex';
        content.style.flexDirection = 'column';
        content.style.flex = '1';

        const nameSpan = document.createElement('span');
        nameSpan.style.fontWeight = '700';
        nameSpan.style.fontFamily = 'monospace';
        nameSpan.textContent = roleName;

        const infoSpan = document.createElement('span');
        infoSpan.style.fontSize = '0.78rem';
        infoSpan.style.color = 'var(--text-secondary)';
        const dates = [item.valid_from, item.valid_to, item.assignment_origin_label || item.assignment_origin].filter(Boolean).join(' · ');
        infoSpan.textContent = dates || 'Atribuída';

        content.appendChild(nameSpan);
        content.appendChild(infoSpan);

        row.appendChild(cb);
        row.appendChild(content);

        listDiv.appendChild(row);
      });

      wrapper.appendChild(listDiv);

      // Botão de ação inferior
      const actionsDiv = document.createElement('div');
      actionsDiv.style.marginTop = '14px';
      actionsDiv.style.display = 'flex';
      actionsDiv.style.justifyContent = 'flex-end';

      const btnSubmit = document.createElement('button');
      btnSubmit.type = 'button';
      btnSubmit.className = isRemoval ? 'btn btn-danger btn-sm' : 'btn btn-primary btn-sm';
      btnSubmit.style.padding = '8px 18px';
      btnSubmit.style.fontWeight = '700';
      btnSubmit.style.fontSize = '0.84rem';
      btnSubmit.disabled = true;

      const updateSubmitButtonText = () => {
        const selected = checkboxes.filter(c => c.checked).map(c => c.value);
        const count = selected.length;
        countBadge.textContent = `${count} selecionada${count === 1 ? '' : 's'}`;
        
        btnSubmit.disabled = count === 0;

        if (isRemoval) {
          btnSubmit.textContent = count > 0 
            ? `❌ Remover ${count} função${count === 1 ? '' : 'ões'} selecionada${count === 1 ? '' : 's'}`
            : '❌ Selecione funções para remover';
        } else if (isEndDate) {
          btnSubmit.textContent = count > 0 
            ? `📅 Alterar data de ${count} função${count === 1 ? '' : 'ões'}`
            : '📅 Selecione funções para alterar';
        } else {
          btnSubmit.textContent = count > 0 
            ? `Confirmar ${count} função${count === 1 ? '' : 'ões'}`
            : 'Selecione funções';
        }

        selectAllCheckbox.checked = checkboxes.length > 0 && checkboxes.every(c => c.checked);
        selectAllCheckbox.indeterminate = count > 0 && count < checkboxes.length;
      };

      // Event listener "Marcar todas"
      selectAllCheckbox.addEventListener('change', () => {
        const isChecked = selectAllCheckbox.checked;
        checkboxes.forEach(c => { c.checked = isChecked; });
        updateSubmitButtonText();
      });

      // Event listeners individuais
      checkboxes.forEach(c => {
        c.addEventListener('change', () => {
          updateSubmitButtonText();
        });
      });

      btnSubmit.onclick = () => {
        const selected = checkboxes.filter(c => c.checked).map(c => c.value);
        if (selected.length === 0) return;
        selectSingleRoleAction(selected.join(', '), targetUser, sys, procName);
      };

      actionsDiv.appendChild(btnSubmit);
      wrapper.appendChild(actionsDiv);

      container.appendChild(wrapper);
      container.scrollTop = container.scrollHeight;

      updateAuthorizationComposer();
    }

    function selectSingleRoleAction(roleName, targetUser, sys, procName) {
      if (authorizationIndividualContext) {
        authorizationIndividualContext.parameters.roles = roleName;
      }
      const sysLabel = sys.label || sys.system || sys.key;
      const roleList = String(roleName).split(',').map(r => r.trim()).filter(Boolean);

      appendAuthorizationMessage('user', roleName);
      appendAuthorizationMessage(
        'assistant',
        `
          <div class="auth-chat-summary">
            <div style="font-weight:700; margin-bottom:8px; color:#10b981; font-size:0.92rem;">✅ Pedido de Alteração Individual Confirmado</div>
            <div style="display:grid; gap:4px; font-size:0.84rem; margin-bottom:10px;">
              <div><b>• Processo:</b> ${escapeAuthorizationText(procName)}</div>
              <div><b>• Utilizador SAP:</b> ${escapeAuthorizationText(targetUser)}</div>
              <div><b>• Sistema/Ambiente:</b> ${escapeAuthorizationText(sysLabel)}</div>
              <div><b>• Função(ões) Selecionada(s):</b> ${escapeAuthorizationText(roleName)}</div>
            </div>
            <div style="font-size:0.8rem; color:var(--text-secondary);">A submeter o job para execução automática no worker SAP...</div>
          </div>
        `,
        true
      );

      authorizationChatState = AUTH_CHAT_STATES.READY;
      updateAuthorizationComposer();

      // Se for um processo de remoção (ex: CUA_REMOVE ou PFCG_DELETE), submeter a remoção via API automaticamente
      const script = (authorizationIndividualContext?.scriptName || procName || '').toUpperCase();
      if (script.includes('CUA_REMOVE') || script.includes('PFCG_DELETE') || procName.includes('CUA_REMOVE') || procName.includes('DELETE')) {
        authorizationPendingRemoval = {
          targetUser: targetUser,
          targetSystemKey: sys.key,
          systemShort: sys.system,
          roles: roleList.map(r => ({ role: r })),
          label: `${roleList.length} função(ões) selecionada(s)`
        };
        window.setTimeout(() => {
          confirmAuthorizationRemoval();
        }, 400);
      }
    }

    function promptIndividualProcessParameters() {
      if (!authorizationIndividualContext) return;
      const ctx = authorizationIndividualContext;
      const script = (ctx.scriptName || '').toUpperCase();

      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS;

      if (script.includes('CUA_ADICIONAR')) {
        appendAuthorizationMessage('assistant', `Por favor, indique a role/perfil (ou lista de roles separadas por vírgula) que pretende **adicionar** ao utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      } else if (script.includes('CUA_REMOVE')) {
        appendAuthorizationMessage('assistant', `Por favor, indique a role/perfil (ou lista de roles separadas por vírgula) que pretende **remover** do utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      } else if (script.includes('CUA_ENDDATE')) {
        appendAuthorizationMessage('assistant', `Por favor, indique a nova **Data de Fim de Validade** (ex: 31.12.2026) para as roles do utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      } else if (script.includes('PFCG_CREATE')) {
        appendAuthorizationMessage('assistant', `Por favor, indique o nome da **Role PFCG Simples** e as transações (ex: Z_MINHA_ROLE | VA01, SE38) para o utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      } else if (script.includes('PFCG_COMPOSTA')) {
        appendAuthorizationMessage('assistant', `Por favor, indique o nome da **Role Composta** e as roles filhas a associar ao utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      } else if (script.includes('PFCG_DELETE')) {
        appendAuthorizationMessage('assistant', `Por favor, indique a **Role PFCG** que pretende eliminar para o utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      } else {
        appendAuthorizationMessage('assistant', `Por favor, especifique os detalhes/instruções para a alteração individual do utilizador **${escapeAuthorizationText(ctx.targetUser)}**:`);
      }

      updateAuthorizationComposer();
    }

    function showPfcgProcessExecutionOptions() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();

      appendAuthorizationAssistantMessage('Selecione qual das tarefas ou processos de **Funções PFCG & Autorizações** pretende realizar:');

      const grid = document.createElement('div');
      grid.style.display = 'grid';
      grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(220px, 1fr))';
      grid.style.gap = '10px';

      const items = [
        {
          label: '🛡️ Análise de Autorizações SAP',
          desc: 'Análise detalhada de acessos, perfis e funções de um utilizador SAP',
          action: () => showAnalysisTypeSelection()
        },
        {
          label: '🔨 PFCG_CREATE (RFC)',
          desc: 'Criar e atualizar roles simples (TCODEs)',
          action: () => promptProcessMode('PFCG_CREATE (RFC)', 'Funções PFCG', 'A. PFCG_CREATE_RFC.py', 'Criar e atualizar roles simples')
        },
        {
          label: '🧩 PFCG_COMPOSTA (RFC)',
          desc: 'Criar e atualizar roles compostas',
          action: () => promptProcessMode('PFCG_COMPOSTA (RFC)', 'Funções PFCG', 'D. PFCG_COMPOSTA_RFC.py', 'Criar e atualizar roles compostas')
        },
        {
          label: '🗑️ PFCG_DELETE (RFC)',
          desc: 'Eliminação de perfis em lote',
          action: () => promptProcessMode('PFCG_DELETE (RFC)', 'Funções PFCG', 'B. PFCG_DELETE.py', 'Eliminação de perfis')
        },
        {
          label: '🔑 CUA_ADICIONAR (RFC)',
          desc: 'Atribuir perfis a utilizadores CUA',
          action: () => promptProcessMode('CUA_ADICIONAR (RFC)', 'Funções PFCG', 'H. CUA_ADICIONAR.py', 'Atribuir perfis a utilizadores CUA')
        },
        {
          label: '📅 CUA_ENDDATE (GUI)',
          desc: 'Alterar data de fim de validade CUA',
          action: () => promptProcessMode('CUA_ENDDATE (GUI)', 'Funções PFCG', 'I. CUA_ENDDATE.py', 'Alterar data de fim de validade CUA')
        },
        {
          label: '❌ CUA_REMOVE (GUI)',
          desc: 'Remover perfis de utilizadores CUA',
          action: () => promptProcessMode('CUA_REMOVE (GUI)', 'Funções PFCG', 'J. CUA_REMOVE.py', 'Remover perfis de utilizadores CUA')
        }
      ];

      items.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-analysis-card';
        btn.style.padding = '10px 12px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          appendAuthorizationMessage('user', item.label);
          item.action();
        };

        btn.innerHTML = `
          <span class="analysis-title">${escapeAuthorizationText(item.label)}</span>
          <span class="analysis-desc">${escapeAuthorizationText(item.desc)}</span>
        `;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }
    async function confirmAuthorizationRemoval() {
      if (!authorizationPendingRemoval) {
        appendAuthorizationMessage('assistant', 'Não tenho funções pendentes para remover.');
        return;
      }

      const pending = authorizationPendingRemoval;
      const roles = Array.isArray(pending.roles) ? pending.roles : [];
      const payload = {
        target_user: pending.targetUser || authorizationTargetUser,
        target_system_key: pending.targetSystemKey || authorizationSelectedSystem?.key || '',
        roles: roles.map(item => ({
          role: item.role || item.function || item.agr_name || item.AGR_NAME || item.name || ''
        })),
        opcao_processamento: 'sistema_user'
      };

      if (!payload.target_user || !payload.target_system_key || payload.roles.length === 0) {
        appendAuthorizationMessage('assistant', 'Não tenho dados suficientes para criar o pedido de remoção.');
        return;
      }

      appendAuthorizationMessage('assistant', 'A preparar o pedido de remoção das funções selecionadas...');

      try {
        const response = await fetch('/api/authorizations/remove', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        const data = await response.json().catch(() => ({}));
        if (!response.ok) {
          const detail = data.detail || data.message || '';
          const routeMsg = response.status === 404
            ? 'A rota /api/authorizations/remove não foi encontrada na API em execução. Reinicia o serviço web para carregar o código novo.'
            : response.status === 503
              ? 'O serviço web ou o worker SAP está indisponível. Confirma que ambos estão ligados e volta a tentar.'
              : `A API devolveu HTTP ${response.status}.`;
          throw new Error(detail ? `${routeMsg} Detalhe: ${detail}` : routeMsg);
        }

        appendAuthorizationMessage(
          'assistant',
          `Job CUA_REMOVE criado com sucesso no CUA. Job #${String(data.job_id || '').slice(0, 8)} para remover ${data.roles_count || roles.length} funções do sistema alvo.`
        );
        authorizationRemovalLastContext = {
          user: payload.target_user,
          system: payload.target_system_key,
          rolesCount: data.roles_count || roles.length
        };
        appendAuthorizationMessage(
          'assistant',
          'Vou acompanhar a execução do job e mostrar o resultado final quando terminar.'
        );
        authorizationRemovalJobRequestId++;
        showAuthorizationTypingIndicator(authorizationRemovalJobRequestId, 'A aguardar o fim do job...');
        pollAuthorizationRemovalJob(data.job_id, authorizationRemovalJobRequestId);
        authorizationPendingRemoval = null;
      } catch (error) {
        const errText = String(error?.message || error || '');
        const statusHint = errText.includes('A rota /api/authorizations/remove não foi encontrada')
          ? 'Falta recarregar o serviço web antes de criar o pedido de remoção.'
          : errText.includes('serviço web ou o worker SAP está indisponível')
            ? 'Falta ligar o worker SAP ou o serviço web.'
            : '';
        appendAuthorizationMessage(
          'assistant',
          statusHint
            ? `Não foi possível criar o pedido de remoção: ${statusHint} ${errText}`
            : `Não foi possível criar o pedido de remoção: ${errText}`
        );
      }
    }

    function cancelAuthorizationRemoval() {
      authorizationRemovalJobRequestId++;
      authorizationPendingRemoval = null;
      authorizationRemovalLastContext = null;
      appendAuthorizationMessage('assistant', 'Pedido de remoção cancelado.');
    }

    function removeAuthorizationTypingIndicator() {
      const indicators = document.querySelectorAll('[data-authorization-typing="true"]');
      indicators.forEach((element) => {
        element.remove();
      });
      // Also remove by legacy ID if present
      const legacy = document.getElementById('auth-chat-typing-indicator');
      if (legacy) legacy.remove();
    }

    function removeAuthorizationTypingIndicatorForRequest(requestId) {
      if (requestId === null || requestId === undefined) return;
      const indicators = document.querySelectorAll(`[data-authorization-typing="true"][data-request-id="${requestId}"]`);
      indicators.forEach((element) => {
        element.remove();
      });
    }

    function showAuthorizationTypingIndicator(requestId = null, label = 'A pensar...') {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;
      
      const existing = container.querySelector('[data-authorization-typing="true"]');
      if (existing) {
        const labelSpan = existing.querySelector('.auth-chat-typing-label');
        if (labelSpan) {
          labelSpan.textContent = label;
          if (requestId !== null) {
            existing.dataset.requestId = requestId;
          }
          container.scrollTop = container.scrollHeight;
          return;
        }
      }

      removeAuthorizationTypingIndicator();

      const typingDiv = document.createElement('div');
      typingDiv.className = 'auth-chat-typing';
      typingDiv.dataset.authorizationTyping = 'true';
      if (requestId !== null) {
        typingDiv.dataset.requestId = requestId;
      }
      typingDiv.innerHTML = `
        <span class="auth-chat-typing-label">${escapeAuthorizationText(label)}</span>
        <span class="auth-chat-typing-dot"></span>
        <span class="auth-chat-typing-dot"></span>
        <span class="auth-chat-typing-dot"></span>
      `;
      container.appendChild(typingDiv);
      container.scrollTop = container.scrollHeight;
    }

    function hideAuthorizationTypingIndicator() {
      removeAuthorizationTypingIndicator();
    }

    function appendAuthorizationMessage(sender, text, isHtml = false, options = {}) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();

      const msgDiv = document.createElement('div');
      msgDiv.className = `auth-chat-bubble auth-chat-bubble-${sender}`;
      if (options.key) {
        msgDiv.dataset.messageKey = options.key;
      }
      
      if (sender === 'assistant') {
        const iconSpan = document.createElement('span');
        iconSpan.className = 'auth-chat-bubble-assistant-icon';
        iconSpan.textContent = '🛡️';
        msgDiv.appendChild(iconSpan);

        const contentSpan = document.createElement('span');
        if (isHtml) {
          contentSpan.innerHTML = text;
        } else {
          contentSpan.textContent = text;
        }
        msgDiv.appendChild(contentSpan);
      } else {
        if (isHtml) {
          msgDiv.innerHTML = text;
        } else {
          msgDiv.textContent = text;
        }
      }

      container.appendChild(msgDiv);
      container.scrollTop = container.scrollHeight;
    }

    function appendAuthorizationAssistantMessage(text, options = {}) {
      appendAuthorizationMessage('assistant', text, options.isHtml || false, options);
    }

    function stopAuthorizationJobPolling() {
      authorizationJobRequestId++;
    }

    function clearAuthorizationTimers() {
      if (authorizationLoadingWatchdog) {
        window.clearTimeout(authorizationLoadingWatchdog);
        authorizationLoadingWatchdog = null;
      }
      authorizationChatRequestId++;
      authorizationLoadRequestId++;
    }

    function clearAuthorizationMessages() {
      const container = document.getElementById('authorization-chat-messages');
      if (container) {
        container.innerHTML = '';
      }
    }

    function resetAuthorizationChat() {
      stopAuthorizationJobPolling();
      authorizationRemovalJobRequestId++;
      clearAuthorizationTimers();
      removeAuthorizationTypingIndicator();

      authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
      authorizationLastDisplayedRoles = [];
      authorizationPendingRemoval = null;
      authorizationRemovalLastContext = null;
      window.hasPrintedAgrUsers = false;
      window.hasPrintedAgrTcodes = false;
      window.hasPrintedUsla04 = false;
      window.hasPrintedUsl04 = false;
      window.hasPrintedUsr02 = false;
      window.hasPrintedUsr21 = false;
      window.hasPrintedUsr04 = false;

      clearAuthorizationMessages();
      renderAuthorizationInitialQuestion();
      updateAuthorizationComposer();
    }

    window.resetAuthorizationChatFlow = function() {
      if (authorizationChatInitialized && authorizationAvailableSystems.length > 0) {
        resetAuthorizationChat();
      } else {
        loadAuthorizationChat({ force: true });
      }
    };

    function updateAuthorizationStatus(status, detailText = '') {
      const statusDiv = document.getElementById('auth-chat-status');
      const statusText = document.getElementById('auth-chat-status-text');
      if (!statusDiv || !statusText) return;

      if (status === 'loading') {
        statusDiv.className = "auth-chat-status connecting";
        statusText.textContent = "A carregar...";
      } else if (status === 'ready') {
        statusDiv.className = "auth-chat-status connected";
        statusText.textContent = `Sessão técnica: ${authorizationTechnicalUser || 'Indefinido'}`;
      } else if (status === 'error') {
        statusDiv.className = "auth-chat-status error";
        statusText.textContent = detailText || "Erro";
      }
    }

    function updateAuthorizationStatusBadge() {
      if (authorizationChatState === AUTH_CHAT_STATES.LOADING) {
        updateAuthorizationStatus('loading');
      } else if (authorizationChatState === AUTH_CHAT_STATES.ERROR) {
        updateAuthorizationStatus('error');
      } else {
        updateAuthorizationStatus('ready');
      }
    }

    function renderAuthorizationLoadError(error) {
      hideAuthorizationTypingIndicator();
      const isTimeout = error && (error.name === 'AbortError' || error.message?.includes('aborted'));
      const msg = isTimeout 
        ? 'A configuração SAP demorou demasiado para responder.' 
        : `⚠️ Erro: ${error?.message || 'Erro ao carregar a configuração.'}`;
      appendAuthorizationMessage('assistant', msg);
      appendRetryButton();
    }

    function handleInitialOptionSelect(text) {
      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) {
        inputEl.value = text;
        const formEl = document.getElementById('authorization-chat-form');
        if (formEl) {
          formEl.dispatchEvent(new Event('submit', { cancelable: true, bubbles: true }));
        }
      }
    }

    function renderRoutineSuggestionsForSystem(sys) {
      hideAuthorizationTypingIndicator();
      const sysLabel = sys.label || sys.system || sys.key;

      appendAuthorizationMessage(
        'assistant',
        `Ambiente **${escapeAuthorizationText(sysLabel)}** registado. Em que processo ou rotina SAP pretende trabalhar neste ambiente?\nSelecione uma das sugestões abaixo ou escreva no campo inferior:`
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '6px';
      grid.style.marginBottom = '8px';

      const items = [
        { label: '🛡️ Perfil de autorização', val: 'Perfil de autorização' },
        { label: '📋 Códigos IVA', val: 'Códigos IVA' },
        { label: '🔄 Reverter documento', val: 'Reverter documento' },
        { label: '🏦 Chave de banco', val: 'Chave de banco' },
        { label: '🔍 Cadeias de pesquisa', val: 'Cadeias de pesquisa' }
      ];

      items.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-system-card';
        btn.style.flex = '0 0 auto';
        btn.style.padding = '8px 12px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          handleInitialOptionSelect(item.val);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    async function ensureAuthorizationInitialQuestion() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      if (!container.querySelector('[data-message-key="authorization-initial-question"]')) {
        appendAuthorizationMessage(
          'assistant',
          'Olá! Em que sistema/ambiente SAP pretende trabalhar hoje?',
          false,
          { key: 'authorization-initial-question' }
        );
        showIndividualSystemOptions((sys) => {
          authorizationSelectedSystem = sys;
          const sysLabel = sys.label || sys.system || sys.key;
          appendAuthorizationMessage('user', sysLabel);
          renderRoutineSuggestionsForSystem(sys);
        });
      }
    }

    function resetAuthorizationChat() {
      authorizationTargetUser = '';
      authorizationSelectedSystem = null;
      authorizationSelectedAnalysisType = null;
      authorizationLastStatusData = null;
      authorizationLastDisplayedRoles = [];
      authorizationPendingRemoval = null;
      authorizationRemovalLastContext = null;
      authorizationActiveJobId = null;

      const container = document.getElementById('authorization-chat-messages');
      const input = document.getElementById('authorization-chat-input');
      if (container) container.innerHTML = '';
      if (input) input.value = '';

      removeAuthorizationTypingIndicator();
      setAuthorizationChatState(AUTH_CHAT_STATES.WAITING_USER);
      ensureAuthorizationInitialQuestion();
    }

    function resetAuthorizationChatFlow() {
      resetAuthorizationChat();
    }

    function renderAuthorizationInitialQuestion() {
      resetAuthorizationChat();
    }

    function validateAuthorizationConfig(payload) {
      if (!payload || payload.success !== true) {
        throw new Error(payload?.message || 'Resposta de configuração inválida.');
      }
      if (!Array.isArray(payload.systems)) {
        throw new Error('A lista de sistemas não foi devolvida.');
      }
    }

    function startAuthorizationLoadingWatchdog() {
      window.clearTimeout(authorizationLoadingWatchdog);
      authorizationLoadingWatchdog = window.setTimeout(() => {
        if (authorizationChatState === AUTH_CHAT_STATES.LOADING) {
          authorizationChatInitialized = false;
          authorizationLoadPromise = null;
          setAuthorizationChatState(AUTH_CHAT_STATES.ERROR);
          renderAuthorizationLoadError(new Error('O carregamento da configuração não terminou.'));
        }
      }, 12000);
    }

    async function ensureAuthorizationViewReady(options = {}) {
      const force = options.force === true;
      
      removeAuthorizationTypingIndicator();
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_USER) {
        ensureAuthorizationInitialQuestion();
      }

      if (authorizationLoadPromise && !force) {
        return authorizationLoadPromise;
      }

      if (force) {
        authorizationLoadPromise = null;
        authorizationChatInitialized = false;
        window.clearTimeout(authorizationLoadingWatchdog);
      }

      if (authorizationChatInitialized && authorizationChatState !== AUTH_CHAT_STATES.LOADING && !force) {
        return Promise.resolve();
      }

      authorizationLoadPromise = performAuthorizationInitialization()
        .finally(() => {
          authorizationLoadPromise = null;
        });

      return authorizationLoadPromise;
    }

    async function performAuthorizationInitialization() {
      const requestId = ++authorizationLoadRequestId;
      authorizationChatLoading = true;

      setAuthorizationChatState(AUTH_CHAT_STATES.LOADING);
      renderAuthorizationLoadingState();
      startAuthorizationLoadingWatchdog();
      showAuthorizationTypingIndicator(requestId);

      try {
        console.debug('[AUTH INIT] Inicialização iniciada', { requestId });

        const requiredElements = {
          root: document.getElementById('view-autorizacoes'),
          body: document.getElementById('authorization-chat-messages'),
          input: document.getElementById('authorization-chat-input'),
          sendButton: document.getElementById('authorization-chat-send'),
          statusBadge: document.getElementById('auth-chat-status')
        };

        const missingElements = Object.entries(requiredElements)
          .filter(([, element]) => !element)
          .map(([name]) => name);

        if (missingElements.length > 0) {
          throw new Error(`Elementos do Assistente ausentes: ${missingElements.join(', ')}`);
        }

        const response = await fetchWithTimeout(
          AUTHORIZATION_CONFIG_ENDPOINT,
          {
            method: 'GET',
            cache: 'no-store',
            headers: {
              Accept: 'application/json'
            }
          },
          10000
        );

        console.debug('[AUTH INIT] Resposta recebida', {
          requestId,
          status: response.status
        });

        if (!response.ok) {
          throw new Error(`Configuração respondeu HTTP ${response.status}`);
        }

        const payload = await response.json();

        if (requestId !== authorizationLoadRequestId) {
          removeAuthorizationTypingIndicatorForRequest(requestId);
          return;
        }

        validateAuthorizationConfig(payload);

        authorizationTechnicalUser = String(payload.user_sap || '').trim().toUpperCase();
        authorizationAvailableSystems = Array.isArray(payload.systems) ? payload.systems : [];
        authorizationAnalysisTypes = Array.isArray(payload.analysis_types) ? payload.analysis_types : [];

        if (!payload.env_found) {
          throw new Error(payload.message || 'Não foi possível carregar a configuração SAP.');
        }

        if (!authorizationTechnicalUser) {
          throw new Error(payload.message || 'SAP_USER não está configurado.');
        }

        if (authorizationAvailableSystems.length === 0) {
          throw new Error(payload.message || 'Nenhum sistema SAP foi configurado.');
        }

        window.clearTimeout(authorizationLoadingWatchdog);
        authorizationChatInitialized = true;

        updateAuthorizationStatus('ready');

        authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
        removeAuthorizationTypingIndicator();

        try {
          renderAuthorizationInitialQuestion();
        } catch (error) {
          console.error(
            '[AUTH UI] Falha ao renderizar pergunta inicial',
            error
          );

          removeAuthorizationTypingIndicator();
          setAuthorizationChatState(AUTH_CHAT_STATES.ERROR);
          renderAuthorizationLoadError(error);
          return;
        }

        updateAuthorizationComposer();

        if (requestId !== authorizationLoadRequestId) {
          removeAuthorizationTypingIndicatorForRequest(requestId);
          return;
        }

        console.debug('[AUTH INIT] Inicialização concluída', {
          requestId,
          systems: authorizationAvailableSystems.length
        });

        window.requestAnimationFrame(() => {
          const inputEl = document.getElementById('authorization-chat-input');
          if (inputEl) inputEl.focus();
        });
      } catch (error) {
        if (requestId !== authorizationLoadRequestId) {
          removeAuthorizationTypingIndicatorForRequest(requestId);
          return;
        }

        window.clearTimeout(authorizationLoadingWatchdog);
        authorizationChatInitialized = false;

        console.error('[AUTH INIT] Falha na inicialização', error);

        setAuthorizationChatState(AUTH_CHAT_STATES.ERROR);
        renderAuthorizationLoadError(error);
        updateAuthorizationStatus('error', error?.name === 'AbortError' ? 'Timeout' : 'Erro');
      } finally {
        if (requestId === authorizationLoadRequestId) {
          removeAuthorizationTypingIndicator();
          authorizationChatLoading = false;
          authorizationLoadPromise = null;
        } else {
          removeAuthorizationTypingIndicatorForRequest(requestId);
        }
      }
    }

    window.loadAuthorizationChat = function(options = {}) {
      return ensureAuthorizationViewReady(options);
    };

    window.authorizationDebugState = () => ({
      state: authorizationChatState,
      initialized: authorizationChatInitialized,
      hasLoadPromise: Boolean(authorizationLoadPromise),
      requestId: authorizationLoadRequestId,
      viewVisible: (() => {
        const el = document.getElementById('view-autorizacoes');
        return el ? el.style.display !== 'none' : false;
      })()
    });

    function appendRetryButton() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const btnDiv = document.createElement('div');
      btnDiv.style.marginTop = '10px';
      btnDiv.style.alignSelf = 'flex-start';

      const btn = document.createElement('button');
      btn.className = 'btn btn-secondary';
      btn.style.padding = '8px 16px';
      btn.textContent = '🔄 Tentar novamente';
      btn.onclick = () => ensureAuthorizationViewReady({ force: true });

      btnDiv.appendChild(btn);
      container.appendChild(btnDiv);
      container.scrollTop = container.scrollHeight;
    }

    const AUTH_USER_STOP_WORDS = new Set([
      'RFC', 'CUA', 'DEV', 'QAD', 'PRD', 'S4P', 'S4D', 'S4Q', 'SPA',
      'QUERO', 'DADOS', 'MESTRES', 'MESTRE', 'ANALISAR', 'ANALISE', 'ANÁLISE', 'QUAIS', 'SISTEMAS', 'SISTEMA',
      'ESTA', 'ATRIBUIDO', 'ATRIBUIDA', 'ATRIBUIDOS', 'ATRIBUIDAS', 'NO', 'NA', 'NOS', 'NAS', 'PARA', 'COM',
      'PERFIL', 'PERFIS', 'AUTORIZACAO', 'AUTORIZAÇÃO', 'AUTORIZACOS', 'AUTORIZAÇÕES',
      'CODIGO', 'CODIGOS', 'CÓDIGO', 'CÓDIGOS', 'IVA', 'IMPOSTO', 'IMPOSTOS',
      'REVERTER', 'ESTORNO', 'DOCUMENTO', 'DOCUMENTOS', 'CHAVE', 'BANCO', 'BANCOS',
      'CADEIA', 'CADEIAS', 'PESQUISA', 'PESQUISAS', 'PROCESSO', 'PROCESSOS', 'SKILL', 'SKILLS',
      'AJUDA', 'SOBRE', 'MENU', 'OPCAO', 'OPÇÃO', 'OPCOES', 'OPÇÕES', 'LISTA', 'LISTAR', 'VER', 'OBTER',
      'USER', 'USERS', 'UTILIZADOR', 'UTILIZADORES', 'CONTA', 'ID', 'SESSAO', 'SESSÃO'
    ]);

    function extractAuthorizationParamsFromText(text) {
      const norm = normalizeAuthorizationSearchText(text);
      
      let extractedUser = '';
      const userMatch = text.match(/(?:utilizador|user|id|conta)\s*[:=]?\s*([A-Z0-9\.\-_]{2,30})/i);
      if (userMatch) {
        const val = userMatch[1].toUpperCase();
        if (!AUTH_USER_STOP_WORDS.has(val) && !AUTH_USER_STOP_WORDS.has(normalizeAuthorizationSearchText(val))) {
          extractedUser = val;
        }
      }
      
      if (!extractedUser) {
        const words = text.split(/\s+/);
        const candidate = words.find(w => {
          const u = w.toUpperCase();
          const normWord = normalizeAuthorizationSearchText(w);
          return /^[A-Z][A-Z0-9\.\-_]{2,20}$/i.test(w) && 
                 !AUTH_USER_STOP_WORDS.has(u) && 
                 !AUTH_USER_STOP_WORDS.has(normWord);
        });
        if (candidate) {
          extractedUser = candidate.toUpperCase();
        }
      }

      let extractedSystem = null;
      if (Array.isArray(authorizationAvailableSystems)) {
        const sysMatch = authorizationAvailableSystems.find(sys => {
          const sName = String(sys.system || '').toUpperCase();
          const sKey = String(sys.key || '').toUpperCase();
          return norm.includes(sName) || norm.includes(sKey);
        });
        if (sysMatch) {
          extractedSystem = sysMatch;
        }
      }

      let extractedType = null;
      if (norm.includes('DADOS MESTRES') || norm.includes('MESTRES') || norm.includes('MESTRE')) {
        extractedType = authorizationAnalysisTypes.find(t => t.id === 'master_data' || t.key === 'master_data') || { id: 'master_data', key: 'master_data', label: 'Dados mestres' };
      } else {
        extractedType = authorizationAnalysisTypes.find(t => t.id === 'authorizations' || t.key === 'authorizations') || { id: 'authorizations', key: 'authorizations', label: 'Autorizações' };
      }

      return { extractedUser, extractedSystem, extractedType };
    }

    function handleAuthorizationChatSubmit(event) {
      if (event) event.preventDefault();
      
      const input = document.getElementById('authorization-chat-input');
      if (!input) return;

      const rawVal = input.value.trim();
      if (!rawVal) return;

      input.value = '';
      updateAuthorizationComposer();

      // Exibir a mensagem do utilizador exatamente como foi escrita
      appendAuthorizationMessage('user', rawVal);

      // Se for um comando explícito de limpeza / nova análise
      const normVal = normalizeAuthorizationSearchText(rawVal);
      if (
        normVal.includes('NOVA ANALISE') ||
        normVal.includes('REINICIAR') ||
        normVal.includes('LIMPAR') ||
        normVal.includes('NOVA PESQUISA')
      ) {
        resetAuthorizationChat();
        return;
      }

      // Se estivermos no fluxo de alteração individual à espera do utilizador alvo:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER) {
        const userVal = rawVal.trim().toUpperCase();
        if (authorizationIndividualContext) {
          authorizationIndividualContext.targetUser = userVal;
        }
        authorizationTargetUser = userVal;

        authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_SYSTEM;
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `Registado o utilizador **${escapeAuthorizationText(userVal)}** para a alteração em **${escapeAuthorizationText(authorizationIndividualContext?.processName || 'Processo')}**.\n\nEm que sistema/ambiente pretende efetuar a alteração?`
        );
        showIndividualSystemOptions();
        return;
      }

      // Se estivermos no fluxo de alteração individual à espera dos parâmetros:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS) {
        const detailsVal = rawVal.trim();
        const ctx = authorizationIndividualContext || {};
        const procName = ctx.processName || 'Processo';
        const user = ctx.targetUser || authorizationTargetUser || 'N/D';
        const sysLabel = ctx.selectedSystem?.label || ctx.selectedSystem?.system || authorizationSelectedSystem?.system || 'N/D';

        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `
            <div class="auth-chat-summary">
              <div style="font-weight:700; margin-bottom:8px; color:#10b981; font-size:0.92rem;">✅ Pedido de Alteração Individual Registado</div>
              <div style="display:grid; gap:4px; font-size:0.84rem; margin-bottom:10px;">
                <div><b>• Processo:</b> ${escapeAuthorizationText(procName)}</div>
                <div><b>• Utilizador SAP:</b> ${escapeAuthorizationText(user)}</div>
                <div><b>• Sistema/Ambiente:</b> ${escapeAuthorizationText(sysLabel)}</div>
                <div><b>• Detalhes/Instruções:</b> ${escapeAuthorizationText(detailsVal)}</div>
              </div>
              <div style="font-size:0.8rem; color:var(--text-secondary);">O pedido foi registado no assistente. Para tarefas em lote com múltiplos utilizadores via Excel, utilize a opção <b>Massiva</b>.</div>
            </div>
          `,
          true
        );

        authorizationChatState = AUTH_CHAT_STATES.READY;
        updateAuthorizationComposer();
        return;
      }

      // Se for seleção da opção 'Perfil de autorização' ou consulta sobre Funções PFCG
      if (
        normVal.includes('PERFIL DE AUTORIZACAO') ||
        normVal.includes('PERFIL DE AUTORIZAÇÃO') ||
        normVal === 'PERFIL DE AUTORIZACAO' ||
        normVal === 'PERFIL' ||
        normVal === 'PERFIS' ||
        normVal.includes('FUNCOES PFCG') ||
        normVal.includes('FUNÇÕES PFCG') ||
        normVal === 'PFCG' ||
        normVal === 'CUA'
      ) {
        showPfcgProcessExecutionOptions();
        return;
      }

      // Se for seleção de outro processo inicial
      if (normVal.includes('CODIGOS IVA') || normVal === 'IVA') {
        promptProcessMode('Criar/Manter Códigos IVA (FTXP)', 'Códigos IVA', 'FTXP_CRIAR_CODIGO_IVA.py', 'Automatização FTXP');
        return;
      }

      if (normVal.includes('REVERTER') || normVal.includes('ESTORNO') || normVal.includes('DOCUMENTO')) {
        promptProcessMode('Reverter Documento Contabilístico', 'Reverter Documento', 'REVERTER_DOCUMENTO.py', 'Anulação de documentos FB08/FB05');
        return;
      }

      if (normVal.includes('BANCO') || normVal.includes('CHAVE DE BANCO')) {
        promptProcessMode('Chave de Banco', 'Chave de Banco', 'CHAVE_DE_BANCO.py', 'Criação de chave de banco FI01/FI02');
        return;
      }

      if (normVal.includes('CADEIA') || normVal.includes('CADEIAS DE PESQUISA')) {
        promptProcessMode('Cadeias de Pesquisa', 'Cadeias de Pesquisa', 'CADEIAS_DE_PESQUISA.py', 'Configuração de cadeias OT83');
        return;
      }

      // Se a mensagem contiver a indicação de um utilizador SAP para analisar (ex: "Quero analisar o user S5441")
      const { extractedUser, extractedSystem, extractedType } = extractAuthorizationParamsFromText(rawVal);

      if (extractedUser) {
        authorizationTargetUser = extractedUser;
        if (extractedSystem) authorizationSelectedSystem = extractedSystem;
        if (extractedType) authorizationSelectedAnalysisType = extractedType;

        if (!authorizationSelectedAnalysisType && authorizationAnalysisTypes.length > 0) {
          authorizationSelectedAnalysisType = authorizationAnalysisTypes.find(t => t.id === 'authorizations') || authorizationAnalysisTypes[0];
        }

        if (authorizationSelectedSystem) {
          authorizationChatState = AUTH_CHAT_STATES.LOADING;
          showAuthorizationTypingIndicator();
          setTimeout(() => {
            appendAuthorizationMessage('assistant', `Perfeito. Registado o utilizador ${escapeAuthorizationText(authorizationTargetUser)}.`);
            showAuthorizationSummary();
          }, 400);
          return;
        } else {
          authorizationChatState = AUTH_CHAT_STATES.WAITING_SYSTEM;
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', `Perfeito. Registado o utilizador ${escapeAuthorizationText(authorizationTargetUser)}. Em que sistema pretende fazer a análise?`);
          showAuthorizationSystemOptions();
          return;
        }
      }

      // Se ainda não temos uma análise concluída ou se estamos na fase de recolha de parâmetros:
      if (authorizationChatState !== AUTH_CHAT_STATES.ANALYSIS_COMPLETE && authorizationChatState !== AUTH_CHAT_STATES.READY) {
        if (extractedSystem) authorizationSelectedSystem = extractedSystem;
        if (extractedType) authorizationSelectedAnalysisType = extractedType;

        if (!authorizationSelectedAnalysisType && authorizationAnalysisTypes.length > 0) {
          authorizationSelectedAnalysisType = authorizationAnalysisTypes.find(t => t.id === 'authorizations') || authorizationAnalysisTypes[0];
        }

        if (authorizationTargetUser && authorizationSelectedSystem && authorizationSelectedAnalysisType) {
          authorizationChatState = AUTH_CHAT_STATES.LOADING;
          showAuthorizationTypingIndicator();
          setTimeout(() => {
            appendAuthorizationMessage('assistant', `Perfeito. Registado o utilizador ${escapeAuthorizationText(authorizationTargetUser)}.`);
            showAuthorizationSummary();
          }, 400);
          return;
        }

        if (!authorizationTargetUser) {
          authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', 'Qual o utilizador SAP que pretende analisar?');
          updateAuthorizationComposer();
          return;
        }

        if (!authorizationSelectedSystem) {
          authorizationChatState = AUTH_CHAT_STATES.WAITING_SYSTEM;
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', 'Em que sistema pretende fazer a análise?');
          showAuthorizationSystemOptions();
          return;
        }
      }

      // Se já temos a análise concluída e é uma pergunta/filtragem sobre o resultado atual:
      handleAuthorizationFollowUp(rawVal);
    }

    function showAuthorizationSystemOptions() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.WAITING_SYSTEM;
      updateAuthorizationComposer();

      const grid = document.createElement('div');
      grid.id = 'authorization-system-options';
      grid.className = 'auth-chat-system-grid';

      if (authorizationAvailableSystems.length === 0) {
        const noSystems = document.createElement('div');
        noSystems.style.color = '#ef4444';
        noSystems.style.fontWeight = '600';
        noSystems.textContent = 'Nenhum sistema SAP foi encontrado no .env.';
        container.appendChild(noSystems);
        container.scrollTop = container.scrollHeight;
        return;
      }

      authorizationAvailableSystems.forEach(sys => {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'auth-chat-system-card';
        card.setAttribute('aria-pressed', 'false');
        card.setAttribute('data-key', sys.key);

        const codeSpan = document.createElement('span');
        codeSpan.className = 'sys-code';
        codeSpan.textContent = sys.system;

        const clientSpan = document.createElement('span');
        clientSpan.className = 'sys-client';
        clientSpan.textContent = `Cliente ${sys.client}`;

        card.appendChild(codeSpan);
        card.appendChild(clientSpan);

        if (sys.connection_name) {
          const connSpan = document.createElement('span');
          connSpan.className = 'sys-conn';
          connSpan.textContent = sys.connection_name;
          card.appendChild(connSpan);
        }

        if (sys.execution_mode) {
          const modeSpan = document.createElement('span');
          modeSpan.className = 'sys-conn';
          modeSpan.textContent = sys.execution_mode;
          card.appendChild(modeSpan);
        }

        card.onclick = () => selectAuthorizationSystem(sys.key);
        card.onkeydown = (e) => {
          if (e.key === ' ' || e.key === 'Enter') {
            e.preventDefault();
            selectAuthorizationSystem(sys.key);
          }
        };

        grid.appendChild(card);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function selectAuthorizationSystem(sysKey) {
      if (authorizationChatState !== AUTH_CHAT_STATES.WAITING_SYSTEM) return;

      const sysInfo = authorizationAvailableSystems.find(s => s.key === sysKey);
      if (!sysInfo) return;

      authorizationSelectedSystem = sysInfo;
      
      // Destacar visualmente no grid
      const cards = document.querySelectorAll('.auth-chat-system-card');
      cards.forEach(card => {
        if (card.getAttribute('data-key') === sysKey) {
          card.classList.add('selected');
          card.setAttribute('aria-pressed', 'true');
        } else {
          card.classList.remove('selected');
          card.setAttribute('aria-pressed', 'false');
        }
      });

      // Transição para a próxima etapa: se o utilizador ainda não foi informado, perguntar qual o utilizador
      authorizationChatState = AUTH_CHAT_STATES.LOADING;
      showAuthorizationTypingIndicator();

      setTimeout(() => {
        if (!authorizationTargetUser) {
          authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', 'Qual o utilizador SAP que pretende analisar?');
          updateAuthorizationComposer();
          const inputEl = document.getElementById('authorization-chat-input');
          if (inputEl) {
            inputEl.placeholder = 'Introduza o utilizador SAP (ex: CSILVA)...';
            inputEl.focus();
          }
        } else {
          showAuthorizationSummary();
        }
      }, 450);
    }

    function showAuthorizationAnalysisOptions() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.WAITING_ANALYSIS_TYPE;
      updateAuthorizationComposer();

      const grid = document.createElement('div');
      grid.id = 'authorization-analysis-options';
      grid.className = 'auth-chat-analysis-grid';

      if (authorizationAnalysisTypes.length === 0) {
        const noTypes = document.createElement('div');
        noTypes.style.color = '#ef4444';
        noTypes.style.fontWeight = '600';
        noTypes.textContent = 'Nenhum tipo de análise configurado no backend.';
        container.appendChild(noTypes);
        container.scrollTop = container.scrollHeight;
        return;
      }

      authorizationAnalysisTypes.forEach(t => {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'auth-chat-analysis-card';
        card.setAttribute('aria-pressed', 'false');
        card.setAttribute('data-key', t.key);

        const titleSpan = document.createElement('span');
        titleSpan.className = 'analysis-title';
        const emoji = t.key === 'master_data' ? '👤 ' : '🛡️ ';
        titleSpan.textContent = emoji + t.label;

        const descSpan = document.createElement('span');
        descSpan.className = 'analysis-desc';
        descSpan.textContent = t.description;

        card.appendChild(titleSpan);
        card.appendChild(descSpan);

        card.onclick = () => selectAuthorizationAnalysisType(t.key);
        card.onkeydown = (e) => {
          if (e.key === ' ' || e.key === 'Enter') {
            e.preventDefault();
            selectAuthorizationAnalysisType(t.key);
          }
        };

        grid.appendChild(card);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function selectAuthorizationAnalysisType(typeKey) {
      if (authorizationChatState !== AUTH_CHAT_STATES.WAITING_ANALYSIS_TYPE) return;

      const typeInfo = authorizationAnalysisTypes.find(t => t.key === typeKey);
      if (!typeInfo) return;

      authorizationSelectedAnalysisType = typeInfo;

      // Destacar visualmente no grid
      const cards = document.querySelectorAll('.auth-chat-analysis-card');
      cards.forEach(card => {
        if (card.getAttribute('data-key') === typeKey) {
          card.classList.add('selected');
          card.setAttribute('aria-pressed', 'true');
        } else {
          card.classList.remove('selected');
          card.setAttribute('aria-pressed', 'false');
        }
      });

      // TransiÃ§Ã£o para a prÃ³xima etapa: perguntar o sistema alvo
      authorizationChatState = AUTH_CHAT_STATES.LOADING;
      showAuthorizationTypingIndicator();

      setTimeout(() => {
        appendAuthorizationMessage('assistant', 'Em que sistema pretende fazer a análise?');
        showAuthorizationSystemOptions();
      }, 500);
    }

    function showAuthorizationWorkerOfflineMessage() {
      // Check if message already exists to avoid duplicates
      if (document.getElementById('authorization-worker-offline')) {
        return;
      }
      
      const warningHtml = `
        <div id="authorization-worker-offline" style="border-left: 4px solid var(--warning, #ff9800); padding-left: 10px; margin: 5px 0;">
          <div style="font-weight: bold; color: var(--warning, #ff9800); margin-bottom: 5px;">O Worker Windows está desligado.</div>
          <div style="margin-bottom: 10px;">Para abrir o SAP CUA, ligue primeiro o Worker através do botão “Ligar Worker” no canto inferior esquerdo.</div>
          <div style="margin-bottom: 12px;">Quando o Worker estiver online, volte a clicar em “Iniciar análise”.</div>
          <button type="button" id="chat-start-worker-btn" style="display: inline-flex; align-items: center; gap: 5px; cursor: pointer; padding: 6px 12px; border: none; border-radius: 4px; background-color: var(--warning, #ff9800); color: #fff; font-weight: bold; font-family: inherit;">
            🖥️ Ligar Worker
          </button>
        </div>
      `;
      appendAuthorizationMessage('assistant', warningHtml, true);
      
      // Bind click event to the button
      const chatBtn = document.getElementById('chat-start-worker-btn');
      if (chatBtn) {
        chatBtn.addEventListener('click', () => {
          const sidebarBtn = document.getElementById('start-worker-btn');
          if (sidebarBtn) {
            sidebarBtn.click();
            // Highlight sidebar btn
            sidebarBtn.style.animation = 'pulse-attention 1s infinite alternate';
            setTimeout(() => {
              sidebarBtn.style.animation = '';
            }, 5000);
          } else {
            window.location.href = 'sap-worker://start';
          }
        });
      }
    }

    async function startAuthorizationAnalysis() {
      if (
        authorizationChatState === AUTH_CHAT_STATES.ANALYZING ||
        authorizationActiveJobId ||
        window.isCheckingWorkerOnline
      ) {
        return;
      }

      if (!authorizationTargetUser || !authorizationSelectedSystem || !authorizationSelectedAnalysisType) {
        return;
      }

      const btnStart = document.querySelector('.auth-chat-summary-actions .btn-primary');
      const btnChange = document.querySelector('.auth-chat-summary-actions .btn-secondary');

      // 1. Pre-validar se o worker está online
      window.isCheckingWorkerOnline = true;
      if (btnStart) {
        btnStart.disabled = true;
        btnStart.textContent = 'A verificar worker...';
      }
      if (btnChange) {
        btnChange.disabled = true;
      }

      const isOnline = await requireOnlineWorker({
        context: 'start_analysis',
        onOffline: () => {
          showAuthorizationWorkerOfflineMessage();
        }
      });

      window.isCheckingWorkerOnline = false;

      if (!isOnline) {
        if (btnStart) {
          btnStart.disabled = false;
          btnStart.textContent = 'Worker desligado';
          setTimeout(() => {
            if (btnStart.textContent === 'Worker desligado') {
              btnStart.textContent = 'Iniciar análise';
            }
          }, 3000);
        }
        if (btnChange) {
          btnChange.disabled = false;
        }
        return;
      }

      // 2. Se estiver online, prosseguir normalmente
      const currentRequestId = ++authorizationJobRequestId;
      authorizationChatState = AUTH_CHAT_STATES.ANALYZING;
      updateAuthorizationComposer();

      // Disable change button and change start button text
      if (btnStart) {
        btnStart.disabled = true;
        btnStart.textContent = 'A analisar...';
      }
      if (btnChange) {
        btnChange.disabled = true;
      }

      // Mostrar mensagem do assistente e indicador de escrita
      const executionMode = getAuthorizationExecutionMode();
      const isDevFlow = isAuthorizationDevFlow();
      const isRfcFlow = executionMode === 'RFC' || isDevFlow;
      const selectedSystemLabel = authorizationSelectedSystem?.system || 'sistema selecionado';
      appendAuthorizationMessage(
        'assistant',
        isRfcFlow
          ? `A iniciar a ligação RFC ao sistema ${escapeAuthorizationText(selectedSystemLabel)}...`
          : `A iniciar a análise no sistema ${escapeAuthorizationText(selectedSystemLabel)}...`
      );
      showAuthorizationTypingIndicator();

      const payload = {
        target_user: authorizationTargetUser,
        target_system_key: authorizationSelectedSystem.key,
        subsystem_filter: authorizationSelectedSystem.key,
        analysis_type: authorizationSelectedAnalysisType.key
      };

      try {
        const response = await fetch('/api/authorizations/start', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(payload)
        });

        if (currentRequestId !== authorizationJobRequestId) {
          return;
        }

        if (!response.ok) {
          const errData = await response.json().catch(() => ({}));
          // Tratar especificamente HTTP 503 com worker_offline
          if (response.status === 503 && errData.detail && errData.detail.code === 'worker_offline') {
            throw { isWorkerOffline: true, message: errData.detail.message };
          }
          throw new Error(errData.detail || `Erro HTTP ${response.status}`);
        }

        const data = await response.json();
        authorizationActiveJobId = data.job_id;
        
        // Iniciar polling
        pollAuthorizationJob(data.job_id, currentRequestId);

      } catch (err) {
        if (currentRequestId !== authorizationJobRequestId) {
          return;
        }
        hideAuthorizationTypingIndicator();
        authorizationChatState = AUTH_CHAT_STATES.READY;
        authorizationActiveJobId = null;
        updateAuthorizationComposer();

        // Restore buttons
        if (btnStart) {
          btnStart.disabled = false;
          btnStart.textContent = 'Iniciar análise';
        }
        if (btnChange) {
          btnChange.disabled = false;
        }

        if (err && err.isWorkerOffline) {
          showAuthorizationWorkerOfflineMessage();
        } else {
          appendAuthorizationMessage(
            'assistant',
            '⚠️ Não foi possível iniciar a análise de autorizações. Verifique o worker e a configuração da ligação.'
          );
        }
      }
    }

    function pollAuthorizationJob(jobId, currentRequestId) {
      const startTime = Date.now();
      const timeoutMs = 120000;
      const executionMode = getAuthorizationExecutionMode();
      const isDevFlow = isAuthorizationDevFlow();
      const isMasterDataFlow = authorizationSelectedAnalysisType && authorizationSelectedAnalysisType.key === 'master_data';
      const isRfcFlow = executionMode === 'RFC' || isDevFlow;
      let hasPrintedAuthorizationReady = false;
      window.hasPrintedAgrUsers = false;
      window.hasPrintedAgrTcodes = false;
      window.hasPrintedUsla04 = false;
      window.hasPrintedUsl04 = false;
      window.hasPrintedUsr02 = false;
      window.hasPrintedUsr21 = false;
      window.hasPrintedUsr04 = false;

      async function check() {
        if (currentRequestId !== authorizationJobRequestId) {
          return;
        }

        try {
          const response = await fetch(`/api/jobs/${jobId}`);
          if (!response.ok) {
            throw new Error(`Erro HTTP ${response.status}`);
          }
          const job = await response.json();

          if (currentRequestId !== authorizationJobRequestId) {
            return;
          }

          const executionMode = getAuthorizationExecutionMode();

          // Escutar os logs para atualizar dinamicamente a mensagem no mesmo balão de pensamento
          if (job.log) {
            if (
              !hasPrintedAuthorizationReady &&
              (
                job.log.includes("Sessão CUA validada") ||
                job.log.includes("Sessão SAP CUA aberta") ||
                job.log.includes("Ligação RFC estabelecida") ||
                job.log.includes("AGR_USERS") ||
                job.log.includes("USR02")
              )
            ) {
              hasPrintedAuthorizationReady = true;
              showAuthorizationTypingIndicator(
                currentRequestId,
                isRfcFlow
                  ? (isMasterDataFlow
                    ? 'Ligação RFC pronta. A consultar dados mestres...'
                    : (isDevFlow
                      ? 'Ligação RFC pronta. A consultar atribuições em AGR_USERS...'
                      : 'Ligação RFC pronta. A consultar USZBVSYS...'))
                  : (isMasterDataFlow
                    ? 'Ligação pronta. A consultar dados mestres...'
                    : (isDevFlow
                      ? 'Ligação pronta. A consultar atribuições em AGR_USERS...'
                      : 'Ligação pronta. A consultar USZBVSYS...'))
              );
            }
            if (
              isMasterDataFlow &&
              (job.log.includes("USR02") || job.log.includes("Tabela USR02 informada")) &&
              !job.log.includes("Dados de validade e bloqueio lidos")
            ) {
              if (!window.hasPrintedUsr02) {
                window.hasPrintedUsr02 = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  executionMode === 'RFC' ? 'A consultar USR02 via RFC...' : 'A consultar USR02...'
                );
              }
            }
            if (
              isMasterDataFlow &&
              (job.log.includes("USR21") || job.log.includes("Tabela USR21 informada")) &&
              !job.log.includes("Ligação do utilizador ao endereço lida")
            ) {
              if (!window.hasPrintedUsr21) {
                window.hasPrintedUsr21 = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  executionMode === 'RFC' ? 'A consultar ligação do utilizador ao endereço via RFC...' : 'A consultar ligação do utilizador ao endereço...'
                );
              }
            }
            if (
              isMasterDataFlow &&
              (job.log.includes("USR04") || job.log.includes("Tabela USR04 informada")) &&
              !job.log.includes("Perfis do utilizador lidos")
            ) {
              if (!window.hasPrintedUsr04) {
                window.hasPrintedUsr04 = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  executionMode === 'RFC' ? 'A consultar perfis em USR04 via RFC...' : 'A consultar perfis em USR04...'
                );
              }
            }
            if (
              (job.log.includes("AGR_USERS") || job.log.includes("Tabela AGR_USERS informada")) &&
              !job.log.includes("Roles do utilizador lidas") &&
              !job.log.includes("Roles lidas via AGR_USERS")
            ) {
              if (!window.hasPrintedAgrUsers) {
                window.hasPrintedAgrUsers = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  isRfcFlow
                    ? (isMasterDataFlow
                      ? 'A consultar roles em AGR_USERS via RFC...'
                      : 'Utilizador validado. A consultar atribuições em AGR_USERS via RFC...')
                    : (isMasterDataFlow
                      ? 'A consultar roles em AGR_USERS...'
                      : 'Utilizador validado. A consultar atribuições em AGR_USERS...')
                );
              }
            }
            if (
              (job.log.includes("AGR_TCODES") || job.log.includes("Tabela AGR_TCODES informada")) &&
              !job.log.includes("Funções AGR_TCODES lidas")
            ) {
              if (!window.hasPrintedAgrTcodes) {
                window.hasPrintedAgrTcodes = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  executionMode === 'RFC' ? 'A consultar transações em AGR_TCODES via RFC...' : 'A consultar transações em AGR_TCODES...'
                );
              }
            }
            if (
              (job.log.includes("A abrir SE16 para USLA04") || job.log.includes("A abrir SE16N para USLA04") || job.log.includes("Tabela USLA04 informada")) &&
              !job.log.includes("USLA04 executada")
            ) {
              if (!window.hasPrintedUsla04) {
                window.hasPrintedUsla04 = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  executionMode === 'RFC' ? 'Utilizador validado. A consultar funções em USLA04 via RFC...' : 'Utilizador validado. A consultar funções em USLA04...'
                );
              }
            }
            if (
              (job.log.includes("A abrir SE16 para USL04") || job.log.includes("A abrir SE16N para USL04") || job.log.includes("Tabela USL04 informada")) &&
              !job.log.includes("USL04 executada")
            ) {
              if (!window.hasPrintedUsl04) {
                window.hasPrintedUsl04 = true;
                showAuthorizationTypingIndicator(
                  currentRequestId,
                  executionMode === 'RFC' ? 'A consultar perfis em USL04 via RFC...' : 'A consultar perfis em USL04...'
                );
              }
            }
          }

          if (job.state === 'succeeded' || job.state === 'succeeded_with_warnings') {
            let statusData = {};
            try {
              statusData = JSON.parse(job.status);
            } catch (e) {
              throw new Error('Status do job não é um JSON válido.');
            }

            hideAuthorizationTypingIndicator();
            
            // Validar identificação de versão do Worker
            if (statusData.worker_feature_version !== "authorization-tables-v1") {
              throw new Error('A análise foi iniciada, mas a consulta das tabelas não foi executada pelo Worker (versão antiga detetada). Reinicie o Worker.');
            }

            // Validar payload estruturado
            if (
              statusData.success !== true ||
              statusData.data_source_verified !== true ||
              !Array.isArray(statusData.queries) ||
              statusData.queries.length === 0
            ) {
              throw new Error(
                executionMode === 'RFC'
                  ? 'A ligação RFC foi estabelecida, mas a consulta das tabelas não foi executada.'
                  : 'A análise foi iniciada, mas a consulta das tabelas não foi executada.'
              );
            }

            // Mapear queries executadas
            const queryMap = {};
            statusData.queries.forEach(q => {
              queryMap[q.table] = q;
            });

            const isMasterData = statusData.analysis_type === 'master_data';
            const isCuaFlow = Boolean(queryMap["USLA04"]);
            const isDevFlow = Boolean(queryMap["AGR_USERS"]);

            if (isMasterData) {
              const requiredMasterTables = ["USR02", "USR21", "USR04", "AGR_USERS"];
              requiredMasterTables.forEach(table => {
                const q = queryMap[table];
                if (!q || q.executed !== true || q.filters_applied !== true) {
                  throw new Error(`A consulta na tabela ${table} não foi realizada ou os filtros não foram aplicados.`);
                }
              });
            } else if (isCuaFlow) {
              const qRoles = queryMap["USLA04"];
              if (!qRoles || qRoles.executed !== true || qRoles.filters_applied !== true) {
                throw new Error('A consulta na tabela USLA04 não foi realizada ou os filtros não foram aplicados.');
              }
            } else if (isDevFlow) {
              const qUsers = queryMap["AGR_USERS"];
              if (!qUsers || qUsers.executed !== true || qUsers.filters_applied !== true) {
                throw new Error('A consulta na tabela AGR_USERS não foi realizada ou os filtros não foram aplicados.');
              }
              const qTcodes = queryMap["AGR_TCODES"];
              if (!qTcodes || qTcodes.executed !== true || qTcodes.filters_applied !== true) {
                throw new Error('A consulta na tabela AGR_TCODES não foi realizada ou os filtros não foram aplicados.');
              }
            }

            authorizationChatState = AUTH_CHAT_STATES.ANALYSIS_COMPLETE;
            authorizationActiveJobId = null;
            authorizationLastStatusData = statusData;
            updateAuthorizationComposer();

              const btnStart = document.querySelector('.auth-chat-summary-actions .btn-primary');
              if (btnStart) {
                btnStart.textContent = 'Análise concluída';
                btnStart.disabled = true;
              }

              if (statusData.code === 'user_not_assigned_to_system') {
                const sysShort = authorizationSelectedSystem ? authorizationSelectedSystem.system : statusData.target_system.system;
                appendAuthorizationMessage(
                  'assistant',
                  executionMode === 'RFC'
                    ? `O utilizador ${escapeAuthorizationText(statusData.target_user)} não está associado ao sistema ${escapeAuthorizationText(sysShort)} via RFC.`
                    : `O utilizador ${escapeAuthorizationText(statusData.target_user)} não está associado ao sistema ${escapeAuthorizationText(sysShort)}.`
                );
                
              } else if (statusData.code === 'analysis_complete') {
                appendAuthorizationMessage(
                  'assistant',
                  executionMode === 'RFC'
                    ? 'Análise de autorizações concluída via RFC.'
                    : 'Análise de autorizações concluída.'
                );

                const s = statusData.summary;
                const sysShort = authorizationSelectedSystem ? authorizationSelectedSystem.system : statusData.target_system.system;
                const functions = Array.isArray(statusData.functions) ? statusData.functions : [];
                const functionCodes = functions
                  .map(f => String(f?.tcode || '').trim())
                  .filter(Boolean);
                const isMasterData = statusData.analysis_type === 'master_data';
                const masterData = statusData.master_data || {};
                const masterDataFullName = masterData.full_name || masterData.name_text || [masterData.name_first, masterData.name_last].filter(Boolean).join(' ').trim();
                
                let rowsHtml = '';
                const roles = statusData.roles || [];
                authorizationLastDisplayedRoles = Array.isArray(roles) ? roles.slice() : [];
                
                roles.forEach((r, idx) => {
                  const validityFrom = r.valid_from || '(aberta)';
                  const validityTo = r.valid_to || '(aberta)';
                  
                  let statusBadge = '';
                  if (r.validity_status === 'active') {
                    statusBadge = '<span class="badge badge-success">Ativa</span>';
                  } else if (r.validity_status === 'expired') {
                    statusBadge = '<span class="badge badge-danger">Expirada</span>';
                  } else if (r.validity_status === 'future') {
                    statusBadge = '<span class="badge badge-warning">Futura</span>';
                  } else {
                    statusBadge = '<span class="badge">Indeterminada</span>';
                  }
                  
                  const isHidden = idx >= 15 ? 'class="auth-row-hidden"' : '';
                  
                  rowsHtml += `
                    <tr ${isHidden}>
                      <td><strong>${escapeAuthorizationText(r.role)}</strong></td>
                      <td>${validityFrom}</td>
                      <td>${validityTo}</td>
                      <td>${statusBadge}</td>
                      <td>${escapeAuthorizationText(r.assignment_origin_label || r.assignment_origin)}</td>
                    </tr>
                  `;
                });

                if (roles.length === 0) {
                  rowsHtml = `<tr><td colspan="5" style="text-align: center; color: var(--text-muted);">Nenhuma role encontrada.</td></tr>`;
                }

                let verTodasBtn = '';
                if (roles.length > 15) {
                  verTodasBtn = `
                    <button class="auth-ver-todas-btn" onclick="
                      const table = this.closest('.auth-table-wrapper');
                      table.querySelectorAll('.auth-row-hidden').forEach(row => row.classList.remove('auth-row-hidden'));
                      this.style.display = 'none';
                    ">Ver todas (${roles.length - 15} mais)</button>
                  `;
                }

                const resultHtml = `
                  <div class="auth-analysis-result">
                    <div class="auth-summary-card">
                      <div class="auth-summary-header">${isMasterData ? 'Resumo dos Dados Mestres' : 'Resumo do Utilizador'}: ${escapeAuthorizationText(statusData.target_user)}</div>
                      <div class="auth-summary-grid">
                        <div class="auth-summary-item"><strong>Utilizador:</strong> <span>${escapeAuthorizationText(statusData.target_user)}</span></div>
                        <div class="auth-summary-item"><strong>${isMasterData ? 'Nome completo' : 'Sistema'}:</strong> <span>${escapeAuthorizationText(isMasterData ? (masterDataFullName || 'N/D') : sysShort)}</span></div>
                        <div class="auth-summary-item"><strong>${isMasterData ? 'Email' : 'Funções encontradas'}:</strong> <span>${escapeAuthorizationText(isMasterData ? (masterData.email || 'N/D') : String(s.total_roles ?? 0))}</span></div>
                        <div class="auth-summary-item"><strong>${isMasterData ? 'Grupo' : 'Ativas'}:</strong> <span class="${isMasterData ? '' : 'badge badge-success'}">${escapeAuthorizationText(isMasterData ? (masterData.user_group || 'N/D') : String(s.active_roles ?? 0))}</span></div>
                        <div class="auth-summary-item"><strong>${isMasterData ? 'Validade' : 'Expiradas'}:</strong> <span class="${isMasterData ? '' : 'badge badge-danger'}">${escapeAuthorizationText(isMasterData ? `${masterData.valid_from || '(aberta)'} - ${masterData.valid_to || '(aberta)'}` : String(s.expired_roles ?? 0))}</span></div>
                        <div class="auth-summary-item"><strong>${isMasterData ? 'Bloqueio' : 'Futuras'}:</strong> <span class="${isMasterData ? '' : 'badge badge-warning'}">${escapeAuthorizationText(isMasterData ? (masterData.lock_status || 'N/D') : String(s.future_roles ?? 0))}</span></div>
                        <div class="auth-summary-item"><strong>Perfis do Utilizador:</strong> <span>${String(s.total_profiles ?? 0)}</span></div>
                        <div class="auth-summary-item"><strong>${isMasterData ? 'Roles' : 'Transações'}:</strong> <span>${isMasterData ? String(s.total_roles ?? 0) : String(functionCodes.length)}</span></div>
                      </div>
                    </div>

                    <div class="auth-table-wrapper">
                      <table class="auth-result-table">
                        <thead>
                          <tr>
                            <th>Role</th>
                            <th>Validade inicial</th>
                            <th>Validade final</th>
                            <th>Estado</th>
                            <th>Origem</th>
                          </tr>
                        </thead>
                        <tbody>
                          ${rowsHtml}
                        </tbody>
                      </table>
                      ${verTodasBtn}
                    </div>
                  </div>
                `;

                appendAuthorizationMessage('assistant', resultHtml, true);
              }
          } else if (job.state === 'failed' || job.state === 'cancelled') {
            let errorMsg = job.status || 'O job falhou ou foi cancelado.';
            try {
              const statusData = JSON.parse(job.status);
              if (statusData && statusData.message) {
                errorMsg = statusData.message;
              }
            } catch (e) {}
            throw new Error(errorMsg);
          } else {
            // Pending or running
            if (!document.querySelector('[data-authorization-typing="true"]')) {
              showAuthorizationTypingIndicator(currentRequestId, job.state === 'running' ? 'A processar...' : 'A aguardar execução...');
            }

            if (Date.now() - startTime > timeoutMs) {
              throw new Error('Tempo limite da análise esgotado.');
            }

            // Verificar se o worker ficou offline durante a execução do job
            const statusRes = await fetch('/api/worker/status', { cache: 'no-store' });
            let workerOnline = true;
            if (statusRes.ok) {
              const statusData = await statusRes.json();
              updateSidebarWorkerStatus(statusData);
              workerOnline = statusData.online === true;
            } else {
              const offlineData = { online: false, state: 'offline', status: 'offline' };
              updateSidebarWorkerStatus(offlineData);
              workerOnline = false;
            }

            if (!workerOnline) {
              throw { isWorkerOfflineDuringExecution: true };
            }

            setTimeout(check, 1200);
          }

        } catch (err) {
          if (currentRequestId !== authorizationJobRequestId) {
            return;
          }
          hideAuthorizationTypingIndicator();
          authorizationChatState = AUTH_CHAT_STATES.READY;
          authorizationActiveJobId = null;
          updateAuthorizationComposer();

          const btnStart = document.querySelector('.auth-chat-summary-actions .btn-primary');
          const btnChange = document.querySelector('.auth-chat-summary-actions .btn-secondary');
          if (btnStart) {
            btnStart.disabled = false;
            btnStart.textContent = 'Iniciar análise';
          }
          if (btnChange) {
            btnChange.disabled = false;
          }

          if (err && err.isWorkerOfflineDuringExecution) {
            showAuthorizationWorkerOfflineMessage();
          } else {
            let userFriendlyMsg = executionMode === 'RFC'
              ? '⚠️ A ligação RFC foi estabelecida, mas a consulta das tabelas não foi executada. Reinicie o Worker e tente novamente.'
              : '⚠️ A análise foi iniciada, mas a consulta das tabelas não foi executada. Reinicie o Worker e tente novamente.';
            const errText = err.message || '';
            if (errText.includes('table_not_authorized')) {
              userFriendlyMsg = '⚠️ O utilizador técnico não possui autorização para consultar a tabela necessária.';
            } else if (errText.includes('filter_not_applied')) {
              userFriendlyMsg = '⚠️ Não foi possível aplicar de forma segura os filtros da análise.';
            } else if (errText.includes('user_not_assigned_to_system')) {
              const sysShort = authorizationSelectedSystem ? authorizationSelectedSystem.system : 'alvo';
              userFriendlyMsg = `⚠️ O utilizador ${escapeAuthorizationText(authorizationTargetUser)} não está associado ao sistema ${escapeAuthorizationText(sysShort)}.`;
            } else if (errText.includes('Não foi possível abrir a sessão SAP CUA')) {
              userFriendlyMsg = executionMode === 'RFC'
                ? '⚠️ Não foi possível abrir a ligação RFC ao sistema selecionado. Verifique as credenciais RFC e o acesso ao SAP.'
                : '⚠️ Não foi possível abrir a sessão SAP. Verifique se o SAP GUI está configurado e com Scripting ativo.';
            } else if (errText.includes('rfc_connection_failed')) {
              userFriendlyMsg = '⚠️ Não foi possível abrir a ligação RFC ao sistema selecionado. Verifique as credenciais RFC e o acesso ao SAP.';
            } else if (errText.includes('rfc_table_read_failed')) {
              userFriendlyMsg = '⚠️ Não foi possível consultar uma das tabelas SAP via RFC.';
            } else if (errText.includes('versão antiga')) {
              userFriendlyMsg = '⚠️ ' + errText;
            }
            appendAuthorizationMessage('assistant', `${userFriendlyMsg}\n*(Detalhe: ${errText || 'Erro desconhecido'})*`);

            const actionsHtml = `
              <div class="auth-actions-row">
                <button class="btn btn-secondary btn-sm" onclick="resetAuthorizationChatFlow()">Nova análise</button>
              </div>
            `;
            appendAuthorizationMessage('assistant', actionsHtml, true);
          }
        }
      }

      setTimeout(check, 1200);
    }

    function showAuthorizationSummary() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.READY;
      updateAuthorizationComposer();

      const summaryDiv = document.createElement('div');
      summaryDiv.className = 'auth-chat-summary';

      const fields = [
        { label: 'Utilizador analisado', val: authorizationTargetUser },
        { label: 'Tipo de análise', val: authorizationSelectedAnalysisType ? authorizationSelectedAnalysisType.label : 'N/D' },
        { label: 'Modo de execução', val: getAuthorizationExecutionMode() || 'N/D' },
        { label: 'Sistema', val: authorizationSelectedSystem ? authorizationSelectedSystem.system : 'N/D' }
      ];

      fields.forEach(f => {
        const row = document.createElement('div');
        row.className = 'auth-chat-summary-row';

        const lSpan = document.createElement('span');
        lSpan.className = 'auth-chat-summary-label';
        lSpan.textContent = f.label;

        const vSpan = document.createElement('span');
        vSpan.className = 'auth-chat-summary-value';
        vSpan.textContent = f.val;

        row.appendChild(lSpan);
        row.appendChild(vSpan);
        summaryDiv.appendChild(row);
      });

      container.appendChild(summaryDiv);
      container.scrollTop = container.scrollHeight;

      // Executar a análise automaticamente sem necessidade de confirmação manual
      setTimeout(() => {
        startAuthorizationAnalysis();
      }, 300);
    }

    function handleAuthorizationChatInput(event) {
      updateAuthorizationComposer();
    }

    let authorizationChatEventsInitialized = false;

    function initializeAuthorizationChatEvents() {
      if (authorizationChatEventsInitialized) {
        return;
      }

      const form = document.getElementById('authorization-chat-form');
      const input = document.getElementById('authorization-chat-input');

      if (!form || !input) {
        return;
      }

      form.addEventListener('submit', handleAuthorizationChatSubmit);
      input.addEventListener('input', handleAuthorizationChatInput);

      authorizationChatEventsInitialized = true;
    }

    document.addEventListener('DOMContentLoaded', () => {
      initializeAuthorizationChatEvents();
      window.setTimeout(() => {
        const authorizationView = document.getElementById('view-autorizacoes');
        const isVisible = authorizationView && (authorizationView.style.display !== 'none');
        if (isVisible) {
          ensureAuthorizationViewReady({ force: true });
        }
      }, 0);
    });

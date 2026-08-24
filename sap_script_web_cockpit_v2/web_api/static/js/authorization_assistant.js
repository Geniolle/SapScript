const AUTH_CHAT_STATES = {
      LOADING: 'loading',
      WAITING_INITIAL_CHOICE: 'waiting_initial_choice',
      WAITING_JIRA_TEAM: 'waiting_jira_team',
      WAITING_JIRA_ASSIGNEE: 'waiting_jira_assignee',
      WAITING_JIRA_FILTER_MODE: 'waiting_jira_filter_mode',
      WAITING_JIRA_PROCESS: 'waiting_jira_process',
      WAITING_JIRA_TICKET: 'waiting_jira_ticket',
      WAITING_TICKET_ACTION: 'waiting_ticket_action',
      WAITING_USER: 'waiting_user',
      WAITING_SYSTEM: 'waiting_system',
      WAITING_ANALYSIS_TYPE: 'waiting_analysis_type',
      WAITING_INDIVIDUAL_USER: 'waiting_individual_user',
      WAITING_HR_SEARCH_QUERY: 'waiting_hr_search_query',
      WAITING_COPY_REFERENCE_USER: 'waiting_copy_reference_user',
      WAITING_CUA_USER_DETAILS: 'waiting_cua_user_details',
      WAITING_CUA_FUNCTION: 'waiting_cua_function',
      WAITING_CUA_DEPARTMENT: 'waiting_cua_department',
      WAITING_CUA_MOB_NUMBER: 'waiting_cua_mob_number',
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
    let authorizationTargetUserDisplayName = '';
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

    let authorizationSelectedJiraTeam = null;
    let authorizationSelectedJiraAssignee = null;
    let authorizationSelectedJiraProcess = null;
    let authorizationSelectedJiraTicket = null;
    let authorizationCachedJiraTickets = null;
    let authorizationUatCreateDocumentFlow = null;
    let authorizationUatCreateDocumentJobRequestId = 0;
    let authorizationUatLastCreatedDocumentContext = null;
    let authorizationUatExecuteF110JobRequestId = 0;

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
        .replace(/&/g, "&")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
    }

    function formatSapUserId(rawUser) {
      if (!rawUser) return '';
      let clean = String(rawUser).trim().toUpperCase();
      if (!clean) return '';
      if (clean.startsWith('S')) return clean;
      if (/^\d+$/.test(clean)) {
        clean = clean.replace(/^0+/, '');
        return 'S' + clean;
      }
      return clean;
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
        authorizationChatState === AUTH_CHAT_STATES.WAITING_INITIAL_CHOICE ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_TEAM ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_ASSIGNEE ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_FILTER_MODE ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_PROCESS ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_TICKET ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_TICKET_ACTION ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_USER ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_COPY_REFERENCE_USER ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_USER_DETAILS ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_FUNCTION ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_DEPARTMENT ||
        authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_MOB_NUMBER;
      const followUpReady =
        authorizationChatState === AUTH_CHAT_STATES.READY ||
        authorizationChatState === AUTH_CHAT_STATES.ANALYSIS_COMPLETE;
      const executionMode = getAuthorizationExecutionMode();

      input.disabled = !(waitingForUser || followUpReady);

      if (authorizationChatState === AUTH_CHAT_STATES.LOADING) {
        input.placeholder = 'A carregar configuração SAP...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INITIAL_CHOICE) {
        input.placeholder = 'Selecione "Ticket" ou "Processo" ou escreva no campo...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_TEAM) {
        input.placeholder = 'Selecione uma equipa Jira acima ou escreva o nome...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_ASSIGNEE) {
        input.placeholder = 'Selecione um responsável acima ou escreva o nome...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_FILTER_MODE) {
        input.placeholder = 'Selecione "Todos os tickets" ou "Filtrar por processo"...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_PROCESS) {
        input.placeholder = 'Selecione um processo acima ou escreva o nome...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_TICKET) {
        input.placeholder = 'Selecione um ticket acima ou escreva a chave (ex: SD-1234)...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_TICKET_ACTION) {
        input.placeholder = 'Selecione uma ação para o ticket acima...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_USER) {
        input.placeholder = 'Escreva a sua mensagem ou utilizador SAP...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER) {
        input.placeholder = 'Escreva o utilizador SAP alvo (ex: CSILVA)...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY) {
        input.placeholder = 'Escreva o PERNR, Nome ou Utilizador para pesquisa no RH...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_COPY_REFERENCE_USER) {
        input.placeholder = 'Escreva o utilizador SAP de referência (ex: JSILVA)...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_FUNCTION) {
        input.placeholder = 'Escreva a Função (FUNCTION) do utilizador no CUA...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_DEPARTMENT) {
        input.placeholder = 'Escreva o Departamento (DEPARTMENT) do utilizador no CUA...';
      } else if (authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_MOB_NUMBER) {
        input.placeholder = 'Escreva o Telefone (MOB_NUMBER) do utilizador no CUA...';
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
      const stateLabel = state === 'succeeded_with_warnings' ? 'concluída com avisos' : state === 'succeeded' ? 'concluída com sucesso' : state || 'terminada';
      const jobIdShort = String(job?.id || '').slice(0, 8);
      const pluralize = (value, singular, plural) => `${value} ${value === 1 ? singular : plural}`;
      const counts = [];

      const script = String(job?.params?.subprocesso || job?.params?.processo || '').toUpperCase();
      const isEndDate = script.includes('ENDDATE');
      const actionName = isEndDate ? 'Alteração de validade' : 'Remoção';

      if (summary.processed !== null) counts.push(pluralize(summary.processed, 'linha processada', 'linhas processadas'));
      if (summary.concluded !== null) counts.push(pluralize(summary.concluded, 'concluída', 'concluídas'));
      if (summary.warnings !== null) counts.push(pluralize(summary.warnings, 'aviso', 'avisos'));
      if (summary.errors !== null) counts.push(pluralize(summary.errors, 'erro', 'erros'));
      if (summary.removed !== null) counts.push(pluralize(summary.removed, isEndDate ? 'função alterada' : 'função eliminada', isEndDate ? 'funções alteradas' : 'funções eliminadas'));

      const subjectUser = summary.user || String(fallbackContext.user || job?.params?.target_user || '').trim();
      const subjectSystem = summary.system || String(fallbackContext.system || job?.params?.target_system_key || '').trim();
      const subject = subjectUser && subjectSystem
        ? `${subjectUser} no sistema ${subjectSystem}`
        : subjectUser
          ? `o utilizador ${subjectUser}`
          : 'o processo';

      let message = `${actionName} ${stateLabel} para ${subject}`;
      if (jobIdShort) {
        message += ` (job #${jobIdShort})`;
      }
      if (counts.length > 0) {
        message += `: ${counts.join(', ')}.`;
      }
      if (summary.noMatches) {
        message += ' O log indica que não foram encontradas funções no sistema alvo.';
      } else if (summary.removed === 0) {
        message += isEndDate ? ' Sem funções alteradas.' : ' Sem funções eliminadas.';
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

          window.setTimeout(() => {
            renderPostAnalysisFollowUpQuestion();
          }, 400);
        } catch (error) {
          if (requestId !== authorizationRemovalJobRequestId) {
            return;
          }

          window.setTimeout(check, 3000);
        }
      }

      window.setTimeout(check, 1200);
    }

    function renderPostAnalysisFollowUpQuestion() {
      hideAuthorizationTypingIndicator();

      appendAuthorizationAssistantMessage(
        'Deseja seguir com alguma ação sobre o processo de **Análise de Autorizações SAP** para este utilizador?'
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '8px';
      grid.style.marginBottom = '14px';

      const actions = [
        {
          label: '✅ Listar funções ativas',
          val: 'Listar funções ativas',
          action: () => {
            appendAuthorizationMessage('user', 'Listar funções ativas');
            handleContextualRolesQuery('Listar funções ativas', 'ATIVAS');
          }
        },
        {
          label: '❌ Listar funções expiradas',
          val: 'Listar funções expiradas',
          action: () => {
            appendAuthorizationMessage('user', 'Listar funções expiradas');
            handleContextualRolesQuery('Listar funções expiradas', 'EXPIRADAS');
          }
        },
        {
          label: '🔄 Nova análise',
          val: 'Nova análise',
          action: () => {
            appendAuthorizationMessage('user', 'Nova análise');
            resetAuthorizationChat();
          }
        },
        {
          label: '⚙️ Seguir com ação',
          val: 'Seguir com ação',
          action: () => {
            appendAuthorizationMessage('user', 'Seguir com ação sobre esta análise');
            showPfcgProcessExecutionOptions();
          }
        }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      actions.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-system-card';
        btn.style.flex = '0 0 auto';
        btn.style.padding = '8px 14px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          item.action();
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
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

      // 0.1 Se o utilizador responder afirmativamente à pergunta pós-resumo
      if (
        queryNorm === 'SIM' ||
        queryNorm === 'QUERO' ||
        queryNorm.startsWith('SIM ') ||
        queryNorm.startsWith('SIM,') ||
        queryNorm.includes('SEGUIR COM ACAO') ||
        queryNorm.includes('SEGUIR COM AÇÃO') ||
        queryNorm.includes('REALIZAR ACAO') ||
        queryNorm.includes('REALIZAR AÇÃO') ||
        queryNorm.includes('PRETENDO REALIZAR')
      ) {
        showPfcgProcessExecutionOptions();
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
      const existingUser = authorizationLastStatusData?.target_user || authorizationTargetUser;
      const existingSys = authorizationSelectedSystem;
      const existingRoles = authorizationLastDisplayedRoles.length > 0 ? authorizationLastDisplayedRoles : (authorizationLastStatusData?.roles || []);

      authorizationIndividualContext = {
        processName,
        category,
        scriptName,
        targetUser: existingUser || '',
        selectedSystem: existingSys || null,
        parameters: {}
      };

      const isCreateUserProc = category === 'CUA_CRIAR_USER' || category === 'CUA_ADICIONAR' || String(scriptName || '').includes('CUA_CRIAR_USER') || String(scriptName || '').includes('CUA_ADICIONAR') || String(processName || '').includes('Criar utilizador');

      if (isCreateUserProc) {
        authorizationTargetUser = '';
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `Como pretende efetuar a **Criar utilizador** (${escapeAuthorizationText(processName)}) em modo individual?`
        );

        const container = document.getElementById('authorization-chat-messages');
        if (!container) return;

        const grid = document.createElement('div');
        grid.style.display = 'grid';
        grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(240px, 1fr))';
        grid.style.gap = '10px';
        grid.style.marginTop = '6px';
        grid.style.marginBottom = '10px';

        // Card 1: Novo
        const btnNew = document.createElement('button');
        btnNew.type = 'button';
        btnNew.className = 'auth-chat-analysis-card';
        btnNew.style.padding = '12px 14px';
        btnNew.onclick = () => {
          if (btnNew.parentElement) {
            btnNew.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btnNew.classList.add('selected');
          btnNew.setAttribute('aria-pressed', 'true');
          appendAuthorizationMessage('user', '🆕 Novo');
          authorizationIndividualContext.creationType = 'NEW';
          promptHrSearchForNewUser(processName);
        };
        btnNew.innerHTML = `
          <span class="analysis-title">🆕 Novo</span>
          <span class="analysis-desc">Criar um utilizador novo selecionando os dados do RH no sistema produtivo (S4P).</span>
        `;
        grid.appendChild(btnNew);

        // Card 2: Por cópia
        const btnCopy = document.createElement('button');
        btnCopy.type = 'button';
        btnCopy.className = 'auth-chat-analysis-card';
        btnCopy.style.padding = '12px 14px';
        btnCopy.onclick = () => {
          if (btnCopy.parentElement) {
            btnCopy.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btnCopy.classList.add('selected');
          btnCopy.setAttribute('aria-pressed', 'true');
          appendAuthorizationMessage('user', '📋 Por cópia');
          authorizationIndividualContext.creationType = 'COPY';
          promptCopyReferenceUser(processName);
        };
        btnCopy.innerHTML = `
          <span class="analysis-title">📋 Por cópia</span>
          <span class="analysis-desc">Criar um utilizador copiando autorizações/dados de um utilizador SAP de referência.</span>
        `;
        grid.appendChild(btnCopy);

        container.appendChild(grid);
        container.scrollTop = container.scrollHeight;
        return;
      }

      if (existingUser && existingSys && existingRoles.length > 0) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationAssistantMessage(
          `A utilizar a lista de autorizações da análise anterior para o utilizador **${escapeAuthorizationText(existingUser)}** no ambiente **${escapeAuthorizationText(existingSys.system || existingSys.key)}** (${escapeAuthorizationText(processName)}):`
        );
        renderIndividualUserRolesList(existingRoles, existingUser, existingSys);
        return;
      }

      authorizationTargetUser = '';
      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER;

      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        `Para a **Alteração Individual** do processo **${escapeAuthorizationText(processName)}**, por favor indique qual é o utilizador SAP sobre o qual pretende efetuar a alteração.`
      );
      updateAuthorizationComposer();

      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    function promptHrSearchForNewUser(processName) {
      authorizationChatState = AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY;
      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        `Para a criação de um utilizador **Novo** (${escapeAuthorizationText(processName)}), vamos selecionar os dados do colaborador na **tabela do RH no Sistema Produtivo (S4P)** para obter o Nome, Email e Equipa.\n\nPor favor, introduza o **N.º Mecanográfico (PERNR)**, **Nome** ou **Utilizador SAP** do colaborador.`
      );
      updateAuthorizationComposer();

      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    function promptCopyReferenceUser(processName) {
      authorizationChatState = AUTH_CHAT_STATES.WAITING_COPY_REFERENCE_USER;
      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        `Para a criação de um utilizador **Por cópia** (${escapeAuthorizationText(processName)}), por favor indique qual é o **Utilizador SAP de referência** a partir do qual pretende copiar as autorizações/dados.\n\nIntroduza o **N.º Mecanográfico (PERNR)**, **Nome** ou **Utilizador SAP** do utilizador de referência para pesquisa na tabela do RH em Produtivo (S4P):`
      );
      updateAuthorizationComposer();

      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    async function performHrReferenceUserSearch(queryVal) {
      hideAuthorizationTypingIndicator();
      showAuthorizationTypingIndicator(null, `A pesquisar utilizador de referência "${escapeAuthorizationText(queryVal)}" na tabela do RH no sistema produtivo (S4P)...`);

      try {
        const response = await fetch('/api/authorizations/hr-search', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query: queryVal, target_system_key: 'S4PCLNT100', max_results: 10 })
        });

        hideAuthorizationTypingIndicator();

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        if (!data || data.success !== true || !Array.isArray(data.data) || data.data.length === 0) {
          const fallbackUser = formatSapUserId(queryVal);
          appendAuthorizationMessage(
            'assistant',
            `⚠️ Nenhum registo de RH foi encontrado no sistema produtivo (S4P) para **"${escapeAuthorizationText(queryVal)}"**.\n\nDeseja utilizar **"${escapeAuthorizationText(fallbackUser)}"** diretamente como Utilizador de Referência ou efetuar outra pesquisa?`
          );

          const container = document.getElementById('authorization-chat-messages');
          if (!container) return;

          const grid = document.createElement('div');
          grid.style.display = 'flex';
          grid.style.flexWrap = 'wrap';
          grid.style.gap = '10px';
          grid.style.marginTop = '8px';
          grid.style.marginBottom = '12px';

          const btnUseDirect = document.createElement('button');
          btnUseDirect.type = 'button';
          btnUseDirect.className = 'auth-chat-system-card selected';
          btnUseDirect.style.padding = '8px 14px';
          btnUseDirect.onclick = () => {
            selectHrReferenceUserForCreation({ user_id: fallbackUser, full_name: fallbackUser, pernr: queryVal, team: 'N/D' });
          };
          btnUseDirect.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">✅ Utilizar "${escapeAuthorizationText(fallbackUser)}" como Referência</span>`;
          grid.appendChild(btnUseDirect);

          const btnRetry = document.createElement('button');
          btnRetry.type = 'button';
          btnRetry.className = 'auth-chat-system-card';
          btnRetry.style.padding = '8px 14px';
          btnRetry.onclick = () => {
            promptCopyReferenceUser(authorizationIndividualContext?.processName || 'Criar utilizador');
          };
          btnRetry.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">🔍 Pesquisar outro nome / PERNR</span>`;
          grid.appendChild(btnRetry);

          container.appendChild(grid);
          window.setTimeout(() => {
            container.scrollTop = container.scrollHeight;
          }, 50);
          return;
        }

        if (data.data.length === 1) {
          selectHrReferenceUserForCreation(data.data[0]);
          return;
        }

        appendAuthorizationMessage(
          'assistant',
          `Foram encontrados **${data.data.length}** registo(s) na tabela do RH no sistema produtivo (**S4P**). Por favor, selecione o **utilizador de referência**:`
        );

        renderHrReferenceResultsCards(data.data);
      } catch (err) {
        hideAuthorizationTypingIndicator();
        console.warn('[HR REF SEARCH] Erro na pesquisa RH:', err);
        const fallbackUser = formatSapUserId(queryVal);
        selectHrReferenceUserForCreation({ user_id: fallbackUser, full_name: fallbackUser, pernr: queryVal, team: 'N/D' });
      }
    }

    function renderHrReferenceResultsCards(items) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'grid';
      grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(280px, 1fr))';
      grid.style.gap = '10px';
      grid.style.marginTop = '8px';
      grid.style.marginBottom = '12px';

      items.forEach(item => {
        const card = document.createElement('div');
        card.className = 'auth-chat-analysis-card';
        card.style.padding = '12px 14px';
        card.style.cursor = 'pointer';

        const formattedUser = formatSapUserId(item.user_id || item.pernr);

        card.innerHTML = `
          <div style="font-weight:700; font-size:0.92rem; color:var(--text-primary); margin-bottom:4px;">👤 ${escapeAuthorizationText(item.full_name || 'N/D')}</div>
          <div style="font-size:0.82rem; color:var(--text-secondary); display:grid; gap:3px;">
            <div><b>• N.º Colaborador (PERNR):</b> ${escapeAuthorizationText(item.pernr)}</div>
            <div><b>• Utilizador SAP:</b> ${escapeAuthorizationText(formattedUser)}</div>
            <div><b>• Email:</b> ${escapeAuthorizationText(item.email || 'Não registado')}</div>
            <div><b>• Equipa:</b> ${escapeAuthorizationText(item.team || 'Geral')}</div>
            <div style="font-size:0.75rem; color:#10b981; margin-top:4px;">Sistema Produtivo S4P</div>
          </div>
          <button type="button" class="auth-chat-system-card selected" style="margin-top:10px; width:100%; justify-content:center; font-weight:700; padding:6px 10px;">
            ✅ Selecionar como Utilizador de Referência (${escapeAuthorizationText(formattedUser)})
          </button>
        `;

        const selectBtn = card.querySelector('button');
        selectBtn.onclick = (e) => {
          e.stopPropagation();
          selectHrReferenceUserForCreation(item);
        };
        card.onclick = () => selectHrReferenceUserForCreation(item);

        grid.appendChild(card);
      });

      container.appendChild(grid);
      window.setTimeout(() => {
        container.scrollTop = container.scrollHeight;
      }, 50);
    }

    function selectHrReferenceUserForCreation(item) {
      const rawUser = item.user_id || item.pernr || '';
      const refUser = formatSapUserId(rawUser);
      if (!authorizationIndividualContext) {
        authorizationIndividualContext = {};
      }
      authorizationIndividualContext.referenceUser = refUser;
      authorizationIndividualContext.referenceHrData = item;

      const nameLabel = item.full_name && item.full_name !== refUser ? ` (${item.full_name})` : '';

      appendAuthorizationMessage('user', `Utilizador de referência: ${refUser}${nameLabel}`);
      appendAuthorizationMessage(
        'assistant',
        `Registado o utilizador de referência **${escapeAuthorizationText(refUser)}**${escapeAuthorizationText(nameLabel)}.\n\nAgora, vamos selecionar os dados do **novo colaborador** na **tabela do RH no Sistema Produtivo (S4P)** para obter o Nome, Email e Equipa.\n\nPor favor, introduza o **N.º Mecanográfico (PERNR)**, **Nome** ou **Utilizador SAP** do novo colaborador.`
      );

      authorizationChatState = AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY;
      updateAuthorizationComposer();

      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    async function performHrUserSearch(queryVal) {
      hideAuthorizationTypingIndicator();
      showAuthorizationTypingIndicator(null, `A pesquisar colaborador "${escapeAuthorizationText(queryVal)}" na tabela do RH no sistema produtivo (S4P)...`);

      try {
        const response = await fetch('/api/authorizations/hr-search', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query: queryVal, target_system_key: 'S4PCLNT100', max_results: 5 })
        });

        hideAuthorizationTypingIndicator();

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        if (!data || data.success !== true || !Array.isArray(data.data) || data.data.length === 0) {
          const fallbackUser = formatSapUserId(queryVal);
          appendAuthorizationMessage(
            'assistant',
            `⚠️ Nenhum registo de RH foi encontrado no sistema produtivo (S4P) para **"${escapeAuthorizationText(queryVal)}"**.\n\nDeseja utilizar **"${escapeAuthorizationText(fallbackUser)}"** diretamente para criação do utilizador ou efetuar outra pesquisa?`
          );

          const container = document.getElementById('authorization-chat-messages');
          if (!container) return;

          const grid = document.createElement('div');
          grid.style.display = 'flex';
          grid.style.flexWrap = 'wrap';
          grid.style.gap = '10px';
          grid.style.marginTop = '8px';
          grid.style.marginBottom = '12px';

          const btnUseDirect = document.createElement('button');
          btnUseDirect.type = 'button';
          btnUseDirect.className = 'auth-chat-system-card selected';
          btnUseDirect.style.padding = '8px 14px';
          btnUseDirect.onclick = () => {
            selectHrUserForCreation({
              user_id: fallbackUser,
              full_name: queryVal,
              pernr: queryVal,
              email: queryVal.includes('@') ? queryVal : '',
              team: 'Geral'
            });
          };
          btnUseDirect.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">✅ Prosseguir com "${escapeAuthorizationText(fallbackUser)}"</span>`;
          grid.appendChild(btnUseDirect);

          const btnRetry = document.createElement('button');
          btnRetry.type = 'button';
          btnRetry.className = 'auth-chat-system-card';
          btnRetry.style.padding = '8px 14px';
          btnRetry.onclick = () => {
            promptHrSearchForNewUser(authorizationIndividualContext?.processName || 'Criar utilizador');
          };
          btnRetry.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">🔍 Pesquisar outro nome / PERNR / e-mail</span>`;
          grid.appendChild(btnRetry);

          container.appendChild(grid);
          window.setTimeout(() => {
            container.scrollTop = container.scrollHeight;
          }, 50);
          authorizationChatState = AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY;
          updateAuthorizationComposer();
          return;
        }

        if (data.data.length === 1) {
          selectHrUserForCreation(data.data[0]);
          return;
        }

        appendAuthorizationMessage(
          'assistant',
          `Foram encontrados **${data.data.length}** registo(s) na tabela do RH no sistema produtivo (**S4P**). Por favor, selecione o colaborador para criar o utilizador:`
        );

        renderHrResultsCards(data.data);
      } catch (err) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `❌ Ocorreu um erro ao consultar a tabela do RH em Produtivo (S4P): ${escapeAuthorizationText(err.message || err)}.\n\nPretende tentar novamente ou introduzir o utilizador manualmente?`
        );
        authorizationChatState = AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY;
        updateAuthorizationComposer();
      }
    }

    function renderHrResultsCards(items) {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'grid';
      grid.style.gridTemplateColumns = 'repeat(auto-fill, minmax(280px, 1fr))';
      grid.style.gap = '10px';
      grid.style.marginTop = '8px';
      grid.style.marginBottom = '12px';

      items.forEach(item => {
        const card = document.createElement('div');
        card.className = 'auth-chat-analysis-card';
        card.style.padding = '12px 14px';
        card.style.cursor = 'pointer';

        const formattedUser = formatSapUserId(item.user_id || item.pernr);

        card.innerHTML = `
          <div style="font-weight:700; font-size:0.92rem; color:var(--text-primary); margin-bottom:4px;">👤 ${escapeAuthorizationText(item.full_name || 'N/D')}</div>
          <div style="font-size:0.82rem; color:var(--text-secondary); display:grid; gap:3px;">
            <div><b>• Primeiro Nome (NAME_FIRST):</b> ${escapeAuthorizationText(item.first_name || 'N/D')} <span style="font-size:0.75rem; color:var(--text-tertiary);">(PA0002-VORNA)</span></div>
            <div><b>• Apelido (NAME_LAST):</b> ${escapeAuthorizationText(item.last_name || 'N/D')} <span style="font-size:0.75rem; color:var(--text-tertiary);">(PA0002-NACHN)</span></div>
            <div><b>• Email (SMTP_ADDR):</b> ${escapeAuthorizationText(item.email || 'Não registado')} <span style="font-size:0.75rem; color:var(--text-tertiary);">(PA0002/PA0105-USRID_LONG)</span></div>
            <div><b>• Equipa:</b> ${escapeAuthorizationText(item.team || 'Geral')}</div>
            <div><b>• ID/PERNR:</b> ${escapeAuthorizationText(item.pernr)} | <b>Utilizador CUA:</b> ${escapeAuthorizationText(formattedUser)}</div>
            <div style="font-size:0.75rem; color:#10b981; margin-top:4px;">Sistema Produtivo S4P</div>
          </div>
          <button type="button" class="auth-chat-system-card selected" style="margin-top:10px; width:100%; justify-content:center; font-weight:700; padding:6px 10px;">
            ✅ Selecionar e Criar Utilizador (${escapeAuthorizationText(formattedUser)})
          </button>
        `;

        const selectBtn = card.querySelector('button');
        selectBtn.onclick = (e) => {
          e.stopPropagation();
          selectHrUserForCreation(item);
        };
        card.onclick = () => selectHrUserForCreation(item);

        grid.appendChild(card);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function selectHrUserForCreation(item) {
      const rawUser = item.user_id || item.pernr || '';
      const selectedUser = formatSapUserId(rawUser);
      authorizationTargetUser = selectedUser;
      if (authorizationIndividualContext) {
        authorizationIndividualContext.targetUser = selectedUser;
        authorizationIndividualContext.hrData = item;
      }

      appendAuthorizationMessage('user', `Selecionado: ${item.full_name} (${selectedUser})`);
      appendAuthorizationMessage(
        'assistant',
        `
          <div class="auth-chat-summary">
            <div style="font-weight:700; margin-bottom:8px; color:#10b981; font-size:0.92rem;">✅ Colaborador Selecionado do RH (S4P)</div>
            <div style="display:grid; gap:4px; font-size:0.84rem; margin-bottom:10px;">
              <div><b>• CUA-NAME_FIRST:</b> ${escapeAuthorizationText(item.first_name || 'N/D')} <span style="font-size:0.75rem; color:var(--text-secondary);">(PA0002-VORNA)</span></div>
              <div><b>• CUA-NAME_LAST:</b> ${escapeAuthorizationText(item.last_name || 'N/D')} <span style="font-size:0.75rem; color:var(--text-secondary);">(PA0002-NACHN)</span></div>
              <div><b>• CUA-SMTP_ADDR:</b> ${escapeAuthorizationText(item.email || 'N/D')} <span style="font-size:0.75rem; color:var(--text-secondary);">(PA0002/PA0105-USRID_LONG)</span></div>
              <div><b>• Equipa / Org:</b> ${escapeAuthorizationText(item.team || 'N/D')}</div>
              <div><b>• Utilizador Alvo CUA:</b> ${escapeAuthorizationText(selectedUser)}</div>
            </div>
          </div>
        `,
        true
      );

      promptCuaUserDetails(item);
    }

    function promptCuaUserDetails(item) {
      const defaultFirstName = item?.first_name || '';
      const defaultLastName = item?.last_name || '';
      const defaultEmail = item?.email || '';
      // Função (FUNCTION), Departamento (DEPARTMENT) e Telefone (MOB_NUMBER) não existem na tabela de RH e têm de ser explicitamente solicitados
      const defaultFunc = '';
      const defaultDept = '';
      const defaultMob = '';

      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        'Por favor, preencha ou confirme os **6 dados para criação da conta no CUA** (Nota: **Função**, **Departamento** e **Telefone** não existem na tabela do RH e têm de ser solicitados e preenchidos):'
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const card = document.createElement('div');
      card.className = 'auth-chat-summary';
      card.style.display = 'grid';
      card.style.gap = '10px';
      card.style.marginTop = '6px';
      card.style.marginBottom = '10px';
      card.style.maxWidth = '480px';

      card.innerHTML = `
        <div style="font-weight:700; font-size:0.92rem; color:var(--text-primary); border-bottom:1px solid var(--border-color, #e2e8f0); padding-bottom:6px;">
          📝 Dados de Utilizador CUA (6 Campos Obrigatórios)
        </div>
        
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
          <div style="display:grid; gap:3px;">
            <label style="font-size:0.78rem; font-weight:700; color:var(--text-secondary);">• Nome (NAME_FIRST):</label>
            <input type="text" id="cua-field-first-name" class="form-control form-control-sm" value="${escapeAuthorizationText(defaultFirstName)}" placeholder="Primeiro Nome" style="padding:6px 10px; font-size:0.84rem; border-radius:6px; border:1px solid var(--border-color, #cbd5e1);">
          </div>
          <div style="display:grid; gap:3px;">
            <label style="font-size:0.78rem; font-weight:700; color:var(--text-secondary);">• Sobrenome (NAME_LAST):</label>
            <input type="text" id="cua-field-last-name" class="form-control form-control-sm" value="${escapeAuthorizationText(defaultLastName)}" placeholder="Apelido / Sobrenome" style="padding:6px 10px; font-size:0.84rem; border-radius:6px; border:1px solid var(--border-color, #cbd5e1);">
          </div>
        </div>

        <div style="display:grid; gap:3px;">
          <label style="font-size:0.78rem; font-weight:700; color:var(--text-secondary);">• Email (SMTP_ADDR):</label>
          <input type="email" id="cua-field-email" class="form-control form-control-sm" value="${escapeAuthorizationText(defaultEmail)}" placeholder="email@empresa.com" style="padding:6px 10px; font-size:0.84rem; border-radius:6px; border:1px solid var(--border-color, #cbd5e1);">
        </div>

        <div style="display:grid; gap:3px;">
          <label style="font-size:0.78rem; font-weight:700; color:#d97706;">• Função (FUNCTION) * (Solicitar):</label>
          <input type="text" id="cua-field-function" class="form-control form-control-sm" value="" placeholder="Escreva a Função (ex: Analista de Sistemas)" style="padding:6px 10px; font-size:0.84rem; border-radius:6px; border:1px solid #f59e0b;">
        </div>

        <div style="display:grid; gap:3px;">
          <label style="font-size:0.78rem; font-weight:700; color:#d97706;">• Departamento (DEPARTMENT) * (Solicitar):</label>
          <input type="text" id="cua-field-department" class="form-control form-control-sm" value="" placeholder="Escreva o Departamento (ex: SI/TI)" style="padding:6px 10px; font-size:0.84rem; border-radius:6px; border:1px solid #f59e0b;">
        </div>

        <div style="display:grid; gap:3px;">
          <label style="font-size:0.78rem; font-weight:700; color:#d97706;">• Telefone (MOB_NUMBER) * (Solicitar):</label>
          <input type="text" id="cua-field-mob-number" class="form-control form-control-sm" value="" placeholder="Escreva o Telefone (ex: +351 912345678)" style="padding:6px 10px; font-size:0.84rem; border-radius:6px; border:1px solid #f59e0b;">
        </div>

        <button type="button" id="cua-fields-submit-btn" class="btn btn-primary btn-sm" style="margin-top:6px; font-weight:700; padding:8px 14px; border-radius:6px; width:100%; justify-content:center;">
          ✅ Guardar os 6 Campos CUA e Continuar ➔
        </button>
      `;

      container.appendChild(card);
      container.scrollTop = container.scrollHeight;

      const submitBtn = card.querySelector('#cua-fields-submit-btn');
      submitBtn.onclick = () => {
        const firstNameVal = (card.querySelector('#cua-field-first-name')?.value || '').trim();
        const lastNameVal = (card.querySelector('#cua-field-last-name')?.value || '').trim();
        const emailVal = (card.querySelector('#cua-field-email')?.value || '').trim();
        const funcVal = (card.querySelector('#cua-field-function')?.value || '').trim();
        const deptVal = (card.querySelector('#cua-field-department')?.value || '').trim();
        const mobVal = (card.querySelector('#cua-field-mob-number')?.value || '').trim();

        saveCuaUserDetails(firstNameVal, lastNameVal, emailVal, funcVal, deptVal, mobVal);
      };

      authorizationChatState = AUTH_CHAT_STATES.WAITING_CUA_FUNCTION;
      updateAuthorizationComposer();
    }

    function saveCuaUserDetails(firstNameVal, lastNameVal, emailVal, funcVal, deptVal, mobVal) {
      if (!authorizationIndividualContext) {
        authorizationIndividualContext = {};
      }
      if (!authorizationIndividualContext.parameters) {
        authorizationIndividualContext.parameters = {};
      }

      authorizationIndividualContext.parameters.NAME_FIRST = firstNameVal;
      authorizationIndividualContext.parameters.NAME_LAST = lastNameVal;
      authorizationIndividualContext.parameters.SMTP_ADDR = emailVal;
      authorizationIndividualContext.parameters.FUNCTION = funcVal;
      authorizationIndividualContext.parameters.DEPARTMENT = deptVal;
      authorizationIndividualContext.parameters.MOB_NUMBER = mobVal;

      appendAuthorizationMessage(
        'assistant',
        `
          <div class="auth-chat-summary">
            <div style="font-weight:700; margin-bottom:8px; color:#10b981; font-size:0.92rem;">✅ Dados dos 6 Campos CUA Registados</div>
            <div style="display:grid; gap:4px; font-size:0.84rem;">
              <div><b>• Nome (NAME_FIRST):</b> ${escapeAuthorizationText(firstNameVal || 'Não preenchido')}</div>
              <div><b>• Sobrenome (NAME_LAST):</b> ${escapeAuthorizationText(lastNameVal || 'Não preenchido')}</div>
              <div><b>• Email (SMTP_ADDR):</b> ${escapeAuthorizationText(emailVal || 'Não preenchido')}</div>
              <div><b>• Função (FUNCTION):</b> ${escapeAuthorizationText(funcVal || 'Não preenchido')}</div>
              <div><b>• Departamento (DEPARTMENT):</b> ${escapeAuthorizationText(deptVal || 'Não preenchido')}</div>
              <div><b>• Telefone (MOB_NUMBER):</b> ${escapeAuthorizationText(mobVal || 'Não preenchido')}</div>
            </div>
          </div>
          Em que sistema/ambiente SAP pretende efetuar a criação da conta CUA?
        `,
        true
      );

      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_SYSTEM;
      showIndividualSystemOptions();
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
        card.setAttribute('data-system', sys.system || sys.key);

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
      appendAuthorizationAssistantMessage(`Selecionada a rotina **${escapeAuthorizationText(categoryName)}**. Em que sistema/ambiente SAP pretende efetuar a execução?`);
      showIndividualSystemOptions((sys) => {
        authorizationSelectedSystem = sys;
        const sysLabel = sys.label || sys.system || sys.key;
        appendAuthorizationMessage('user', sysLabel);
        renderRoutineSuggestionsForSystem(sys);
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

      const scriptName = authorizationIndividualContext.scriptName;
      if (scriptName) {
        hideAuthorizationTypingIndicator();
        showAuthorizationTypingIndicator(
          null,
          `A criar job de execução do script **${escapeAuthorizationText(scriptName)}** (${escapeAuthorizationText(proc)}) para o utilizador **${escapeAuthorizationText(user)}**...`
        );
        try {
          const isCuaScript = scriptName.includes('CUA') || scriptName.includes('su01_reset') || (authorizationIndividualContext.category || '').includes('CUA');
          let executionAmb = isCuaScript ? 'CUA' : (sys.system || sys.key || 'PRD').toUpperCase();
          let targetSysKey = sys.key || sys.system || 'S4PCLNT100';

          const targetFolder = (scriptName && scriptName.includes('su01_reset')) ? 'CUA Login' : 'Funções PFCG';
          const formData = new FormData();
          formData.append('task', 'sap_cockpit');
          formData.append('ambiente', executionAmb);
          formData.append('processo', targetFolder);
          formData.append('subprocesso', scriptName);
          formData.append('target_user', user);
          formData.append('target_system', targetSysKey);
          formData.append('target_env', targetSysKey);
          formData.append('subsystem_filter', targetSysKey);
          formData.append('opcao_processamento', '1');

          if (authorizationIndividualContext.referenceUser) {
            formData.append('reference_user', authorizationIndividualContext.referenceUser);
            formData.append('target_user_ref', authorizationIndividualContext.referenceUser);
          }

          if (authorizationIndividualContext.parameters) {
            if (authorizationIndividualContext.parameters.NAME_FIRST) {
              formData.append('first_name', authorizationIndividualContext.parameters.NAME_FIRST);
            }
            if (authorizationIndividualContext.parameters.NAME_LAST) {
              formData.append('last_name', authorizationIndividualContext.parameters.NAME_LAST);
            }
            if (authorizationIndividualContext.parameters.SMTP_ADDR) {
              formData.append('email', authorizationIndividualContext.parameters.SMTP_ADDR);
            }
            if (authorizationIndividualContext.parameters.FUNCTION) {
              formData.append('function', authorizationIndividualContext.parameters.FUNCTION);
            }
            if (authorizationIndividualContext.parameters.DEPARTMENT) {
              formData.append('department', authorizationIndividualContext.parameters.DEPARTMENT);
            }
            if (authorizationIndividualContext.parameters.MOB_NUMBER) {
              formData.append('mob_number', authorizationIndividualContext.parameters.MOB_NUMBER);
            }
            formData.append('parameters', JSON.stringify(authorizationIndividualContext.parameters));
          }

          const response = await fetch('/jobs', {
            method: 'POST',
            body: formData
          });

          if (!response.ok) {
            const errData = await response.json().catch(() => ({}));
            throw new Error(errData.detail || `HTTP ${response.status}`);
          }

          const data = await response.json();
          const processLabel = authorizationIndividualContext.processName || proc;
          const techCategory = authorizationIndividualContext.category || 'CUA_CRIAR_USER';
          const displayProcess = (processLabel && processLabel !== techCategory)
            ? `${processLabel} (${techCategory})`
            : techCategory;

          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage(
            'assistant',
            `
              <div class="auth-chat-summary">
                <div style="font-weight:700; margin-bottom:8px; color:#10b981; font-size:0.92rem;">✅ Job de Execução Criado no SAP Cockpit!</div>
                <div style="display:grid; gap:4px; font-size:0.84rem; margin-bottom:10px;">
                  <div><b>• Script:</b> ${escapeAuthorizationText(scriptName)}</div>
                  <div><b>• Processo:</b> ${escapeAuthorizationText(displayProcess)}</div>
                  <div><b>• Utilizador Alvo:</b> ${escapeAuthorizationText(user)}</div>
                  <div><b>• Ambiente SAP:</b> ${escapeAuthorizationText(sys.system || sys.key)}</div>
                  <div><b>• Job ID:</b> #${escapeAuthorizationText(String(data.job_id || data.id || '').slice(0, 8))}</div>
                </div>
                <div style="font-size:0.8rem; color:var(--text-secondary);">O worker SAP foi notificado e iniciou a execução da rotina <b>${escapeAuthorizationText(scriptName)}</b> no ambiente SAP.</div>
              </div>
            `,
            true
          );
          showNextActionsPrompt('O job foi submetido com sucesso para execução no SAP! Pretende efetuar mais alguma ação?');
          authorizationChatState = AUTH_CHAT_STATES.READY;
          updateAuthorizationComposer();
          return;
        } catch (jobErr) {
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', `❌ Não foi possível criar o job para ${escapeAuthorizationText(scriptName)}: ${jobErr.message || jobErr}`);
          showNextActionsPrompt('Ocorreu um erro ao submeter o job. Pretende tentar novamente ou escolher outro processo?');
          return;
        }
      }

    function showNextActionsPrompt(messageText = 'O processo foi concluído. Deseja efetuar mais alguma ação?') {
      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage('assistant', messageText);

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '8px';
      grid.style.marginBottom = '10px';

      const actions = [
        {
          label: '🔄 Executar outro processo',
          action: () => {
            appendAuthorizationMessage('user', 'Executar outro processo');
            renderRoutineSuggestionsInitial();
          }
        },
        {
          label: '🔑 Alterar outra senha',
          action: () => {
            appendAuthorizationMessage('user', 'Alterar outra senha');
            selectUserDataSubroutine({ label: '🔑 Alterar Senha', val: 'Alterar Senha', scriptName: 'su01_reset_password.py', category: 'CUA Login' });
          }
        }
      ];

      actions.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-system-card';
        btn.style.flex = '0 0 auto';
        btn.style.padding = '8px 14px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          item.action();
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

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

      // Submissão automática do job de acordo com o processo escolhido
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
      } else if (script.includes('CUA_ENDDATE') || procName.includes('CUA_ENDDATE') || procName.includes('ENDDATE')) {
        window.setTimeout(() => {
          confirmAuthorizationEndDate(targetUser, sys, roleList);
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
    async function confirmAuthorizationRemoval(targetUser, sys, roleList) {
      const user = targetUser || authorizationPendingRemoval?.targetUser || authorizationTargetUser;
      const systemKey = (sys && (sys.key || sys.system)) || authorizationPendingRemoval?.targetSystemKey || authorizationSelectedSystem?.key || '';
      const rawRoles = Array.isArray(roleList) ? roleList : (Array.isArray(authorizationPendingRemoval?.roles) ? authorizationPendingRemoval.roles : []);

      const roles = rawRoles.map(item => {
        if (typeof item === 'string') return { role: item };
        return { role: item.role || item.function || item.agr_name || item.AGR_NAME || item.name || '' };
      }).filter(r => Boolean(r.role));

      const payload = {
        target_user: user,
        target_system_key: systemKey,
        roles: roles,
        opcao_processamento: 'sistema_user'
      };

      if (!payload.target_user || !payload.target_system_key || payload.roles.length === 0) {
        appendAuthorizationMessage('assistant', 'Não tenho funções pendentes para remover.');
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

    async function confirmAuthorizationEndDate(targetUser, sys, roleList) {
      const user = targetUser || authorizationPendingRemoval?.targetUser || authorizationTargetUser;
      const systemKey = (sys && (sys.key || sys.system)) || authorizationPendingRemoval?.targetSystemKey || authorizationSelectedSystem?.key || '';
      const rawRoles = Array.isArray(roleList) ? roleList : (Array.isArray(authorizationPendingRemoval?.roles) ? authorizationPendingRemoval.roles : []);

      const roles = rawRoles.map(item => {
        if (typeof item === 'string') return { role: item };
        return { role: item.role || item.function || item.agr_name || item.AGR_NAME || item.name || '' };
      }).filter(r => Boolean(r.role));

      const payload = {
        target_user: user,
        target_system_key: systemKey,
        roles: roles
      };

      if (!payload.target_user || !payload.target_system_key || payload.roles.length === 0) {
        appendAuthorizationMessage('assistant', 'Não tenho dados suficientes para criar o pedido de CUA_ENDDATE.');
        return;
      }

      appendAuthorizationMessage('assistant', 'A preparar o pedido de alteração de validade (CUA_ENDDATE) das funções selecionadas...');

      try {
        const response = await fetch('/api/authorizations/enddate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        const data = await response.json().catch(() => ({}));
        if (!response.ok) {
          const detail = data.detail || data.message || '';
          throw new Error(detail || `A API devolveu HTTP ${response.status}`);
        }

        appendAuthorizationMessage(
          'assistant',
          `Job CUA_ENDDATE criado com sucesso no CUA. Job #${String(data.job_id || '').slice(0, 8)} para alterar validade de ${data.roles_count || roleList.length} funções no sistema alvo.`
        );
        appendAuthorizationMessage(
          'assistant',
          'Vou acompanhar a execução do job no SAP GUI e mostrar o resultado final quando terminar.'
        );
        authorizationRemovalJobRequestId++;
        showAuthorizationTypingIndicator(authorizationRemovalJobRequestId, 'A aguardar a execução do CUA_ENDDATE...');
        pollAuthorizationRemovalJob(data.job_id, authorizationRemovalJobRequestId);
      } catch (error) {
        appendAuthorizationMessage(
          'assistant',
          `Não foi possível criar o pedido CUA_ENDDATE: ${error?.message || error}`
        );
      }
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
        if (requestId !== null) {
          existing.dataset.requestId = requestId;
        }
        const existingLabel = existing.querySelector('.auth-chat-typing-label');
        if (existingLabel) {
          existingLabel.innerHTML = label ? String(label) : '';
        }
        container.scrollTop = container.scrollHeight;
        return;
      }

      removeAuthorizationTypingIndicator();

      const typingDiv = document.createElement('div');
      typingDiv.className = 'auth-chat-typing';
      typingDiv.dataset.authorizationTyping = 'true';
      if (requestId !== null) {
        typingDiv.dataset.requestId = requestId;
      }
      typingDiv.innerHTML = `
        <span class="auth-chat-typing-dot"></span>
        <span class="auth-chat-typing-dot"></span>
        <span class="auth-chat-typing-dot"></span>
        <span class="auth-chat-typing-label">${label ? String(label) : ''}</span>
      `;
      container.appendChild(typingDiv);
      container.scrollTop = container.scrollHeight;
    }

    function hideAuthorizationTypingIndicator() {
      removeAuthorizationTypingIndicator();
    }

    function parseSimpleAuthorizationMarkdown(text) {
      if (!text) return '';
      let html = String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')
        .replace(/\n\n/g, '<br><br>')
        .replace(/\n/g, '<br>');
      return html;
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
          contentSpan.innerHTML = parseSimpleAuthorizationMarkdown(text);
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
      authorizationUatCreateDocumentJobRequestId++;
      clearAuthorizationTimers();
      removeAuthorizationTypingIndicator();
      resetUatCreateDocumentFlow();

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
      const normText = String(text || '').trim().toUpperCase();

      if (!normText) {
        return;
      }

      if (normText.includes('UAT SIMULACAO') || normText.includes('UAT SIMULAÇÃO')) {
        appendAuthorizationMessage('user', text);
        showUatSimulationSubroutineOptions();
        return;
      }

      if (normText.includes('DADOS DE UTILIZADOR') || normText.includes('DADOS DO UTILIZADOR') || normText.includes('DADOS MESTRES') || normText.includes('UTILIZADOR')) {
        appendAuthorizationMessage('user', text);
        showUserDataSubroutineOptions();
        return;
      }

      if (normText.includes('PERFIL DE AUTORIZACAO') || normText.includes('FUNCOES PFCG') || normText.includes('PFCG') || normText === 'CUA') {
        appendAuthorizationMessage('user', text);
        showAuthorizationProfileSubroutineOptions();
        return;
      }

      if (normText.includes('CODIGOS IVA') || normText === 'IVA') {
        appendAuthorizationMessage('user', text);
        promptProcessMode('Criar/Manter Códigos IVA (FTXP)', 'Códigos IVA', 'FTXP_CRIAR_CODIGO_IVA.py', 'Automatização FTXP');
        return;
      }

      if (normText.includes('REVERTER') || normText.includes('ESTORNO') || normText.includes('DOCUMENTO')) {
        appendAuthorizationMessage('user', text);
        promptProcessMode('Reverter Documento Contabilístico', 'Reverter Documento', 'REVERTER_DOCUMENTO.py', 'Anulação de documentos FB08/FB05');
        return;
      }

      if (normText.includes('BANCO') || normText.includes('CHAVE DE BANCO')) {
        appendAuthorizationMessage('user', text);
        promptProcessMode('Chave de Banco', 'Chave de Banco', 'CHAVE_DE_BANCO.py', 'Criação de chave de banco FI01/FI02');
        return;
      }

      if (normText.includes('CADEIA') || normText.includes('CADEIAS DE PESQUISA')) {
        appendAuthorizationMessage('user', text);
        promptProcessMode('Cadeias de Pesquisa', 'Cadeias de Pesquisa', 'CADEIAS_DE_PESQUISA.py', 'Configuração de cadeias OT83');
        return;
      }

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
        `Ambiente **${escapeAuthorizationText(sysLabel)}** registado. Qual é o utilizador SAP que pretende analisar?`
      );
      authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
      updateAuthorizationComposer();
    }

    function renderAuthorizationInitialChoice() {
      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.WAITING_INITIAL_CHOICE;
      updateAuthorizationComposer();

      appendAuthorizationMessage(
        'assistant',
        'Olá! O que pretende fazer hoje?\nSelecione uma das opções abaixo ou escreva no campo inferior:',
        false,
        { key: 'authorization-initial-question' }
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '6px';
      grid.style.marginBottom = '8px';

      const choices = [
        { label: '🎫 Analisar ticket', val: 'Ticket' },
        { label: '⚙️ Executar processo', val: 'Processo' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      choices.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-system-card';
        btn.style.flex = '0 0 auto';
        btn.style.padding = '10px 16px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          handleInitialChoiceSelect(item.val);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.86rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function handleInitialChoiceSelect(choiceVal, skipUserMsg = false) {
      const normChoice = String(choiceVal || '').trim().toLowerCase();
      if (normChoice.includes('ticket')) {
        if (!skipUserMsg) appendAuthorizationMessage('user', 'Analisar ticket');
        showJiraTeamOptions();
      } else {
        if (!skipUserMsg) appendAuthorizationMessage('user', 'Executar processo');
        renderRoutineSuggestionsInitial();
      }
    }

    function showJiraTeamOptions() {
      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.WAITING_JIRA_TEAM;
      updateAuthorizationComposer();

      appendAuthorizationMessage(
        'assistant',
        'De qual Equipa é o ticket que pretende analisar?'
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '6px';
      grid.style.marginBottom = '8px';

      const predefinedTeams = [
        { label: '🏢 Business Intelligence', val: 'Business Intelligence' },
        { label: '⚙️ Core Systems', val: 'Core Systems' },
        { label: '💻 Development', val: 'Development' },
        { label: '🌐 Digital', val: 'Digital' },
        { label: '🎧 Helpdesk', val: 'Helpdesk' },
        { label: '🛍️ Retail Systems', val: 'Retail Systems' },
        { label: '🖥️ Systems Administration and Network', val: 'Systems Administration and Network' },
        { label: '🌐 Todas as Equipas', val: 'Todas as Equipas' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      predefinedTeams.forEach(item => {
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
          selectJiraTeam(item.val);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function selectJiraTeam(teamName, skipUserMsg = false) {
      showJiraAssigneeOptions(teamName, skipUserMsg);
    }

    function getAssigneesForTeam(teamName, tickets) {
      const isAllTeams = teamName.toLowerCase().includes('todas');
      const teamTickets = isAllTeams
        ? tickets
        : tickets.filter(t => {
            const tTeam = (t.team || '').trim().toLowerCase();
            const sTeam = teamName.trim().toLowerCase();
            return tTeam === sTeam || tTeam.includes(sTeam) || sTeam.includes(tTeam);
          });

      const assigneeSet = new Set();
      const TEAM_MEMBERS_MAP = {
        "Core Systems": ["Clayton Lopes", "Rita Rodrigues", "Filipe Galego", "Paula Silva", "José Pereira"],
        "Helpdesk": ["Filipe Abreu", "Miguel Ribeiro", "Alexandre Rodrigues"],
        "Retail Systems": ["Vitor.Pereira", "Marisa Moreira", "Sandra Gomes"],
        "Digital": ["Sandra Gomes", "Vitor Silva", "Diogo Oliveira"],
        "Systems Administration and Network": ["Alexandre Rodrigues"],
        "Business Intelligence": ["Mariana Pinto"],
        "Development": ["Joao.Pinheiro", "Pedro Silva"]
      };

      if (TEAM_MEMBERS_MAP[teamName]) {
        TEAM_MEMBERS_MAP[teamName].forEach(m => assigneeSet.add(m.trim()));
      }

      teamTickets.forEach(t => {
        if (t.assignee && typeof t.assignee === 'string' && t.assignee.trim() && t.assignee.trim() !== 'Sem responsável') {
          assigneeSet.add(t.assignee.trim());
        }
      });

      const list = Array.from(assigneeSet)
        .map(name => ({ label: `👤 ${name}`, val: name }))
        .sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      list.push({ label: '🌐 Todos os responsáveis', val: 'Todos os responsáveis' });
      return list;
    }

    async function showJiraAssigneeOptions(teamName, skipUserMsg = false) {
      if (!skipUserMsg) {
        appendAuthorizationMessage('user', teamName);
      }
      authorizationSelectedJiraTeam = teamName;
      authorizationChatState = AUTH_CHAT_STATES.WAITING_JIRA_ASSIGNEE;
      updateAuthorizationComposer();

      showAuthorizationTypingIndicator(null, `A carregar responsáveis da equipa ${escapeAuthorizationText(teamName)}...`);

      try {
        let tickets = authorizationCachedJiraTickets;
        if (!tickets) {
          const res = await fetch('/api/jira/tickets?limit=50000&exclude_closed=true');
          if (res.ok) {
            const data = await res.json();
            tickets = data.tickets || [];
            authorizationCachedJiraTickets = tickets;
          } else {
            tickets = [];
          }
        }

        hideAuthorizationTypingIndicator();

        const assigneeItems = getAssigneesForTeam(teamName, tickets);

        appendAuthorizationMessage(
          'assistant',
          `De qual responsável da equipa **${escapeAuthorizationText(teamName)}** pretende ver os tickets?`
        );

        const container = document.getElementById('authorization-chat-messages');
        if (!container) return;

        const grid = document.createElement('div');
        grid.style.display = 'flex';
        grid.style.flexWrap = 'wrap';
        grid.style.gap = '10px';
        grid.style.marginTop = '6px';
        grid.style.marginBottom = '8px';

        assigneeItems.forEach(item => {
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
            selectJiraAssignee(item.val);
          };
          btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
          grid.appendChild(btn);
        });

        container.appendChild(grid);
        container.scrollTop = container.scrollHeight;
      } catch (err) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage('assistant', `⚠️ Erro ao obter dados do Jira: ${escapeAuthorizationText(err.message)}`);
      }
    }

    async function selectJiraAssignee(assigneeName, skipUserMsg = false) {
      if (!skipUserMsg) {
        appendAuthorizationMessage('user', assigneeName);
      }
      authorizationSelectedJiraAssignee = assigneeName;
      authorizationChatState = AUTH_CHAT_STATES.WAITING_JIRA_FILTER_MODE;
      updateAuthorizationComposer();

      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage(
        'assistant',
        `Pretende listar todos os tickets de **${escapeAuthorizationText(assigneeName)}** ou filtrar por Processo?`
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '6px';
      grid.style.marginBottom = '8px';

      const filterChoices = [
        { label: '📋 Todos os tickets', val: 'Todos os tickets' },
        { label: '⚙️ Filtrar por processo', val: 'Filtrar por processo' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      filterChoices.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-system-card';
        btn.style.flex = '0 0 auto';
        btn.style.padding = '8px 14px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          handleFilterModeSelect(item.val);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function handleFilterModeSelect(modeVal, skipUserMsg = false) {
      const normMode = String(modeVal || '').trim().toLowerCase();
      if (normMode.includes('processo')) {
        if (!skipUserMsg) appendAuthorizationMessage('user', 'Filtrar por processo');
        showJiraProcessOptions(authorizationSelectedJiraTeam, authorizationSelectedJiraAssignee, true);
      } else {
        if (!skipUserMsg) appendAuthorizationMessage('user', 'Todos os tickets');
        authorizationSelectedJiraProcess = null;
        renderJiraTicketsList({
          teamName: authorizationSelectedJiraTeam,
          assigneeName: authorizationSelectedJiraAssignee,
          processName: null
        });
      }
    }

    function getProcessesForAssignee(teamName, assigneeName, tickets) {
      const isAllTeams = !teamName || teamName.toLowerCase().includes('todas');
      const isAllAssignees = !assigneeName || assigneeName.toLowerCase().includes('todos');

      const filtered = tickets.filter(t => {
        const matchTeam = isAllTeams || (t.team && (t.team.trim().toLowerCase() === teamName.trim().toLowerCase() || t.team.trim().toLowerCase().includes(teamName.trim().toLowerCase())));
        const matchAssignee = isAllAssignees || (t.assignee && (t.assignee.trim().toLowerCase() === assigneeName.trim().toLowerCase() || t.assignee.trim().toLowerCase().includes(assigneeName.trim().toLowerCase())));
        return matchTeam && matchAssignee;
      });

      const processSet = new Set();
      filtered.forEach(t => {
        if (t.process && typeof t.process === 'string' && t.process.trim()) {
          processSet.add(t.process.trim());
        }
        if (t.ticket_type && typeof t.ticket_type === 'string' && t.ticket_type.trim()) {
          processSet.add(t.ticket_type.trim());
        }
      });

      // Default processes as fallback/options
      const defaultProcesses = [
        'Cadeias de pesquisa',
        'Chave de banco',
        'Códigos IVA',
        'Dados de utilizador',
        'Perfil de autorização',
        'Reverter documento'
      ];
      defaultProcesses.forEach(p => processSet.add(p));

      const list = Array.from(processSet)
        .map(proc => ({ label: `⚙️ ${proc}`, val: proc }))
        .sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      list.push({ label: '🌐 Todos os processos', val: 'Todos os processos' });
      return list;
    }

    async function showJiraProcessOptions(teamName, assigneeName, skipUserMsg = false) {
      authorizationChatState = AUTH_CHAT_STATES.WAITING_JIRA_PROCESS;
      updateAuthorizationComposer();

      showAuthorizationTypingIndicator(null, `A carregar processos de ${escapeAuthorizationText(assigneeName || teamName)}...`);

      try {
        let tickets = authorizationCachedJiraTickets;
        if (!tickets) {
          const res = await fetch('/api/jira/tickets?limit=50000&exclude_closed=true');
          if (res.ok) {
            const data = await res.json();
            tickets = data.tickets || [];
            authorizationCachedJiraTickets = tickets;
          } else {
            tickets = [];
          }
        }

        hideAuthorizationTypingIndicator();

        const processItems = getProcessesForAssignee(teamName, assigneeName, tickets);

        appendAuthorizationMessage(
          'assistant',
          `Qual é o Processo que pretende filtrar para **${escapeAuthorizationText(assigneeName || 'a equipa')}**?`
        );

        const container = document.getElementById('authorization-chat-messages');
        if (!container) return;

        const grid = document.createElement('div');
        grid.style.display = 'flex';
        grid.style.flexWrap = 'wrap';
        grid.style.gap = '10px';
        grid.style.marginTop = '6px';
        grid.style.marginBottom = '8px';

        processItems.forEach(item => {
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
            selectJiraProcess(item.val);
          };
          btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
          grid.appendChild(btn);
        });

        container.appendChild(grid);
        container.scrollTop = container.scrollHeight;
      } catch (err) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage('assistant', `⚠️ Erro ao obter os processos do Jira: ${escapeAuthorizationText(err.message)}`);
      }
    }

    function selectJiraProcess(processName, skipUserMsg = false) {
      if (!skipUserMsg) {
        appendAuthorizationMessage('user', processName);
      }
      authorizationSelectedJiraProcess = processName;
      renderJiraTicketsList({
        teamName: authorizationSelectedJiraTeam,
        assigneeName: authorizationSelectedJiraAssignee,
        processName: processName
      });
    }

    async function renderJiraTicketsList({ teamName, assigneeName, processName }) {
      authorizationChatState = AUTH_CHAT_STATES.WAITING_JIRA_TICKET;
      updateAuthorizationComposer();

      const displayAssignee = assigneeName || 'Responsável';
      showAuthorizationTypingIndicator(null, `A filtrar tickets de ${escapeAuthorizationText(displayAssignee)}...`);

      try {
        let tickets = authorizationCachedJiraTickets;
        if (!tickets) {
          const res = await fetch('/api/jira/tickets?limit=50000&exclude_closed=true');
          if (res.ok) {
            const data = await res.json();
            tickets = data.tickets || [];
            authorizationCachedJiraTickets = tickets;
          } else {
            tickets = [];
          }
        }

        hideAuthorizationTypingIndicator();

        const normalizeText = str => String(str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
        const isAllTeams = !teamName || teamName.toLowerCase().includes('todas');
        const isAllAssignees = !assigneeName || assigneeName.toLowerCase().includes('todos');
        const isAllProcesses = !processName || processName.toLowerCase().includes('todos');

        const filtered = tickets.filter(t => {
          const matchTeam = isAllTeams || (t.team && (normalizeText(t.team) === normalizeText(teamName) || normalizeText(t.team).includes(normalizeText(teamName))));
          const matchAssignee = isAllAssignees || (t.assignee && (normalizeText(t.assignee) === normalizeText(assigneeName) || normalizeText(t.assignee).includes(normalizeText(assigneeName))));
          
          let matchProcess = true;
          if (!isAllProcesses) {
            const procNorm = normalizeText(processName);
            const tProc = normalizeText(t.process);
            const tType = normalizeText(t.ticket_type);
            const tSum = normalizeText(t.summary);
            matchProcess = tProc.includes(procNorm) || procNorm.includes(tProc) || tType.includes(procNorm) || procNorm.includes(tType) || tSum.includes(procNorm);
          }
          return matchTeam && matchAssignee && matchProcess;
        });

        const container = document.getElementById('authorization-chat-messages');
        if (!container) return;

        if (filtered.length === 0) {
          const processLabel = !isAllProcesses ? ` no processo **${escapeAuthorizationText(processName)}**` : '';
          const assigneeLabel = !isAllAssignees ? ` para **${escapeAuthorizationText(assigneeName)}**` : ` para a equipa **${escapeAuthorizationText(teamName)}**`;
          appendAuthorizationMessage(
            'assistant',
            `Não foram encontrados tickets abertos${assigneeLabel}${processLabel}.\n\nDeseja alterar o filtro ou ver outros responsáveis?`
          );

          const grid = document.createElement('div');
          grid.style.display = 'flex';
          grid.style.flexWrap = 'wrap';
          grid.style.gap = '10px';
          grid.style.marginTop = '6px';
          grid.style.marginBottom = '8px';

          const recoveryOptions = [
            { label: '⚙️ Ver outros processos', val: 'Ver processos' },
            { label: '👤 Ver responsáveis', val: 'Ver responsáveis' },
            { label: '🌐 Ver equipas Jira', val: 'Equipas Jira' }
          ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

          recoveryOptions.forEach(opt => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'auth-chat-system-card';
            btn.style.flex = '0 0 auto';
            btn.style.padding = '8px 12px';
            btn.onclick = () => {
              if (opt.val === 'Ver processos') {
                showJiraProcessOptions(teamName, assigneeName, true);
              } else if (opt.val === 'Ver responsáveis') {
                showJiraAssigneeOptions(teamName, true);
              } else {
                appendAuthorizationMessage('user', 'Ver equipas Jira');
                showJiraTeamOptions();
              }
            };
            btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(opt.label)}</span>`;
            grid.appendChild(btn);
          });
          container.appendChild(grid);
          container.scrollTop = container.scrollHeight;
          return;
        }

        const processLabelText = (!isAllProcesses && processName) ? ` no processo **${escapeAuthorizationText(processName)}**` : '';
        const assigneeLabelText = (!isAllAssignees && assigneeName) ? `para **${escapeAuthorizationText(assigneeName)}**` : `para a equipa **${escapeAuthorizationText(teamName)}**`;

        appendAuthorizationMessage(
          'assistant',
          `Encontrei **${filtered.length}** ticket(s) ${assigneeLabelText}${processLabelText}.\nSelecione um dos tickets abaixo para analisar:`
        );

        const grid = document.createElement('div');
        grid.style.display = 'flex';
        grid.style.flexDirection = 'column';
        grid.style.gap = '8px';
        grid.style.marginTop = '8px';
        grid.style.marginBottom = '12px';
        grid.style.width = '100%';

        const displayTickets = filtered.slice(0, 25);

        displayTickets.forEach(t => {
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'auth-chat-system-card';
          btn.style.width = '100%';
          btn.style.textAlign = 'left';
          btn.style.padding = '10px 14px';
          btn.style.display = 'flex';
          btn.style.alignItems = 'center';
          btn.style.justifyContent = 'space-between';
          btn.style.background = '#ffffff';
          btn.style.border = '1px solid #cbd5e1';
          btn.style.borderRadius = '8px';
          btn.style.boxShadow = '0 1px 3px rgba(0,0,0,0.05)';
          btn.style.cursor = 'pointer';
          btn.onclick = () => {
            if (btn.parentElement) {
              btn.parentElement.querySelectorAll('button').forEach(b => {
                b.classList.remove('selected');
                b.setAttribute('aria-pressed', 'false');
              });
            }
            btn.classList.add('selected');
            btn.setAttribute('aria-pressed', 'true');
            selectJiraTicket(t);
          };

          const summaryText = t.summary ? (t.summary.length > 60 ? t.summary.substring(0, 57) + '...' : t.summary) : 'Sem resumo';
          const assigneeText = t.assignee || 'Sem responsável';
          const keyText = t.key || 'TICKET';

          btn.innerHTML = `
            <div style="display:flex; flex-direction:column; gap:2px; overflow:hidden;">
              <span class="sys-code" style="font-size:0.84rem; font-weight:700; color:var(--primary, #3b82f6);">🎫 ${escapeAuthorizationText(keyText)} — ${escapeAuthorizationText(summaryText)}</span>
              <span style="font-size:0.76rem; color:var(--text-secondary, #64748b);">Responsável: ${escapeAuthorizationText(assigneeText)} | Estado: ${escapeAuthorizationText(t.status || 'Aberto')}</span>
            </div>
            <span style="font-size:0.8rem; font-weight:600; color:#3b82f6; white-space:nowrap; margin-left:8px;">Ver →</span>
          `;
          grid.appendChild(btn);
        });

        container.appendChild(grid);
        window.setTimeout(() => {
          container.scrollTop = container.scrollHeight;
        }, 50);
      } catch (err) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage('assistant', `⚠️ Erro ao obter os tickets do Jira: ${escapeAuthorizationText(err.message)}`);
      }
    }

    async function selectJiraTicket(ticketObj, skipUserMsg = false) {
      authorizationSelectedJiraTicket = ticketObj;
      authorizationChatState = AUTH_CHAT_STATES.WAITING_TICKET_ACTION;
      updateAuthorizationComposer();

      const key = ticketObj.key || 'Ticket';
      const summary = ticketObj.summary || '';
      if (!skipUserMsg) {
        appendAuthorizationMessage('user', `🎫 ${key}: ${summary}`);
      }

      showAuthorizationTypingIndicator(null, `A obter descrição do ticket ${escapeAuthorizationText(key)}...`);

      let descriptionText = ticketObj.description || '';
      let detailsObj = ticketObj;

      try {
        const res = await fetch(`/api/jira/tickets/${encodeURIComponent(key)}/details`, { cache: 'no-store' });
        if (res.ok) {
          const data = await res.json();
          if (data && data.description) {
            descriptionText = data.description;
          }
          detailsObj = { ...ticketObj, ...data };
          authorizationSelectedJiraTicket = detailsObj;
        }
      } catch (err) {
        console.warn('[TICKET DETAILS] Não foi possível carregar a descrição:', err);
      }

      hideAuthorizationTypingIndicator();

      const cleanDescription = descriptionText ? escapeAuthorizationText(descriptionText).replace(/\n/g, '<br>') : '<i>Sem descrição disponível</i>';

      appendAuthorizationMessage(
        'assistant',
        `
          <div class="auth-chat-summary">
            <div style="font-weight:700; margin-bottom:8px; color:var(--primary, #3b82f6); font-size:0.92rem;">
              🎫 Ticket ${escapeAuthorizationText(key)} Selecionado
            </div>
            <div style="display:grid; gap:6px; font-size:0.84rem; margin-bottom:12px;">
              <div><b>• Resumo:</b> ${escapeAuthorizationText(summary)}</div>
              <div style="background:#f8fafc; border:1px solid #e2e8f0; border-left:3px solid #3b82f6; border-radius:6px; padding:8px 10px; margin:4px 0;">
                <b style="color:#1e293b;">• Descrição / Erro:</b><br>
                <span style="color:#334155; display:block; margin-top:4px; max-height:180px; overflow-y:auto; font-size:0.82rem; line-height:1.4;">${cleanDescription}</span>
              </div>
              <div><b>• Responsável:</b> ${escapeAuthorizationText(detailsObj.assignee || 'Sem responsável')}</div>
              <div><b>• Estado:</b> ${escapeAuthorizationText(detailsObj.status || 'Aberto')}</div>
              ${detailsObj.team ? `<div><b>• Equipa:</b> ${escapeAuthorizationText(detailsObj.team)}</div>` : ''}
              ${detailsObj.ticket_type ? `<div><b>• Tipo de Ticket:</b> ${escapeAuthorizationText(detailsObj.ticket_type)}</div>` : ''}
            </div>
            <div style="font-size:0.82rem; color:var(--text-secondary);">
              Selecione uma das opções de ação abaixo para este ticket:
            </div>
          </div>
        `,
        true
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '6px';
      grid.style.marginBottom = '8px';

      const actions = [
        { label: '🔍 Analisar autorizações de utilizador', val: 'Analisar autorizações' },
        { label: '👤 Dados de utilizador', val: 'Dados de utilizador' },
        { label: '🛡️ Perfil de autorização', val: 'Perfil de autorização' },
        { label: '🔄 Executar outro processo', val: 'Processo' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      actions.forEach(item => {
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
          handleTicketActionSelect(item.val, detailsObj);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      window.setTimeout(() => {
        container.scrollTop = container.scrollHeight;
      }, 50);
    }

    function handleTicketActionSelect(actionVal, ticketObj) {
      if (actionVal === 'Analisar autorizações') {
        appendAuthorizationMessage('user', 'Analisar autorizações');
        authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
        updateAuthorizationComposer();
        appendAuthorizationMessage(
          'assistant',
          `Ação selecionada para o ticket **${escapeAuthorizationText(ticketObj?.key || '')}**.\nQual é o utilizador SAP que pretende analisar?`
        );
      } else if (actionVal === 'Dados de utilizador') {
        appendAuthorizationMessage('user', 'Dados de utilizador');
        showUserDataSubroutineOptions();
      } else if (actionVal === 'Perfil de autorização') {
        appendAuthorizationMessage('user', 'Perfil de autorização');
        showAuthorizationProfileSubroutineOptions();
      } else {
        appendAuthorizationMessage('user', 'Executar outro processo');
        renderAuthorizationInitialChoice();
      }
    }

    function renderRoutineSuggestionsInitial() {
      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage(
        'assistant',
        'Que processo ou rotina SAP pretende executar hoje?\nSelecione uma das sugestões abaixo ou escreva no campo inferior:'
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
        { label: '🔍 Cadeias de pesquisa', val: 'Cadeias de pesquisa' },
        { label: '🏦 Chave de banco', val: 'Chave de banco' },
        { label: '📋 Códigos IVA', val: 'Códigos IVA' },
        { label: '👤 Dados de utilizador', val: 'Dados de utilizador' },
        { label: '🧪 UAT Simulação', val: 'UAT Simulação' },
        { label: '🛡️ Perfil de autorização', val: 'Perfil de autorização' },
        { label: '🔄 Reverter documento', val: 'Reverter documento' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

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

    function showAuthorizationProfileSubroutineOptions() {
      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage(
        'assistant',
        'Perfeito. Selecionou a rotina **Perfil de autorização**.\nQual sub-rotina ou ação pretende executar?'
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
        { label: '🔍 Analisar autorizações', val: 'Analisar autorizações', type: 'ANALYSIS' },
        { label: '➕ Criar role composta', val: 'Criar role composta', scriptName: 'D. PFCG_COMPOSTA_RFC.py', category: 'Funções PFCG' },
        { label: '➕ Criar role simples', val: 'Criar role simples', scriptName: 'A. PFCG_CREATE_RFC.py', category: 'Funções PFCG' },
        { label: '❌ Eliminar role', val: 'Eliminar role', scriptName: 'B. PFCG_DELETE_RFC.py', category: 'Funções PFCG' },
        { label: '🛡️ Gerir objetos de autorização', val: 'Gerir objetos de autorização', scriptName: 'C. PFCG_AUTHORITY.py', category: 'Funções PFCG' },
        { label: '➖ Remover role de utilizador', val: 'Remover role de utilizador', scriptName: 'J. CUA_REMOVE.py', category: 'Funções PFCG' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

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
          selectAuthorizationProfileSubroutine(item);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function selectAuthorizationProfileSubroutine(item) {
      if (item.type === 'ANALYSIS' || item.val.includes('Analisar')) {
        appendAuthorizationMessage('user', item.val);
        authorizationIndividualContext = null;
        showAuthorizationTypingIndicator(null, 'A preparar o ambiente de análise de autorizações...');
        setTimeout(() => {
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage(
            'assistant',
            'Qual é o utilizador SAP que pretende analisar?'
          );
          authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
          updateAuthorizationComposer();
        }, 300);
        return;
      }

      selectUserDataSubroutine(item);
    }

    function showUserDataSubroutineOptions() {
      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage(
        'assistant',
        'Perfeito. Selecionou a rotina **Dados de utilizador**.\nQual sub-rotina pretende executar?'
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
        { label: '🔑 Alterar Senha', val: 'Alterar Senha', scriptName: 'su01_reset_password.py', category: 'CUA Login' },
        { label: 'âž• Criar utilizador', val: 'Criar utilizador', scriptName: 'L. CUA_CRIAR_USER.py', category: 'CUA_CRIAR_USER' },
        { label: '📅 Delimitar data fim', val: 'Delimitar data fim', scriptName: 'I. CUA_ENDDATE.py', category: 'CUA_ENDDATE' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

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
          selectUserDataSubroutine(item);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function showUatSimulationSubroutineOptions() {
      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage(
        'assistant',
        'Perfeito. Selecionou a pasta **UAT Simulação**.\nQual rotina pretende executar?'
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
        { label: '🧾 Criar Documento', val: 'Criar Documento', scriptName: 'Criar Documento.py', category: 'UAT Simulação' },
        { label: '⚙️ Executar F110', val: 'Executar F110', scriptName: 'Executar F110.py', category: 'UAT Simulação' },
        { label: '🧪 Simular F110', val: 'Simular F110', scriptName: 'simular_f110.py', category: 'UAT Simulação' },
        { label: '📄 Proposta Pagamento F110', val: 'Proposta Pagamento F110', scriptName: 'proposta_pagamento_f110.py', category: 'UAT Simulação' },
        { label: '🔁 RFF110S', val: 'RFF110S', scriptName: 'RFF110S.py', category: 'UAT Simulação' }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

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
          selectUatSimulationSubroutine(item);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    function getUatCreateDocumentDefaults() {
      const today = new Date();
      const y = today.getFullYear();
      const m = String(today.getMonth() + 1).padStart(2, '0');
      const d = String(today.getDate()).padStart(2, '0');
      const todayYyyyMmDd = `${y}${m}${d}`;

      return {
        system_key: 'QAD',
        company_code: '2010',
        vendor: '10000040',
        gl_account: '12010741',
        amount: '88,88',
        currency: 'EUR',
        document_date: todayYyyyMmDd,
        posting_date: todayYyyyMmDd,
        payment_method: 'S',
        doc_type: 'KR',
        reference: 'UAT-F110-TEST',
        header_text: 'UAT TESTE F110',
        item_text: 'UAT F110 TESTE',
      };
    }

    function formatUatDisplayDate(value) {
      const raw = String(value || '').trim();
      if (!raw) {
        return '';
      }

      const digits = raw.replace(/[^0-9]/g, '');
      if (/^\d{8}$/.test(digits)) {
        return `${digits.slice(6, 8)}.${digits.slice(4, 6)}.${digits.slice(0, 4)}`;
      }

      const parsed = new Date(raw);
      if (!Number.isNaN(parsed.getTime())) {
        const day = String(parsed.getDate()).padStart(2, '0');
        const month = String(parsed.getMonth() + 1).padStart(2, '0');
        const year = String(parsed.getFullYear());
        return `${day}.${month}.${year}`;
      }

      return raw;
    }

    function getUatExecuteF110Schedule() {
      const today = new Date();
      const todayYyyyMmDd = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
      const tomorrow = new Date(today.getTime());
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowYyyyMmDd = `${tomorrow.getFullYear()}${String(tomorrow.getMonth() + 1).padStart(2, '0')}${String(tomorrow.getDate()).padStart(2, '0')}`;

      return {
        runDate: todayYyyyMmDd,
        docsEnteredUpTo: tomorrowYyyyMmDd,
        runDateDisplay: formatUatDisplayDate(todayYyyyMmDd),
        docsEnteredUpToDisplay: formatUatDisplayDate(tomorrowYyyyMmDd),
      };
    }

    function buildUatExecuteF110PreparationLabel(context, schedule) {
      const defaults = getUatCreateDocumentDefaults();
      const companyCode = String(context?.companyCode || defaults.company_code || '').trim();
      const vendor = String(context?.vendor || defaults.vendor || '').trim();
      const fiscalYear = String(context?.fiscalYear || new Date().getFullYear()).trim();
      const documentNumber = String(context?.documentNumber || '').trim();
      const runDateDisplay = String(schedule?.runDateDisplay || formatUatDisplayDate(getUatExecuteF110Schedule().runDate) || '').trim();

      return [
        'A preparar a execução completa da F110:',
        `• Run date: ${escapeAuthorizationText(runDateDisplay)}`,
        '• Identificação: sequencial automática',
        `• Documento SAP: ${escapeAuthorizationText(documentNumber)}`,
        `• Empresa: ${escapeAuthorizationText(companyCode)}`,
        `• Fornecedor: ${escapeAuthorizationText(vendor)}`,
        `• Exercicio: ${escapeAuthorizationText(fiscalYear)}`,
      ].join('<br>');
    }

    async function resolveLatestUatExecuteF110Identification() {
      try {
        const response = await fetch('/api/jobs?include_archived=true&limit=50', { cache: 'no-store' });
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`);
        }

        const payload = await response.json();
        const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
        for (const job of jobs) {
          const params = job?.params || {};
          if (String(params.subprocesso || '').trim() !== 'Executar F110.py') {
            continue;
          }

          const paramIdentification = String(params.identification || '').trim().toUpperCase();
          if (paramIdentification && paramIdentification !== 'AUTO') {
            return paramIdentification;
          }

          const log = String(job?.log || '');
          const match = log.match(/^\s*Identifica(?:c|ç)ão\s*:\s*([A-Z0-9]+)\s*$/im) || log.match(/^\s*Identificacao\s*:\s*([A-Z0-9]+)\s*$/im);
          if (match?.[1]) {
            return String(match[1]).trim().toUpperCase();
          }
        }
      } catch (error) {
        console.warn('Falha a resolver a identificação da F110:', error);
      }

      return 'UAT01';
    }

    const UAT_CREATE_DOCUMENT_STEPS = [
      { key: 'company_code', label: 'Código da empresa (BUKRS)', defaultValue: '2010', uppercase: true },
      { key: 'vendor', label: 'Fornecedor', defaultValue: '10000040', uppercase: true },
      { key: 'gl_account', label: 'Conta GL', defaultValue: '12010741', uppercase: true },
      { key: 'amount', label: 'Valor do documento', defaultValue: '88,88' },
      { key: 'currency', label: 'Moeda', defaultValue: 'EUR', uppercase: true },
      { key: 'document_date', label: 'Data do documento', defaultValue: null, date: true },
      { key: 'posting_date', label: 'Data de lançamento', defaultValue: null, date: true },
      { key: 'payment_method', label: 'Método de pagamento', defaultValue: 'S', uppercase: true },
      { key: 'doc_type', label: 'Tipo de documento', defaultValue: 'KR', uppercase: true },
      { key: 'reference', label: 'Referência externa', defaultValue: 'UAT-F110-TEST' },
      { key: 'header_text', label: 'Texto do cabeçalho', defaultValue: 'UAT TESTE F110' },
      { key: 'item_text', label: 'Texto das linhas', defaultValue: 'UAT F110 TESTE' },
    ];

    function normalizeUatCreateDocumentText(value, uppercase = false) {
      const text = String(value || '').trim();
      return uppercase ? text.toUpperCase() : text;
    }

    function normalizeUatCreateDocumentDate(value, fallbackValue) {
      const raw = String(value || '').trim();
      if (!raw) {
        return fallbackValue || '';
      }

      const digits = raw.replace(/[^0-9]/g, '');
      if (/^\d{8}$/.test(digits)) {
        return digits;
      }

      const slashMatch = raw.match(/^(\d{2})[\/.-](\d{2})[\/.-](\d{4})$/);
      if (slashMatch) {
        return `${slashMatch[3]}${slashMatch[2]}${slashMatch[1]}`;
      }

      return raw;
    }

    function resetUatCreateDocumentFlow() {
      authorizationUatCreateDocumentFlow = null;
    }

    function appendUatCreateDocumentPrompt() {
      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage(
        'assistant',
        'Pretende usar os **Dados Default** ou introduzir **Novos dados** para a criação do documento?'
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '6px';
      grid.style.marginBottom = '8px';

      const options = [
        { label: '📋 Dados Default', val: 'default' },
        { label: '✍️ Novos dados', val: 'new' },
      ];

      options.forEach(item => {
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
          selectUatCreateDocumentMode(item.val);
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.82rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

    async function submitUatCreateDocumentJob(params) {
      const defaults = getUatCreateDocumentDefaults();
      const payload = {
        task: 'sap_cockpit',
        ambiente: defaults.system_key,
        processo: 'UAT Simulação',
        subprocesso: 'Criar Documento.py',
        request_option: '4',
        request_type: '1',
        caminho_ficheiro: '',
        nome_pasta: '',
        system_key: defaults.system_key,
        ...defaults,
        ...params,
      };

      const formData = new FormData();
      Object.entries(payload).forEach(([key, value]) => {
        formData.set(key, value == null ? '' : String(value));
      });

      const response = await fetch('/jobs', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const detail = await response.text().catch(() => '');
        throw new Error(detail || `Falha ao criar o job. HTTP ${response.status}`);
      }

      const job = await response.json();
      if (typeof loadJobs === 'function') {
        loadJobs().catch(() => {});
      }
      return job;
    }

    function parseUatCreateDocumentJobSummary(job) {
      const log = String(job?.log || '');
      const state = String(job?.state || '').trim() || 'unknown';
      const status = String(job?.status || '').trim() || 'N/D';

      const docMatch = log.match(/\bBELNR\s*:\s*([0-9]+)/i);
      const companyMatch = log.match(/\bBUKRS\s*:\s*([0-9]+)/i);
      const yearMatch = log.match(/\bGJAHR\s*:\s*([0-9]+)/i);
      const refMatch = log.match(/\bReferencia\s*:\s*(.+)/i);

      return {
        state,
        status,
        documentNumber: String(docMatch?.[1] || '').trim(),
        companyCode: String(companyMatch?.[1] || '').trim(),
        fiscalYear: String(yearMatch?.[1] || '').trim(),
        reference: String(refMatch?.[1] || '').trim(),
      };
    }

    function buildUatCreateDocumentFinalHtml(job) {
      const summary = parseUatCreateDocumentJobSummary(job);
      const lines = [
        `state: ${escapeAuthorizationText(summary.state)}`,
        `status: ${escapeAuthorizationText(summary.status)}`,
      ];

      if (summary.documentNumber) {
        lines.push(`Documento SAP criado: ${escapeAuthorizationText(summary.documentNumber)}`);
      }
      if (summary.companyCode) {
        lines.push(`Empresa: ${escapeAuthorizationText(summary.companyCode)}`);
      }
      if (summary.fiscalYear) {
        lines.push(`Exercicio: ${escapeAuthorizationText(summary.fiscalYear)}`);
      }
      if (summary.reference) {
        lines.push(`Referencia no SAP: ${escapeAuthorizationText(summary.reference)}`);
      }

      return lines.map(line => `&bull; ${line}`).join('<br>');
    }

    function appendUatCreateDocumentFinalResultAndPrompt(job) {
      rememberUatCreateDocumentContext(job);
      const finalHtml = buildUatCreateDocumentFinalHtml(job);
      appendAuthorizationMessage('assistant', finalHtml, true);
      showUatSimulationSubroutineOptions();
    }

    function rememberUatCreateDocumentContext(job) {
      const summary = parseUatCreateDocumentJobSummary(job);
      const flowValues = authorizationUatCreateDocumentFlow?.values || {};

      authorizationUatLastCreatedDocumentContext = {
        documentNumber: String(summary.documentNumber || '').trim(),
        fiscalYear: String(summary.fiscalYear || flowValues.posting_date?.slice?.(0, 4) || new Date().getFullYear()).trim(),
        companyCode: String(flowValues.company_code || '').trim(),
        vendor: String(flowValues.vendor || '').trim(),
        systemKey: String(flowValues.system_key || getUatCreateDocumentDefaults().system_key).trim(),
        reference: String(flowValues.reference || '').trim(),
        glAccount: String(flowValues.gl_account || '').trim(),
        amount: String(flowValues.amount || '').trim(),
        currency: String(flowValues.currency || '').trim(),
        paymentMethod: String(flowValues.payment_method || '').trim(),
        docType: String(flowValues.doc_type || '').trim(),
        documentDate: String(flowValues.document_date || '').trim(),
        postingDate: String(flowValues.posting_date || '').trim(),
      };
    }

    function extractUatLogField(log, label) {
      const text = String(log || '');
      const pattern = new RegExp(`^\\s*${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*:\\s*(.+?)\\s*$`, 'im');
      const match = text.match(pattern);
      return String(match?.[1] || '').trim();
    }

    function parseUatExecuteF110JobSummary(job) {
      const log = String(job?.log || '');
      const params = job?.params || {};
      return {
        state: String(job?.state || '').trim() || 'unknown',
        companyCode: extractUatLogField(log, 'Empresa'),
        vendor: extractUatLogField(log, 'Fornecedor'),
        documentNumber: extractUatLogField(log, 'Documento SAP utilizado') || extractUatLogField(log, 'Documento utilizado') || extractUatLogField(log, 'Documento novo') || extractUatLogField(log, 'Documento'),
        fiscalYear: extractUatLogField(log, 'Exercicio'),
        identification: extractUatLogField(log, 'Identificacao') || extractUatLogField(log, 'Identificação') || extractUatLogField(log, 'Identificação'),
        paymentDocumentNumber: extractUatLogField(log, 'Documento de pagamento') || extractUatLogField(log, 'AUGBL'),
        runDate: String(params.run_date || '').trim(),
      };
    }

    function buildUatExecuteF110FinalHtml(job) {
      const summary = parseUatExecuteF110JobSummary(job);
      const lines = [
        ['state', summary.state],
        ['Run date', summary.runDate ? formatUatDisplayDate(summary.runDate) : ''],
        ['Identificacao', summary.identification],
        ['Documento SAP utilizado', summary.documentNumber],
        ['Empresa', summary.companyCode],
        ['Fornecedor', summary.vendor],
        ['Exercicio', summary.fiscalYear],
        ['Documento de pagamento', summary.paymentDocumentNumber],
      ];

      return lines
        .filter(([, value]) => String(value || '').trim())
        .map(([label, value]) => `&bull; <b>${escapeAuthorizationText(label)}:</b> ${escapeAuthorizationText(String(value || '').trim())}`)
        .join('<br>');
    }

    async function submitUatExecuteF110Job(context) {
      const defaults = getUatCreateDocumentDefaults();
      const schedule = getUatExecuteF110Schedule();
      const payload = {
        task: 'sap_cockpit',
        ambiente: String(context?.systemKey || defaults.system_key || 'QAD').trim(),
        processo: 'UAT Simulação',
        subprocesso: 'Executar F110.py',
        request_option: '4',
        request_type: '1',
        caminho_ficheiro: '',
        nome_pasta: '',
        system_key: String(context?.systemKey || defaults.system_key || 'QAD').trim(),
        company_code: String(context?.companyCode || defaults.company_code || '').trim(),
        vendor: String(context?.vendor || defaults.vendor || '').trim(),
        gl_account: String(context?.glAccount || defaults.gl_account || '').trim(),
        amount: String(context?.amount || defaults.amount || '').trim(),
        currency: String(context?.currency || defaults.currency || '').trim(),
        document_date: String(context?.documentDate || '').trim(),
        posting_date: String(context?.postingDate || '').trim(),
        payment_method: String(context?.paymentMethod || defaults.payment_method || 'S').trim(),
        doc_type: String(context?.docType || defaults.doc_type || 'KR').trim(),
        reference: String(context?.reference || defaults.reference || '').trim(),
        header_text: String(context?.headerText || defaults.header_text || '').trim(),
        item_text: String(context?.itemText || defaults.item_text || '').trim(),
        document_number: String(context?.documentNumber || '').trim(),
        fiscal_year: String(context?.fiscalYear || '').trim(),
        run_date: schedule.runDate,
        docs_entered_up_to: schedule.docsEnteredUpTo,
        step: 'full',
        skip_create_document: 'true',
        proposal_only: 'false',
        execute_payment: 'true',
        wait_seconds: '120',
      };

      const formData = new FormData();
      Object.entries(payload).forEach(([key, value]) => {
        formData.set(key, value == null ? '' : String(value));
      });

      const response = await fetch('/jobs', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const detail = await response.text().catch(() => '');
        throw new Error(detail || `Falha ao criar o job. HTTP ${response.status}`);
      }

      const job = await response.json();
      if (typeof loadJobs === 'function') {
        loadJobs().catch(() => {});
      }
      return job;
    }

    async function pollUatExecuteF110Job(jobId, requestId) {
      const startTime = Date.now();
      const timeoutMs = 300000;

      async function check() {
        if (requestId !== authorizationUatExecuteF110JobRequestId) {
          return;
        }

        if (Date.now() - startTime > timeoutMs) {
          if (requestId === authorizationUatExecuteF110JobRequestId) {
            hideAuthorizationTypingIndicator();
            appendAuthorizationMessage(
              'assistant',
              `O fluxo completo da F110 ainda está a ser processado no worker. Job #${escapeAuthorizationText(String(jobId || '').slice(0, 8))}.`
            );
            showUatSimulationSubroutineOptions();
          }
          return;
        }

        try {
          const response = await fetchWithTimeout(`/api/jobs/${jobId}`, {}, 10000);
          if (!response.ok) {
            throw new Error(`Erro HTTP ${response.status}`);
          }

          const job = await response.json();
          if (requestId !== authorizationUatExecuteF110JobRequestId) {
            return;
          }

          if (!isAuthorizationTerminalJobState(job.state)) {
            window.setTimeout(check, 2000);
            return;
          }

          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', buildUatExecuteF110FinalHtml(job), true);
          showUatSimulationSubroutineOptions();
        } catch (error) {
          if (requestId !== authorizationUatExecuteF110JobRequestId) {
            return;
          }

          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage(
            'assistant',
            `⚠️ Não foi possível obter o resultado final da F110: ${escapeAuthorizationText(error?.message || 'erro desconhecido')}`
          );
          showUatSimulationSubroutineOptions();
        }
      }

      window.setTimeout(check, 1000);
    }

    async function finalizeUatExecuteF110Flow() {
      const context = authorizationUatLastCreatedDocumentContext;
      if (!context || !String(context.documentNumber || '').trim()) {
        appendAuthorizationMessage(
          'assistant',
          'Não encontro um documento SAP válido gerado anteriormente neste chat. Primeiro execute `Criar Documento` para gerar o número e depois volte a `Executar F110`.'
        );
        showUatSimulationSubroutineOptions();
        return;
      }

      const currentRequestId = ++authorizationUatExecuteF110JobRequestId;
      authorizationChatState = AUTH_CHAT_STATES.LOADING;
      updateAuthorizationComposer();
      const schedule = getUatExecuteF110Schedule();
      showAuthorizationTypingIndicator(currentRequestId, buildUatExecuteF110PreparationLabel(context, schedule));

      try {
        const job = await submitUatExecuteF110Job(context);
        const jobId = String(job?.id || '').trim();
        if (!jobId) {
          throw new Error('Job sem identificador');
        }

        showAuthorizationTypingIndicator(currentRequestId, buildUatExecuteF110PreparationLabel(context, schedule));
        await pollUatExecuteF110Job(jobId, currentRequestId);
      } catch (error) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `⚠️ Não foi possível executar o fluxo completo da F110: ${escapeAuthorizationText(error?.message || 'erro desconhecido')}`
        );
        showUatSimulationSubroutineOptions();
      } finally {
        authorizationChatState = AUTH_CHAT_STATES.READY;
        updateAuthorizationComposer();
      }
    }

    async function pollUatCreateDocumentJob(jobId, requestId) {
      const startTime = Date.now();
      const timeoutMs = 180000;

      async function check() {
        if (requestId !== authorizationUatCreateDocumentJobRequestId) {
          return;
        }

        if (Date.now() - startTime > timeoutMs) {
          if (requestId === authorizationUatCreateDocumentJobRequestId) {
            hideAuthorizationTypingIndicator();
            appendAuthorizationMessage(
              'assistant',
              `O documento ainda está a ser processado no worker. Job #${escapeAuthorizationText(String(jobId || '').slice(0, 8))}.`
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
          if (requestId !== authorizationUatCreateDocumentJobRequestId) {
            return;
          }

          if (!isAuthorizationTerminalJobState(job.state)) {
            window.setTimeout(check, 2000);
            return;
          }

          hideAuthorizationTypingIndicator();
          appendUatCreateDocumentFinalResultAndPrompt(job);
        } catch (error) {
          if (requestId !== authorizationUatCreateDocumentJobRequestId) {
            return;
          }

          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage(
            'assistant',
            `⚠️ Não foi possível obter o resultado final do documento: ${escapeAuthorizationText(error?.message || 'erro desconhecido')}`
          );
        }
      }

      window.setTimeout(check, 1000);
    }

    function getUatCreateDocumentStepPrompt(step, currentValue = '') {
      const defaultValue = String(currentValue || step.defaultValue || '').trim();
      const suffix = defaultValue ? `\nValor sugerido: ${defaultValue}. Prima Enter para aceitar.` : '';
      return `Informe ${step.label}.${suffix}`;
    }

    function resolveUatCreateDocumentFieldValue(step, rawValue, currentValues = {}) {
      const fallbackValue = currentValues?.[step.key] ?? step.defaultValue ?? '';
      const trimmed = String(rawValue || '').trim();
      const baseValue = trimmed || fallbackValue || '';

      if (step.date) {
        return normalizeUatCreateDocumentDate(baseValue, fallbackValue);
      }

      return normalizeUatCreateDocumentText(baseValue, Boolean(step.uppercase));
    }

    function formatUatCreateDocumentSummary(values) {
      return [
        `• Empresa: ${escapeAuthorizationText(values.company_code || '')}`,
        `• Fornecedor: ${escapeAuthorizationText(values.vendor || '')}`,
        `• Conta GL: ${escapeAuthorizationText(values.gl_account || '')}`,
        `• Valor: ${escapeAuthorizationText(values.amount || '')} ${escapeAuthorizationText(values.currency || '')}`,
        `• Data doc.: ${escapeAuthorizationText(values.document_date || '')}`,
        `• Data lanc.: ${escapeAuthorizationText(values.posting_date || '')}`,
        `• Metodo: ${escapeAuthorizationText(values.payment_method || '')}`,
        `• Tipo doc.: ${escapeAuthorizationText(values.doc_type || '')}`,
        `• Referencia: ${escapeAuthorizationText(values.reference || '')}`,
      ].join('<br>');
    }

    function askNextUatCreateDocumentStep() {
      const flow = authorizationUatCreateDocumentFlow;
      if (!flow || !flow.active || flow.mode !== 'new') {
        return;
      }

      const step = UAT_CREATE_DOCUMENT_STEPS[flow.stepIndex || 0];
      if (!step) {
        void finalizeUatCreateDocumentFlow();
        return;
      }

      hideAuthorizationTypingIndicator();
      appendAuthorizationMessage('assistant', getUatCreateDocumentStepPrompt(step, flow.values?.[step.key]));
      updateAuthorizationComposer();

      const inputEl = document.getElementById('authorization-chat-input');
      if (inputEl) inputEl.focus();
    }

    async function finalizeUatCreateDocumentFlow() {
      const flow = authorizationUatCreateDocumentFlow;
      if (!flow) return;

      const values = flow.values || {};
      const currentRequestId = ++authorizationUatCreateDocumentJobRequestId;
      showAuthorizationTypingIndicator(currentRequestId, 'A aguardar o resultado final do documento...');

      try {
        const job = await submitUatCreateDocumentJob(values);
        const jobId = String(job?.id || '').trim();
        if (!jobId) {
          throw new Error('Job sem identificador');
        }

        showAuthorizationTypingIndicator(currentRequestId, 'A acompanhar o processamento final do documento...');
        await pollUatCreateDocumentJob(jobId, currentRequestId);
      } catch (error) {
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `⚠️ Não foi possível criar o documento: ${escapeAuthorizationText(error?.message || 'erro desconhecido')}`
        );
      } finally {
        resetUatCreateDocumentFlow();
        authorizationChatState = AUTH_CHAT_STATES.READY;
        updateAuthorizationComposer();
      }
    }

    function selectUatCreateDocumentMode(mode) {
      const normalizedMode = String(mode || '').trim().toLowerCase();
      if (!authorizationUatCreateDocumentFlow) {
        authorizationUatCreateDocumentFlow = {
          active: true,
          mode: 'default',
          stepIndex: 0,
          values: getUatCreateDocumentDefaults(),
        };
      }

      if (normalizedMode === 'default') {
        appendAuthorizationMessage('user', 'Dados Default');
        authorizationUatCreateDocumentFlow.mode = 'default';
        void finalizeUatCreateDocumentFlow();
        return;
      }

      appendAuthorizationMessage('user', 'Novos dados');
      authorizationUatCreateDocumentFlow.mode = 'new';
      authorizationUatCreateDocumentFlow.stepIndex = 0;
      authorizationUatCreateDocumentFlow.values = {
        ...getUatCreateDocumentDefaults(),
      };
      authorizationChatState = AUTH_CHAT_STATES.READY;
      askNextUatCreateDocumentStep();
    }

    function handleUatCreateDocumentChatSubmit(rawValue) {
      const flow = authorizationUatCreateDocumentFlow;
      if (!flow || !flow.active) {
        return false;
      }

      const normValue = normalizeAuthorizationSearchText(rawValue);

      if (flow.mode === 'pending') {
        if (normValue.includes('DEFAULT') || normValue.includes('DADOS DEFAULT')) {
          selectUatCreateDocumentMode('default');
          return true;
        }
        if (normValue.includes('NOVOS DADOS') || normValue.includes('NOVOS') || normValue.includes('NOVO')) {
          selectUatCreateDocumentMode('new');
          return true;
        }
        appendAuthorizationMessage(
          'assistant',
          'Escolha uma das opções: **Dados Default** ou **Novos dados**.'
        );
        appendUatCreateDocumentPrompt();
        return true;
      }

      if (flow.mode !== 'new') {
        return false;
      }

      const step = UAT_CREATE_DOCUMENT_STEPS[flow.stepIndex || 0];
      if (!step) {
        return false;
      }

      const resolvedValue = resolveUatCreateDocumentFieldValue(step, rawValue, flow.values || {});
      flow.values[step.key] = resolvedValue;

      const displayValue = resolvedValue || '(vazio)';
      appendAuthorizationMessage('user', `${step.label}: ${displayValue}`);

      flow.stepIndex = (flow.stepIndex || 0) + 1;

      if (flow.stepIndex >= UAT_CREATE_DOCUMENT_STEPS.length) {
        void finalizeUatCreateDocumentFlow();
      } else {
        askNextUatCreateDocumentStep();
      }

      return true;
    }

    function selectUatSimulationSubroutine(item) {
      appendAuthorizationMessage('user', item.val);

      if (String(item.val || '').trim().toLowerCase() === 'criar documento') {
        authorizationUatCreateDocumentFlow = {
          active: true,
          mode: 'pending',
          stepIndex: 0,
          values: getUatCreateDocumentDefaults(),
        };
        authorizationChatState = AUTH_CHAT_STATES.READY;
        appendUatCreateDocumentPrompt();
        return;
      }

      if (String(item.val || '').trim().toLowerCase() === 'executar f110') {
        authorizationChatState = AUTH_CHAT_STATES.READY;
        void finalizeUatExecuteF110Flow();
        return;
      }

      // As restantes rotinas UAT continuam a usar o formulário direto do cockpit.
      abrirSubprocessoModal(item.category, item.scriptName);

      appendAuthorizationMessage(
        'assistant',
        `Formulário de **${escapeAuthorizationText(item.val)}** aberto.\n\nPreencha os campos necessários no modal e clique em **Executar** para submeter o job.`
      );
    }

    function selectUserDataSubroutine(item) {
      appendAuthorizationMessage('user', item.val);
      promptProcessMode(item.val, item.category, item.scriptName, item.label);
    }

    async function ensureAuthorizationInitialQuestion() {
      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      if (!container.querySelector('[data-message-key="authorization-initial-question"]')) {
        renderAuthorizationInitialChoice();
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
      authorizationSelectedJiraTeam = null;
      authorizationSelectedJiraAssignee = null;
      authorizationSelectedJiraProcess = null;
      authorizationSelectedJiraTicket = null;
      resetUatCreateDocumentFlow();

      const container = document.getElementById('authorization-chat-messages');
      const input = document.getElementById('authorization-chat-input');
      if (container) container.innerHTML = '';
      if (input) input.value = '';

      removeAuthorizationTypingIndicator();
      setAuthorizationChatState(AUTH_CHAT_STATES.WAITING_INITIAL_CHOICE);
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
      'USER', 'USERS', 'UTILIZADOR', 'UTILIZADORES', 'CONTA', 'ID', 'SESSAO', 'SESSÃO',
      'UMA', 'UM', 'DAS', 'DOS', 'DE', 'DO', 'DA', 'FUNCAO', 'FUNCOES', 'FUNÇÕES',
      'EXPIRADA', 'EXPIRADAS', 'EXPIRADO', 'EXPIRADOS', 'VENCIDA', 'VENCIDAS',
      'ATIVA', 'ATIVAS', 'ATIVO', 'ATIVOS', 'DIRETA', 'DIRETAS', 'COMPOSTA', 'COMPOSTAS',
      'MOSTRA', 'MOSTRAR', 'EXIBIR', 'FILTRAR', 'SO', 'SÓ', 'APENAS'
    ]);

    function renderRolesTableInChat(roles) {
      if (!Array.isArray(roles)) roles = [];

      let rowsHtml = '';
      roles.forEach((r, idx) => {
        const validityFrom = r.valid_from || r.validity_from || '(aberta)';
        const validityTo = r.valid_to || r.validity_to || '(aberta)';

        let statusBadge = '';
        if (r.validity_status === 'active' || r.isActive || r.status === 'ATIVA' || String(r.status || '').toUpperCase().includes('ATIV')) {
          statusBadge = '<span class="badge badge-success" style="background:#dcfce7; color:#15803d; font-size:0.75rem; font-weight:700; padding:2px 6px; border-radius:4px;">ATIVA</span>';
        } else if (r.validity_status === 'expired' || r.isExpired || r.status === 'EXPIRADA' || String(r.status || '').toUpperCase().includes('EXPIRAD')) {
          statusBadge = '<span class="badge badge-danger" style="background:#fee2e2; color:#b91c1c; font-size:0.75rem; font-weight:700; padding:2px 6px; border-radius:4px;">EXPIRADA</span>';
        } else if (r.validity_status === 'future') {
          statusBadge = '<span class="badge badge-warning" style="background:#fef3c7; color:#b45309; font-size:0.75rem; font-weight:700; padding:2px 6px; border-radius:4px;">FUTURA</span>';
        } else {
          statusBadge = '<span class="badge" style="font-size:0.75rem; font-weight:600;">INDETERMINADA</span>';
        }

        const isHidden = idx >= 15 ? 'class="auth-row-hidden" style="display:none;"' : '';
        const roleName = r.role || r.AGR_NAME || r.name || '';
        const originLabel = r.assignment_origin_label || r.assignment_origin || r.origin || r.atribuicao || 'Direta';

        rowsHtml += `
          <tr ${isHidden}>
            <td style="padding:6px 10px; border-bottom:1px solid #f1f5f9;"><strong>${escapeAuthorizationText(roleName)}</strong></td>
            <td style="padding:6px 10px; border-bottom:1px solid #f1f5f9;">${escapeAuthorizationText(validityFrom)}</td>
            <td style="padding:6px 10px; border-bottom:1px solid #f1f5f9;">${escapeAuthorizationText(validityTo)}</td>
            <td style="padding:6px 10px; border-bottom:1px solid #f1f5f9;">${statusBadge}</td>
            <td style="padding:6px 10px; border-bottom:1px solid #f1f5f9;">${escapeAuthorizationText(originLabel)}</td>
          </tr>
        `;
      });

      if (roles.length === 0) {
        rowsHtml = `<tr><td colspan="5" style="text-align: center; padding:12px; color: var(--text-muted);">Nenhuma role encontrada para os critérios selecionados.</td></tr>`;
      }

      let verTodasBtn = '';
      if (roles.length > 15) {
        verTodasBtn = `
          <button class="auth-ver-todas-btn" style="margin-top:8px; cursor:pointer; background:none; border:none; color:#2563eb; font-weight:600; font-size:0.82rem;" onclick="
            const table = this.closest('.auth-table-wrapper');
            table.querySelectorAll('.auth-row-hidden').forEach(row => row.style.display = '');
            this.style.display = 'none';
          ">Ver todas (${roles.length - 15} mais)</button>
        `;
      }

      const tableHtml = `
        <div class="auth-table-wrapper" style="margin-top:8px; margin-bottom:8px; width:100%; overflow-x:auto;">
          <table class="auth-roles-table" style="width:100%; border-collapse:collapse; font-size:0.84rem;">
            <thead>
              <tr style="background:#f8fafc; border-bottom:2px solid #e2e8f0; text-align:left;">
                <th style="padding:8px 10px;">Função / Role</th>
                <th style="padding:8px 10px;">Início Validade</th>
                <th style="padding:8px 10px;">Fim Validade</th>
                <th style="padding:8px 10px;">Estado</th>
                <th style="padding:8px 10px;">Atribuição</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>
          ${verTodasBtn}
        </div>
      `;

      appendAuthorizationMessage('assistant', tableHtml, true);
    }

    function handleContextualRolesQuery(rawVal, normVal) {
      if (!Array.isArray(authorizationLastDisplayedRoles) || authorizationLastDisplayedRoles.length === 0) {
        return false;
      }

      const isQueryExpired = normVal.includes('EXPIRAD') || normVal.includes('VENCID');
      const isQueryActive = normVal.includes('ATIV') || normVal.includes('VALID');
      const isQueryDirect = normVal.includes('DIRET');
      const isQueryComposta = normVal.includes('COMPOST');
      const isQueryCount = normVal.includes('QUANT') || normVal.includes('CONTAGEM') || normVal.includes('TOTAL');

      if (!isQueryExpired && !isQueryActive && !isQueryDirect && !isQueryComposta && !isQueryCount) {
        return false;
      }

      hideAuthorizationTypingIndicator();

      let filteredRoles = [...authorizationLastDisplayedRoles];
      let titleMsg = '';

      if (isQueryExpired) {
        filteredRoles = filteredRoles.filter(r => r.isExpired || r.validity_status === 'expired' || String(r.status || '').toUpperCase().includes('EXPIRAD'));
        titleMsg = `📋 **Lista de Funções Expiradas (${filteredRoles.length}) para o utilizador ${escapeAuthorizationText(authorizationTargetUser || '')}:**`;
      } else if (isQueryActive) {
        filteredRoles = filteredRoles.filter(r => r.isActive || r.validity_status === 'active' || String(r.status || '').toUpperCase().includes('ATIV'));
        titleMsg = `📋 **Lista de Funções Ativas (${filteredRoles.length}) para o utilizador ${escapeAuthorizationText(authorizationTargetUser || '')}:**`;
      } else if (isQueryDirect) {
        filteredRoles = filteredRoles.filter(r => (r.origin || r.atribuicao || r.assignment_origin_label || '').toLowerCase().includes('diret'));
        titleMsg = `📋 **Lista de Atribuições Diretas (${filteredRoles.length}) para o utilizador ${escapeAuthorizationText(authorizationTargetUser || '')}:**`;
      } else if (isQueryComposta) {
        filteredRoles = filteredRoles.filter(r => (r.origin || r.atribuicao || r.assignment_origin_label || '').toLowerCase().includes('compost'));
        titleMsg = `📋 **Lista de Roles Compostas (${filteredRoles.length}) para o utilizador ${escapeAuthorizationText(authorizationTargetUser || '')}:**`;
      } else if (isQueryCount) {
        const expiredCount = authorizationLastDisplayedRoles.filter(r => r.isExpired || r.validity_status === 'expired' || String(r.status || '').toUpperCase().includes('EXPIRAD')).length;
        const activeCount = authorizationLastDisplayedRoles.filter(r => r.isActive || r.validity_status === 'active' || String(r.status || '').toUpperCase().includes('ATIV')).length;
        appendAuthorizationMessage(
          'assistant',
          `📊 **Resumo de Funções para ${escapeAuthorizationText(authorizationTargetUser || '')}:**\n` +
          `• Total de funções: **${authorizationLastDisplayedRoles.length}**\n` +
          `• Funções Ativas: **${activeCount}**\n` +
          `• Funções Expiradas: **${expiredCount}**`
        );
        showNextActionsPrompt('Deseja efetuar mais alguma ação ou filtrar outra lista?');
        return true;
      }

      if (filteredRoles.length === 0) {
        appendAuthorizationMessage(
          'assistant',
          `ℹ️ Não foram encontradas funções correspondentes aos critérios pesquisados para o utilizador **${escapeAuthorizationText(authorizationTargetUser || '')}**.`
        );
        showNextActionsPrompt('Deseja fazer nova pesquisa ou escolher outra rotina?');
      } else {
        appendAuthorizationMessage('assistant', titleMsg);
        renderRolesTableInChat(filteredRoles);
        showFilteredRolesActionsPrompt(filteredRoles, titleMsg);
      }
      return true;
    }

    function showFilteredRolesActionsPrompt(filteredRoles, filterTypeLabel) {
      hideAuthorizationTypingIndicator();

      appendAuthorizationMessage(
        'assistant',
        `Pretende efetuar alguma ação (como **CUA_ENDDATE** ou **CUA_REMOVE**) sobre as **${filteredRoles.length}** funções de **${escapeAuthorizationText(authorizationTargetUser || '')}** apresentadas na lista acima?`
      );

      const container = document.getElementById('authorization-chat-messages');
      if (!container) return;

      const grid = document.createElement('div');
      grid.style.display = 'flex';
      grid.style.flexWrap = 'wrap';
      grid.style.gap = '10px';
      grid.style.marginTop = '8px';
      grid.style.marginBottom = '14px';

      const actions = [
        {
          label: '📅 Delimitar data fim (CUA_ENDDATE)',
          val: 'Delimitar data fim (CUA_ENDDATE)',
          action: () => {
            appendAuthorizationMessage('user', 'Delimitar data fim (CUA_ENDDATE)');
            confirmAuthorizationEndDate(
              authorizationTargetUser,
              authorizationSelectedSystem || { key: 'S4PCLNT100', system: 'S4P' },
              filteredRoles
            );
          }
        },
        {
          label: '➖ Remover funções (CUA_REMOVE)',
          val: 'Remover funções (CUA_REMOVE)',
          action: () => {
            appendAuthorizationMessage('user', 'Remover funções (CUA_REMOVE)');
            confirmAuthorizationRemoval(
              authorizationTargetUser,
              authorizationSelectedSystem || { key: 'S4PCLNT100', system: 'S4P' },
              filteredRoles
            );
          }
        },
        {
          label: '🔄 Nova análise',
          val: 'Nova análise',
          action: () => {
            appendAuthorizationMessage('user', 'Nova análise');
            resetAuthorizationChat();
          }
        }
      ].sort((a, b) => a.val.localeCompare(b.val, 'pt', { sensitivity: 'base' }));

      actions.forEach(item => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'auth-chat-system-card';
        btn.style.flex = '0 0 auto';
        btn.style.padding = '8px 14px';
        btn.onclick = () => {
          if (btn.parentElement) {
            btn.parentElement.querySelectorAll('button').forEach(b => {
              b.classList.remove('selected');
              b.setAttribute('aria-pressed', 'false');
            });
          }
          btn.classList.add('selected');
          btn.setAttribute('aria-pressed', 'true');
          item.action();
        };
        btn.innerHTML = `<span class="sys-code" style="font-size:0.84rem; font-weight:700;">${escapeAuthorizationText(item.label)}</span>`;
        grid.appendChild(btn);
      });

      container.appendChild(grid);
      container.scrollTop = container.scrollHeight;
    }

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

    async function processIndividualUserSubmit(rawVal) {
      let userVal = rawVal.trim().toUpperCase();
      if (/^\d+$/.test(userVal)) {
        userVal = 'S' + userVal.replace(/^0+/, '');
      }

      showAuthorizationTypingIndicator(null, `A pesquisar o nome do utilizador ${escapeAuthorizationText(userVal)} via RFC em PRD...`);

      let userFullName = '';
      try {
        const resp = await fetch('/api/authorizations/hr-search', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query: userVal, target_system_key: 'S4PCLNT100', max_results: 5 })
        });
        if (resp.ok) {
          const hrRes = await resp.json();
          if (hrRes && hrRes.success && Array.isArray(hrRes.data) && hrRes.data.length > 0) {
            const matched = hrRes.data.find(d => {
              const u = String(d.user_id || d.sap_user || '').trim().toUpperCase();
              const p = String(d.pernr || '').trim();
              return u === userVal || ('S' + p.replace(/^0+/, '')) === userVal;
            }) || hrRes.data[0];

            userFullName = (matched.full_name || `${matched.first_name || ''} ${matched.last_name || ''}`).trim();
          }
        }
      } catch (err) {
        console.warn('[HR SEARCH RFC] Não foi possível obter o nome do utilizador:', err);
      }

      const userDisplay = userFullName ? `${userVal} - ${userFullName}` : userVal;

      if (authorizationIndividualContext) {
        authorizationIndividualContext.targetUser = userVal;
        authorizationIndividualContext.userDisplayName = userDisplay;
        authorizationIndividualContext.userFullName = userFullName;
      }
      authorizationTargetUser = userVal;

      authorizationChatState = AUTH_CHAT_STATES.WAITING_INDIVIDUAL_SYSTEM;
      hideAuthorizationTypingIndicator();

      const procName = authorizationIndividualContext?.processName || 'Processo';
      const msgHtml = `Registado o utilizador <b>${escapeAuthorizationText(userDisplay)}</b> para <b>${escapeAuthorizationText(procName)}</b>.<br><br>Em que sistema/ambiente pretende efetuar a alteração?`;

      appendAuthorizationMessage(
        'assistant',
        msgHtml,
        true
      );
      showIndividualSystemOptions();
    }

    async function handleAuthorizationChatSubmit(event) {
      if (event) event.preventDefault();
      
      const input = document.getElementById('authorization-chat-input');
      if (!input) return;

      const rawVal = input.value.trim();
      input.value = '';
      updateAuthorizationComposer();

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

      if (authorizationUatCreateDocumentFlow?.active) {
        if (handleUatCreateDocumentChatSubmit(rawVal)) {
          return;
        }
      }

      if (!rawVal) return;

      // Exibir a mensagem do utilizador exatamente como foi escrita
      appendAuthorizationMessage('user', rawVal);

      // Se estivermos na escolha inicial (Ticket vs Processo):
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INITIAL_CHOICE) {
        if (normVal.includes('TICKET') || normVal === 'TICKET') {
          handleInitialChoiceSelect('Ticket', true);
        } else {
          handleInitialChoiceSelect('Processo', true);
        }
        return;
      }

      // Se estivermos a aguardar a seleção de equipa Jira:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_TEAM) {
        if (normVal.includes('PROCESSO')) {
          renderRoutineSuggestionsInitial();
        } else {
          showJiraAssigneeOptions(rawVal.trim(), true);
        }
        return;
      }

      // Se estivermos a aguardar a escolha do modo de filtro (Todos os tickets vs Processo):
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_FILTER_MODE) {
        handleFilterModeSelect(rawVal.trim(), true);
        return;
      }

      // Se estivermos a aguardar a seleção de processo Jira:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_PROCESS) {
        if (normVal.includes('RESPONSAVEL') || normVal.includes('RESPONSÁVEL')) {
          showJiraAssigneeOptions(authorizationSelectedJiraTeam, true);
        } else if (normVal.includes('EQUIPA')) {
          showJiraTeamOptions();
        } else {
          selectJiraProcess(rawVal.trim(), true);
        }
        return;
      }

      // Se estivermos a aguardar a seleção de ticket Jira:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_JIRA_TICKET) {
        if (normVal.includes('EQUIPA') || normVal.includes('TROCAR EQUIPA')) {
          showJiraTeamOptions();
          return;
        }
        if (normVal.includes('PROCESSO')) {
          renderRoutineSuggestionsInitial();
          return;
        }
        let matched = null;
        if (Array.isArray(authorizationCachedJiraTickets)) {
          const cleanRaw = rawVal.trim().toUpperCase();
          matched = authorizationCachedJiraTickets.find(t =>
            t.key.toUpperCase() === cleanRaw ||
            cleanRaw.includes(t.key.toUpperCase()) ||
            (t.summary && t.summary.toUpperCase().includes(cleanRaw))
          );
        }
        if (matched) {
          selectJiraTicket(matched, true);
        } else {
          selectJiraTicket({ key: rawVal.trim().toUpperCase(), summary: rawVal.trim() }, true);
        }
        return;
      }

      // Se estivermos a aguardar a ação do ticket Jira:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_TICKET_ACTION) {
        if (normVal.includes('ANALISAR') || normVal.includes('AUTORIZACAO') || normVal.includes('AUTORIZAÇÃO')) {
          handleTicketActionSelect('Analisar autorizações', authorizationSelectedJiraTicket);
        } else if (normVal.includes('DADOS DE UTILIZADOR') || normVal.includes('UTILIZADOR')) {
          handleTicketActionSelect('Dados de utilizador', authorizationSelectedJiraTicket);
        } else if (normVal.includes('PERFIL')) {
          handleTicketActionSelect('Perfil de autorização', authorizationSelectedJiraTicket);
        } else {
          handleTicketActionSelect('Processo', authorizationSelectedJiraTicket);
        }
        return;
      }

      // Se o utilizador fizer uma pergunta sobre a tabela de roles exibida no ecra (expiradas, ativas, contagem):
      if (handleContextualRolesQuery(rawVal, normVal)) {
        return;
      }

      // Se estivermos no fluxo de alteração individual à espera do utilizador de referência (Por cópia):
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_COPY_REFERENCE_USER) {
        performHrReferenceUserSearch(rawVal.trim());
        return;
      }

      // Se estivermos no fluxo de pesquisa no RH Produtivo:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_HR_SEARCH_QUERY) {
        performHrUserSearch(rawVal.trim());
        return;
      }

      // Se estivermos a recolher o campo FUNCTION no chat:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_FUNCTION) {
        const funcVal = rawVal.trim();
        if (!authorizationIndividualContext) authorizationIndividualContext = {};
        if (!authorizationIndividualContext.parameters) authorizationIndividualContext.parameters = {};
        authorizationIndividualContext.parameters.FUNCTION = funcVal;

        authorizationChatState = AUTH_CHAT_STATES.WAITING_CUA_DEPARTMENT;
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `Registada a Função: **${escapeAuthorizationText(funcVal)}**.\n\nAgora, indique o **Departamento (DEPARTMENT)** do utilizador no CUA:`
        );
        updateAuthorizationComposer();
        return;
      }

      // Se estivermos a recolher o campo DEPARTMENT no chat:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_DEPARTMENT) {
        const deptVal = rawVal.trim();
        if (!authorizationIndividualContext) authorizationIndividualContext = {};
        if (!authorizationIndividualContext.parameters) authorizationIndividualContext.parameters = {};
        authorizationIndividualContext.parameters.DEPARTMENT = deptVal;

        authorizationChatState = AUTH_CHAT_STATES.WAITING_CUA_MOB_NUMBER;
        hideAuthorizationTypingIndicator();
        appendAuthorizationMessage(
          'assistant',
          `Registado o Departamento: **${escapeAuthorizationText(deptVal)}**.\n\nPor fim, indique o **Telefone / Telemóvel (MOB_NUMBER)** do utilizador no CUA:`
        );
        updateAuthorizationComposer();
        return;
      }

      // Se estivermos a recolher o campo MOB_NUMBER no chat:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_CUA_MOB_NUMBER) {
        const mobVal = rawVal.trim();
        if (!authorizationIndividualContext) authorizationIndividualContext = {};
        if (!authorizationIndividualContext.parameters) authorizationIndividualContext.parameters = {};
        authorizationIndividualContext.parameters.MOB_NUMBER = mobVal;

        const firstNameVal = authorizationIndividualContext.parameters.NAME_FIRST || authorizationIndividualContext.hrData?.first_name || '';
        const lastNameVal = authorizationIndividualContext.parameters.NAME_LAST || authorizationIndividualContext.hrData?.last_name || '';
        const emailVal = authorizationIndividualContext.parameters.SMTP_ADDR || authorizationIndividualContext.hrData?.email || '';
        const funcVal = authorizationIndividualContext.parameters.FUNCTION || 'N/D';
        const deptVal = authorizationIndividualContext.parameters.DEPARTMENT || 'N/D';

        saveCuaUserDetails(firstNameVal, lastNameVal, emailVal, funcVal, deptVal, mobVal);
        return;
      }

      // Se estivermos no fluxo de alteração individual à espera do utilizador alvo:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_USER) {
        await processIndividualUserSubmit(rawVal);
        return;
      }

      // Se estivermos no fluxo de alteração individual à espera dos parâmetros:
      if (authorizationChatState === AUTH_CHAT_STATES.WAITING_INDIVIDUAL_PARAMS) {
        const detailsVal = rawVal.trim();
        const ctx = authorizationIndividualContext || {};
        const procName = ctx.processName || 'Processo';
        const user = ctx.userDisplayName || ctx.targetUser || authorizationTargetUser || 'N/D';
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
        showAuthorizationProfileSubroutineOptions();
        return;
      }

      // Se for seleção da opção 'Dados de utilizador' ou 'Dados mestres'
      if (
        normVal.includes('DADOS DE UTILIZADOR') ||
        normVal.includes('DADOS DO UTILIZADOR') ||
        normVal.includes('DADOS UTILIZADOR') ||
        normVal.includes('DADOS DE USER') ||
        normVal.includes('DADOS MESTRES') ||
        normVal.includes('DADOS MESTRE')
      ) {
        showUserDataSubroutineOptions();
        return;
      }

      // Se for seleção de uma sub-rotina específica de Dados de Utilizador
      if (normVal.includes('CRIAR UTILIZADOR') || normVal.includes('ADICIONAR UTILIZADOR') || normVal.includes('CRIAR USER')) {
        selectUserDataSubroutine({ label: 'âž• Criar utilizador', val: 'Criar utilizador', scriptName: 'L. CUA_CRIAR_USER.py', category: 'CUA_CRIAR_USER' });
        return;
      }

      if (normVal.includes('ALTERAR SENHA') || normVal.includes('RESET PASSWORD') || normVal.includes('MUDAR SENHA') || normVal.includes('ALTERAR PASSWORD')) {
        selectUserDataSubroutine({ label: '🔑 Alterar Senha', val: 'Alterar Senha', scriptName: 'su01_reset_password.py', category: 'CUA Login' });
        return;
      }

      if (normVal.includes('DELIMITAR DATA FIM') || normVal.includes('DELIMITAR DATA') || normVal.includes('DATA FIM') || normVal.includes('ENDDATE')) {
        selectUserDataSubroutine({ label: '📅 Delimitar data fim', val: 'Delimitar data fim', scriptName: 'I. CUA_ENDDATE.py', category: 'CUA_ENDDATE' });
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

      if (normVal.includes('UAT SIMULACAO')) {
        showUatSimulationSubroutineOptions();
        return;
      }

    async function processAnalysisUserSubmit(userRaw, extractedSystem, extractedType) {
      let userCode = userRaw.trim().toUpperCase();
      if (/^\d+$/.test(userCode)) {
        userCode = 'S' + userCode.replace(/^0+/, '');
      }

      showAuthorizationTypingIndicator(null, `A pesquisar o nome do utilizador ${escapeAuthorizationText(userCode)} via RFC em PRD...`);

      let userFullName = '';
      try {
        const resp = await fetch('/api/authorizations/hr-search', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ query: userCode, target_system_key: 'S4PCLNT100', max_results: 5 })
        });
        if (resp.ok) {
          const hrRes = await resp.json();
          if (hrRes && hrRes.success && Array.isArray(hrRes.data) && hrRes.data.length > 0) {
            const matched = hrRes.data.find(d => {
              const u = String(d.user_id || d.sap_user || '').trim().toUpperCase();
              const p = String(d.pernr || '').trim();
              return u === userCode || ('S' + p.replace(/^0+/, '')) === userCode;
            }) || hrRes.data[0];

            userFullName = (matched.full_name || `${matched.first_name || ''} ${matched.last_name || ''}`).trim();
          }
        }
      } catch (err) {
        console.warn('[HR SEARCH RFC] Não foi possível obter o nome do utilizador:', err);
      }

      const userDisplay = userFullName ? `${userCode} - ${userFullName}` : userCode;
      authorizationTargetUser = userCode;
      authorizationTargetUserDisplayName = userDisplay;

      if (extractedSystem) authorizationSelectedSystem = extractedSystem;
      if (extractedType) authorizationSelectedAnalysisType = extractedType;

      if (!authorizationSelectedAnalysisType && authorizationAnalysisTypes.length > 0) {
        authorizationSelectedAnalysisType = authorizationAnalysisTypes.find(t => t.id === 'authorizations') || authorizationAnalysisTypes[0];
      }

      hideAuthorizationTypingIndicator();

      if (authorizationSelectedSystem) {
        authorizationChatState = AUTH_CHAT_STATES.LOADING;
        showAuthorizationTypingIndicator();
        setTimeout(() => {
          appendAuthorizationMessage('assistant', `Perfeito. Registado o utilizador <b>${escapeAuthorizationText(userDisplay)}</b>.`, true);
          showAuthorizationSummary();
        }, 400);
      } else {
        authorizationChatState = AUTH_CHAT_STATES.WAITING_SYSTEM;
        appendAuthorizationMessage('assistant', `Perfeito. Registado o utilizador <b>${escapeAuthorizationText(userDisplay)}</b>.<br><br>Em que sistema pretende fazer a análise?`, true);
        showAuthorizationSystemOptions();
      }
    }

      // Se a mensagem contiver a indicação de um utilizador SAP para analisar (ex: "Quero analisar o user S5441")
      const { extractedUser, extractedSystem, extractedType } = extractAuthorizationParamsFromText(rawVal);

      if (extractedUser) {
        await processAnalysisUserSubmit(extractedUser, extractedSystem, extractedType);
        return;
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
            appendAuthorizationMessage('assistant', `Perfeito. Registado o utilizador <b>${escapeAuthorizationText(authorizationTargetUserDisplayName || authorizationTargetUser)}</b>.`, true);
            showAuthorizationSummary();
          }, 400);
          return;
        }

        if (!authorizationTargetUser) {
          await processAnalysisUserSubmit(rawVal.trim(), extractedSystem, extractedType);
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
        card.setAttribute('data-system', sys.system || sys.key);

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
        if (authorizationIndividualContext && authorizationIndividualContext.processName) {
          const ctx = authorizationIndividualContext;
          ctx.selectedSystem = sysInfo;
          if (!ctx.targetUser && !authorizationTargetUser) {
            authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
            hideAuthorizationTypingIndicator();
            appendAuthorizationMessage(
              'assistant',
              `Qual é o utilizador SAP que pretende processar na sub-rotina **${escapeAuthorizationText(ctx.processName)}** no ambiente **${escapeAuthorizationText(sysInfo.system || sysInfo.key)}** (ex: CSILVA)?`
            );
            updateAuthorizationComposer();
            const inputEl = document.getElementById('authorization-chat-input');
            if (inputEl) {
              inputEl.placeholder = 'Introduza o utilizador SAP (ex: CSILVA)...';
              inputEl.focus();
            }
          } else {
            promptIndividualProcessParameters();
          }
          return;
        }

        if (!authorizationTargetUser) {
          authorizationChatState = AUTH_CHAT_STATES.WAITING_USER;
          hideAuthorizationTypingIndicator();
          appendAuthorizationMessage('assistant', 'Qual o utilizador SAP que pretende analisar (ex: CSILVA)?');
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

      // Transição para a próxima etapa: perguntar o sistema alvo
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
          <div style="margin-bottom: 10px;">Para abrir o SAP CUA, ligue primeiro o Worker através do botão "Ligar Worker" no canto inferior esquerdo.</div>
          <div style="margin-bottom: 12px;">Quando o Worker estiver online, volte a clicar em "Iniciar análise".</div>
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
      const userDisplay = authorizationTargetUserDisplayName || authorizationTargetUser;
      const analysisTypeLabel = authorizationSelectedAnalysisType ? authorizationSelectedAnalysisType.label : 'Autorizações';
      const modeText = isRfcFlow ? 'RFC Direta (pyrfc)' : 'SAP GUI';

      const statusMsgHtml = `
        <div style="display:flex; flex-direction:column; gap:4px;">
          <div style="font-weight:700; color:var(--text-primary);">A iniciar a análise no sistema <b>${escapeAuthorizationText(selectedSystemLabel)}</b>...</div>
          <div style="font-size:0.83rem; color:var(--text-secondary); line-height:1.5; margin-top:2px;">
            <b>• Utilizador:</b> ${escapeAuthorizationText(userDisplay)}<br>
            <b>• Tipo de análise:</b> ${escapeAuthorizationText(analysisTypeLabel)}<br>
            <b>• Modo de execução:</b> ${escapeAuthorizationText(modeText)}
          </div>
        </div>
      `;

      appendAuthorizationMessage('assistant', statusMsgHtml, true);
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

                window.setTimeout(() => {
                  if (typeof renderPostAnalysisFollowUpQuestion === 'function') {
                    renderPostAnalysisFollowUpQuestion();
                  }
                }, 400);
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

            // Verificar status do worker apenas para atualizar a barra lateral
            try {
              const statusRes = await fetch('/api/worker/status', { cache: 'no-store' });
              if (statusRes.ok) {
                const statusData = await statusRes.json();
                updateSidebarWorkerStatus(statusData);
                if (job.state === 'pending' && statusData.online === false && statusData.last_seen_seconds > 20) {
                  throw { isWorkerOfflineDuringExecution: true };
                }
              }
            } catch (statusErr) {
              if (statusErr && statusErr.isWorkerOfflineDuringExecution) {
                throw statusErr;
              }
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
      hideAuthorizationTypingIndicator();
      authorizationChatState = AUTH_CHAT_STATES.READY;
      updateAuthorizationComposer();

      // Executar a análise diretamente incorporando os detalhes no balão de resposta
      setTimeout(() => {
        startAuthorizationAnalysis();
      }, 200);
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

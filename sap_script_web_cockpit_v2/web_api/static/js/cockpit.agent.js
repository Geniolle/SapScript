/*
 * cockpit.agent.js - escapeHtml e utilitarios, Agente Salsa IT (chat/PFCG/FI),
 * sidebar de tickets Jira, tooltips e o bootstrap (loadJobs -> switchView).
 * Depende de cockpit.core.js (loadJobs, renderQueue, consts do DOM, ...).
 */
    // Close menus when clicking outside
    document.addEventListener('click', function() {
        const allMenus = document.querySelectorAll('.job-options-menu');
        allMenus.forEach(menu => {
            menu.style.display = 'none';
        });
    });

    function escapeHtml(value) {
      return String(value)
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
    }

    let asiChatHistory = [];
    let asiChatInitialized = false;
    let asiChatMockTimer = null;
    let asiChatMessageCounter = 0;
    let asiPfcgPollingTimer = null;
    let asiPfcgPollingInFlight = false;
    let asiSelectedActions = { '__root__': null };
    const ASI_DEFAULT_PLACEHOLDER = 'Escreva a sua mensagem...';
    const ASI_PFCG_PLACEHOLDER = 'Ex.: Z_AUTHORITY_LISTADEPRECOS';
    const ASI_PFCG_AWAITING_INPUT = 'pfcg_role_analysis_name';
    const ASI_PFCG_TCODE_INPUT = 'pfcg_transaction_code';
    const ASI_PFCG_TCODE_PATTERN = /^[A-Z0-9_/$+.\-]{1,40}$/;
    const ASI_PFCG_AUTHOBJ_INPUT = 'pfcg_auth_object';
    const ASI_PFCG_AUTHOBJ_PATTERN = /^[A-Z0-9_/]{1,40}$/;
    const ASI_PFCG_POLL_INTERVAL_MS = 1000;
    const ASI_PFCG_POLL_TIMEOUT_MS = 60000;
    const ASI_PFCG_INVALID_MESSAGE = 'O nome do Perfil de Autorização contém caracteres inválidos.\nUtilize apenas letras, números, "_", "-", "/" ou ":".';
    const ASI_PFCG_ROLE_PATTERN = /^[A-Z0-9_/:-]+$/;

    // Estado do último resultado EXISTE de análise de função PFCG.
    // Mantido fora de asiConversationState para sobreviver a resets de navegação (ex.: "Voltar").
    let asiPfcgRoleState = null;
    let asiLastAnalyzedPfcgRole = '';
    let asiPfcgSubPollingTimer = null;
    let asiPfcgSubPollingInFlight = false;
    const ASI_PFCG_SUB_POLL_INTERVAL_MS = 1000;
    const ASI_PFCG_SUB_POLL_TIMEOUT_MS = 60000;
    let asiPfcgIndividualPollingTimer = null;
    let asiPfcgIndividualPollingInFlight = false;
    let asiPfcgTransportSearchPollingTimer = null;
    let asiPfcgTransportSearchPollingInFlight = false;

    const ASI_PFCG_ROLE_BACK_ACTION = { id: 'pfcg-role-back', label: '← Voltar', icon: 'analysis' };
    const ASI_PFCG_ROLE_ANALYZE_TRANSACTIONS_ACTION = { id: 'pfcg-role-analyze-transactions', label: 'Analisar por Transação', icon: 'analysis' };
    const ASI_PFCG_ROLE_ANALYZE_USERS_ACTION = { id: 'pfcg-role-analyze-users', label: 'Analisar por Utilizador', icon: 'user-plus' };
    const ASI_PFCG_ROLE_BACK_TO_CARD_ACTION = { id: 'pfcg-role-back-to-card', label: '← Voltar para função', icon: 'analysis' };
    const ASI_PFCG_ROLE_RESULT_ACTIONS = [
        ASI_PFCG_ROLE_BACK_ACTION,
        ASI_PFCG_ROLE_ANALYZE_TRANSACTIONS_ACTION,
        ASI_PFCG_ROLE_ANALYZE_USERS_ACTION
    ];
    // Opções de "O que deseja fazer com o Perfil de Autorização?" (filhas de
    // perfil-autorizacao no salsaAgentActions), para re-apresentar após um resultado.
    function asiPfcgRootMenuActions() {
        const node = asiFindQuickAction('perfil-autorizacao', salsaAgentActions);
        return node && Array.isArray(node.children) ? node.children : [];
    }
    const ASI_PFCG_ROOT_MENU_META = {
        actionLevel: 2,
        parentActionId: 'perfil-autorizacao',
        selectionGroupKey: 'perfil-autorizacao'
    };
    const ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT = 'pfcg_individual_role_name';
    const ASI_PFCG_COMPOSTA_ROLE_NAME_INPUT = 'pfcg_composta_role_name';
    const ASI_PFCG_COMPOSTA_DESCRIPTION_INPUT = 'pfcg_composta_description';
    const ASI_PFCG_COMPOSTA_CHILDREN_INPUT = 'pfcg_composta_children';
    const ASI_PFCG_INDIVIDUAL_ROLE_NAME_PLACEHOLDER = 'Ex.: Z_FI_CLIENTES';
    const ASI_PFCG_INDIVIDUAL_DESCRIPTION_INPUT = 'pfcg_individual_description';
    const ASI_PFCG_INDIVIDUAL_TCODES_INPUT = 'pfcg_individual_tcodes';
    const ASI_PFCG_INDIVIDUAL_DESCRIPTION_PLACEHOLDER = 'Ex.: Acesso a relatórios de vendas';
    const ASI_PFCG_INDIVIDUAL_TCODES_PLACEHOLDER = 'Ex.: FB01, VL03N';
    const ASI_PFCG_INDIVIDUAL_BACK_ACTION = { id: 'pfcg-create-individual-back', label: '← Voltar', icon: 'analysis' };
    const ASI_PFCG_INDIVIDUAL_CONFIRM_ACTION = { id: 'pfcg-create-individual-confirm', label: 'Confirmar criação', icon: 'shield-plus' };
    const ASI_PFCG_TRANSPORT_CREATE_DESCRIPTION_INPUT = 'pfcg_transport_create_description';
    const ASI_PFCG_TRANSPORT_LOCAL_ACTION = { id: 'pfcg-transport-local', label: 'Sem transporte (Local)', icon: 'analysis' };
    const ASI_PFCG_TRANSPORT_CREATE_ACTION = { id: 'pfcg-transport-create', label: 'Criar nova Request', icon: 'shield-plus' };
    const ASI_PFCG_TRANSPORT_EXISTING_ACTION = { id: 'pfcg-transport-existing', label: 'Usar Request existente', icon: 'analysis' };
    const ASI_PFCG_TRANSPORT_BACK_ACTION = { id: 'pfcg-transport-back', label: '← Voltar', icon: 'analysis' };
    const ASI_PFCG_DELETE_ROLE_NAME_INPUT = 'pfcg_delete_role_name';
    const ASI_PFCG_DELETE_ROLE_NAME_PLACEHOLDER = 'Ex.: Z_FI_CLIENTES';
    const ASI_PFCG_DELETE_TRANSPORT_CREATE_DESCRIPTION_INPUT = 'pfcg_delete_transport_create_description';
    const ASI_PFCG_DELETE_BACK_ACTION = { id: 'pfcg-delete-individual-back', label: '← Voltar', icon: 'analysis' };
    const ASI_PFCG_DELETE_CONFIRM_ACTION = { id: 'pfcg-delete-individual-confirm', label: 'Confirmar eliminação', icon: 'trash' };
    const ASI_PFCG_DELETE_TRANSPORT_LOCAL_ACTION = { id: 'pfcg-delete-transport-local', label: 'Sem transporte (Local)', icon: 'analysis' };
    const ASI_PFCG_DELETE_TRANSPORT_CREATE_ACTION = { id: 'pfcg-delete-transport-create', label: 'Criar nova Request', icon: 'shield-plus' };
    const ASI_PFCG_DELETE_TRANSPORT_EXISTING_ACTION = { id: 'pfcg-delete-transport-existing', label: 'Usar Request existente', icon: 'analysis' };
    const ASI_PFCG_DYNAMIC_ACTION_IDS = new Set([
        'pfcg-role-back',
        'pfcg-role-back-to-card',
        'pfcg-role-analyze-transactions',
        'pfcg-role-analyze-users',
        'pfcg-create-individual-back',
        'pfcg-create-individual-confirm',
        'pfcg-transport-local',
        'pfcg-transport-create',
        'pfcg-transport-existing',
        'pfcg-transport-back',
        'pfcg-delete-individual-back',
        'pfcg-delete-individual-confirm',
        'pfcg-delete-transport-local',
        'pfcg-delete-transport-create',
        'pfcg-delete-transport-existing'
    ]);

    function asiDefaultConversationState() {
        return {
            processo: '',
            subprocesso: '',
            actionId: '',
            mode: '',
            selectedFiEnvironment: '',
            selectedFiBranch: '',
            selectedFiWorkflow: '',
            lastFiDocumentNumber: '',
            lastFiDocumentEnvironment: '',
            lastFiDocumentBranch: '',
            lastFiDocumentWorkflow: '',
            lastF110ProposalPayload: null,
            lastF110ProposalResult: null,
            awaitingInput: '',
            pendingJobId: '',
            pendingRoleName: '',
            pendingMessageId: '',
            lastPfcgRoleName: '',
            pendingExcelSelectionJobId: '',
            pendingExcelAnalyzeJobId: '',
            pendingExcelMessageId: '',
            pendingExcelFileName: '',
            pfcgCreateRoleName: '',
            pfcgCreateDescription: '',
            pfcgCreateTcodes: [],
            pfcgCreateTransportMode: '',
            pfcgCreateTransportRequestNumber: '',
            pfcgCreateTransportRequestDescription: '',
            pfcgCreatePreviewJobId: '',
            pfcgCreateMessageId: '',
            pfcgDeleteRoleName: '',
            pfcgDeleteTransportMode: '',
            pfcgDeleteTransportRequestNumber: '',
            pfcgDeleteTransportRequestDescription: '',
            pfcgDeletePreviewJobId: '',
            isBusy: false
        };
    }

    function asiNormalizeFiBranch(branch) {
        const value = String(branch || '').trim().toLowerCase();
        if (value === 'fornecedor' || value === 'vendor') return 'fornecedor';
        if (value === 'razao' || value === 'razão' || value === 'gl') return 'razao';
        return 'cliente';
    }

    // asiGetFiActionContext: definicao unica mais abaixo (antes de asiBuildFiDefaultPayload).
    // A copia identica que existia aqui foi removida na Fase 1 da refatoracao.

    function asiBuildF110DefaultInfoAction(environment, branch) {
        const env = String(environment || 'QAD').trim().toUpperCase();
        const normalizedBranch = asiNormalizeFiBranch(branch);
        const branchLabel = normalizedBranch === 'fornecedor'
            ? 'Fornecedor'
            : normalizedBranch === 'razao'
                ? 'Razão'
                : 'Cliente';

        return {
            id: `testes-unitarios-executar-f110-${env.toLowerCase()}-default-${normalizedBranch}-info-default`,
            label: 'Default',
            icon: 'analysis',
            processo: 'Testes Unitários',
            environment: env,
            branch: normalizedBranch,
            mode: 'default',
            workflow: 'f110_default_document',
            prompt: `Usar informações Default para ${branchLabel}.`,
            children: []
        };
    }

    function asiBuildF110DefaultAccountAction(environment, branch) {
        const env = String(environment || 'QAD').trim().toUpperCase();
        const normalizedBranch = asiNormalizeFiBranch(branch);
        const branchLabel = normalizedBranch === 'fornecedor'
            ? 'Fornecedor'
            : normalizedBranch === 'razao'
                ? 'Razão'
                : 'Cliente';

        return {
            id: `testes-unitarios-executar-f110-${env.toLowerCase()}-default-${normalizedBranch}`,
            label: branchLabel,
            icon: 'analysis',
            processo: 'Testes Unitários',
            environment: env,
            branch: normalizedBranch,
            mode: 'default',
            prompt: `Quero validar um documento de ${branchLabel} no F110 em ${env}.`,
            workflow: 'f110_default_document',
            children: []
        };
    }

    function asiBuildF110DefaultChildren(environment) {
        const env = String(environment || 'QAD').trim().toUpperCase();
        return [
            {
                id: `testes-unitarios-executar-f110-${env.toLowerCase()}-default`,
                label: 'Execução Default',
                icon: 'analysis',
                processo: 'Testes Unitários',
                environment: env,
                mode: 'default',
                prompt: `Quero executar o F110 em ${env} com a execução Default.`,
                followupText: 'Escolha o tipo de conta para validar:',
                followupActionsSource: 'children',
                children: [
                    asiBuildF110DefaultAccountAction(env, 'cliente'),
                    asiBuildF110DefaultAccountAction(env, 'fornecedor'),
                    asiBuildF110DefaultAccountAction(env, 'razao')
                ]
            },
            {
                id: `testes-unitarios-executar-f110-${env.toLowerCase()}-manual`,
                label: 'Execução Manual',
                icon: 'analysis',
                processo: 'Testes Unitários',
                environment: env,
                mode: 'manual',
                prompt: `Quero executar o F110 em ${env} com a execução Manual.`,
                children: []
            }
        ];
    }

    let asiConversationState = asiDefaultConversationState();
    // Declarado antes de salsaAgentActions porque este é espalhado (...) dentro do array.
    const ASI_MAIN_MENU_ACTION = {
        id: '__asi-main-menu__',
        label: 'Menu Inicial',
        icon: 'analysis'
    };
    const salsaAgentActions = [
        {
            id: 'configuracoes',
            label: 'Configurações',
            icon: 'settings',
            prompt: 'Quero ajuda com configurações SAP.',
            followupText: 'Escolha uma opção de Configurações:',
            followupActionsSource: 'children',
            children: [
                {
                    id: 'perfil-autorizacao',
                    label: 'Perfil de Autorização',
                    icon: 'authorization',
                    prompt: 'Quero trabalhar com Perfil de Autorização.',
                    followupText: 'O que deseja fazer com o Perfil de Autorização?',
                    followupActionsSource: 'children',
                    children: [
                        {
                            id: 'pfcg-role-analyze',
                            label: 'Analisar',
                            icon: 'analysis',
                            processo: 'Funções PFCG',
                            subprocesso: 'A. PFCG_CREATE.py',
                            prompt: 'Quero analisar um Perfil de Autorização.',
                            followupText: 'O que deseja analisar?',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'pfcg-role-analyze-funcao',
                                    label: 'Função',
                                    icon: 'analysis',
                                    mode: 'analyze',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'A. PFCG_CREATE.py',
                                    prompt: 'Quero analisar a função.',
                                    followupText: 'Qual é o nome do Perfil de Autorização que deseja analisar em PRD?',
                                    children: []
                                },
                                {
                                    id: 'pfcg-role-analyze-transacao',
                                    label: 'Transação',
                                    icon: 'analysis',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'A. PFCG_CREATE.py',
                                    prompt: 'Quero analisar por transação.',
                                    children: []
                                },
                                {
                                    id: 'pfcg-role-analyze-objeto',
                                    label: 'Objeto de autorização',
                                    icon: 'authorization',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'A. PFCG_CREATE.py',
                                    prompt: 'Quero analisar por objeto de autorização.',
                                    children: []
                                }
                            ]
                        },
                        {
                            id: 'pfcg-create',
                            label: 'Criar funções',
                            icon: 'shield-plus',
                            processo: 'Funções PFCG',
                            subprocesso: 'A. PFCG_CREATE.py',
                            prompt: 'Quero criar um Perfil de Autorização.',
                            followupText: 'Como deseja prosseguir?',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'pfcg-create-execute',
                                    label: 'Preparar criação',
                                    icon: 'shield-plus',
                                    mode: 'prepare',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'A. PFCG_CREATE.py',
                                    prompt: 'Quero preparar a criação do Perfil de Autorização.',
                                    followupText: 'Para preparar a criação, selecione o ficheiro Excel com a configuração do Perfil de Autorização.',
                                    followupActionsSource: 'children',
                                    children: [
                                        {
                                            id: 'pfcg-create-select-excel',
                                            label: 'Selecionar Excel',
                                            icon: 'upload',
                                            mode: 'select_excel',
                                            processo: 'Funções PFCG',
                                            subprocesso: 'A. PFCG_CREATE.py',
                                            prompt: 'Selecionar Excel',
                                            followupText: '',
                                            children: []
                                        },
                                        {
                                            id: 'pfcg-create-individual',
                                            label: 'Criar Individualmente',
                                            icon: 'shield-plus',
                                            mode: 'create_individual',
                                            processo: 'Funções PFCG',
                                            subprocesso: 'A. PFCG_CREATE.py',
                                            prompt: 'Quero criar o Perfil de Autorização individualmente via RFC.',
                                            followupText: '',
                                            children: []
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            ...ASI_MAIN_MENU_ACTION,
                            prompt: 'Quero voltar ao menu principal.'
                        },
                        {
                            id: 'pfcg-composta',
                            label: 'Função Composta',
                            icon: 'layers',
                            processo: 'Funções PFCG',
                            subprocesso: 'D. PFCG_COMPOSTA.py',
                            prompt: 'Quero trabalhar com uma Função Composta.',
                            followupText: 'Como deseja prosseguir?',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'pfcg-composta-analyze',
                                    label: 'Analisar',
                                    icon: 'analysis',
                                    mode: 'analyze',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'D. PFCG_COMPOSTA.py',
                                    prompt: 'Quero analisar a Função Composta.',
                                    followupText: 'Qual é o nome da Função Composta que deseja analisar em PRD?',
                                    children: []
                                },
                                {
                                    id: 'pfcg-composta-execute',
                                    label: 'Preparar criação',
                                    icon: 'shield-plus',
                                    mode: 'prepare',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'D. PFCG_COMPOSTA.py',
                                    prompt: 'Quero preparar a criação da Função Composta.',
                                    followupText: 'Para preparar a criação, selecione o ficheiro Excel com a configuração da Função Composta.',
                                    followupActionsSource: 'children',
                                    children: [
                                        {
                                            id: 'pfcg-composta-select-excel',
                                            label: 'Selecionar Excel',
                                            icon: 'upload',
                                            mode: 'select_excel',
                                            processo: 'Funções PFCG',
                                            subprocesso: 'D. PFCG_COMPOSTA.py',
                                            prompt: 'Selecionar Excel',
                                            followupText: '',
                                            children: []
                                        },
                                        {
                                            id: 'pfcg-composta-individual',
                                            label: 'Criar Individualmente',
                                            icon: 'shield-plus',
                                            mode: 'create_individual',
                                            processo: 'Funções PFCG',
                                            subprocesso: 'D. PFCG_COMPOSTA.py',
                                            prompt: 'Quero criar a Função Composta individualmente via RFC.',
                                            followupText: '',
                                            children: []
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            id: 'pfcg-delete',
                            label: 'Eliminar Perfil',
                            icon: 'trash',
                            processo: 'Funções PFCG',
                            subprocesso: 'B. PFCG_DELETE.py',
                            prompt: 'Quero eliminar um Perfil de Autorização.',
                            followupText: 'Como deseja prosseguir com a eliminação?',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'pfcg-delete-select-excel',
                                    label: 'Selecionar Excel',
                                    icon: 'upload',
                                    mode: 'select_excel_delete',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'B. PFCG_DELETE.py',
                                    prompt: 'Selecionar Excel para eliminação em massa.',
                                    followupText: '',
                                    children: []
                                },
                                {
                                    id: 'pfcg-delete-individual',
                                    label: 'Eliminar Individualmente',
                                    icon: 'trash',
                                    mode: 'delete_individual',
                                    processo: 'Funções PFCG',
                                    subprocesso: 'B. PFCG_DELETE.py',
                                    prompt: 'Quero eliminar o Perfil de Autorização individualmente via RFC.',
                                    followupText: '',
                                    children: []
                                }
                            ]
                        },
                        {
                            id: 'pfcg-authority',
                            label: 'Atualizar Autorizações',
                            icon: 'shield-check',
                            processo: 'Funções PFCG',
                            subprocesso: 'C. PFCG_AUTHORITY.py',
                            prompt: 'Quero atualizar as autorizações de um perfil.',
                            followupText: 'Certo. Vamos preparar a atualização das autorizações do perfil.',
                            children: []
                        },
                        {
                            id: 'cua-adicionar',
                            label: 'Adicionar Utilizador',
                            icon: 'user-plus',
                            processo: 'Funções PFCG',
                            subprocesso: 'H. CUA_ADICIONAR.py',
                            prompt: 'Quero adicionar um utilizador.',
                            followupText: 'Certo. Vamos preparar a adição do utilizador.',
                            children: []
                        },
                        {
                            id: 'cua-enddate',
                            label: 'Alterar Data Fim',
                            icon: 'calendar',
                            processo: 'Funções PFCG',
                            subprocesso: 'I. CUA_ENDDATE.py',
                            prompt: 'Quero alterar a data fim de um utilizador.',
                            followupText: 'Certo. Vamos preparar a alteração da data fim do utilizador.',
                            children: []
                        },
                        {
                            id: 'cua-remove',
                            label: 'Remover Utilizador',
                            icon: 'user-minus',
                            processo: 'Funções PFCG',
                            subprocesso: 'J. CUA_REMOVE.py',
                            prompt: 'Quero remover um utilizador.',
                            followupText: 'Certo. Vamos preparar a remoção do utilizador.',
                            children: []
                        }
                    ]
                }
            ]
        },
        {
            id: 'analise',
            label: 'Análise Geral',
            icon: 'analysis',
            prompt: 'Quero fazer uma análise geral.',
            children: []
        },
        {
            id: 'tickets',
            label: 'Tickets',
            icon: 'tickets',
            prompt: 'Quero consultar ou analisar tickets.',
            children: []
        },
        {
            id: 'testes-unitarios',
            label: 'Testes Unitários',
            icon: 'analysis',
            prompt: 'Quero ver os processos de testes unitários.',
            followupText: 'Escolha um processo de Testes Unitários:',
            followupActionsSource: 'children',
            children: [
                {
                    id: 'testes-unitarios-criar-documento-fi',
                    label: 'Criar Documento FI',
                    icon: 'analysis',
                    processo: 'Testes Unitários',
                    prompt: 'Quero abrir o processo de Criar Documento FI.',
                    followupText: 'Escolha o ambiente SAP:',
                    followupActionsSource: 'children',
                    children: [
                        {
                            id: 'testes-unitarios-criar-documento-fi-dev',
                            label: 'DEV',
                            icon: 'analysis',
                            processo: 'Testes Unitários',
                            environment: 'DEV',
                            prompt: 'Quero trabalhar no ambiente DEV.',
                            followupText: 'Escolha o tipo de documento FI (D/K/S):',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'testes-unitarios-criar-documento-fi-dev-cliente',
                                    label: 'D',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'DEV',
                                    subprocesso: 'A1. Criar Documento FI - Cliente.py',
                                    prompt: 'Quero criar um documento FI de Cliente.',
                                    children: []
                                },
                                {
                                    id: 'testes-unitarios-criar-documento-fi-dev-fornecedor',
                                    label: 'K',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'DEV',
                                    subprocesso: 'A2. Criar Documento FI - Fornecedor.py',
                                    prompt: 'Quero criar um documento FI de Fornecedor.',
                                    children: []
                                },
                                {
                                    id: 'testes-unitarios-criar-documento-fi-dev-razao',
                                    label: 'S',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'DEV',
                                    subprocesso: 'A3. Criar Documento FI - Razao.py',
                                    prompt: 'Quero criar um documento FI de Razão.',
                                    children: []
                                }
                            ]
                        },
                        {
                            id: 'testes-unitarios-criar-documento-fi-qad',
                            label: 'QAD',
                            icon: 'analysis',
                            processo: 'Testes Unitários',
                            environment: 'QAD',
                            prompt: 'Quero trabalhar no ambiente QAD.',
                            followupText: 'Escolha o tipo de conta para validar em QAD:',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'testes-unitarios-criar-documento-fi-qad-cliente',
                                    label: 'Cliente',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'QAD',
                                    prompt: 'Quero validar um documento de Cliente no ambiente QAD.',
                                    followupText: 'Escolha as informações de teste:',
                                    followupActionsSource: 'children',
                                    children: [
                                        {
                                            id: 'testes-unitarios-criar-documento-fi-qad-cliente-default',
                                            label: 'Default',
                                            icon: 'analysis',
                                            processo: 'Testes Unitários',
                                            environment: 'QAD',
                                            subprocesso: 'A1. Criar Documento FI - Cliente - QAD.py',
                                            mode: 'default',
                                            prompt: 'Usar informações Default.',
                                            children: []
                                        },
                                        {
                                            id: 'testes-unitarios-criar-documento-fi-qad-cliente-manual',
                                            label: 'Manual',
                                            icon: 'analysis',
                                            processo: 'Testes Unitários',
                                            environment: 'QAD',
                                            subprocesso: 'A1. Criar Documento FI - Cliente - QAD.py',
                                            mode: 'manual',
                                            prompt: 'Vou informar os campos manualmente.',
                                            children: []
                                        }
                                    ]
                                },
                                {
                                    id: 'testes-unitarios-criar-documento-fi-qad-fornecedor',
                                    label: 'Fornecedor',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'QAD',
                                    prompt: 'Quero validar um documento de Fornecedor no ambiente QAD.',
                                    followupText: 'Escolha as informações de teste:',
                                    followupActionsSource: 'children',
                                    children: [
                                        {
                                            id: 'testes-unitarios-criar-documento-fi-qad-fornecedor-default',
                                            label: 'Default',
                                            icon: 'analysis',
                                            processo: 'Testes Unitários',
                                            environment: 'QAD',
                                            subprocesso: 'A2. Criar Documento FI - Fornecedor - QAD.py',
                                            mode: 'default',
                                            prompt: 'Usar informações Default.',
                                            children: []
                                        },
                                        {
                                            id: 'testes-unitarios-criar-documento-fi-qad-fornecedor-manual',
                                            label: 'Manual',
                                            icon: 'analysis',
                                            processo: 'Testes Unitários',
                                            environment: 'QAD',
                                            subprocesso: 'A2. Criar Documento FI - Fornecedor - QAD.py',
                                            mode: 'manual',
                                            prompt: 'Vou informar os campos manualmente.',
                                            children: []
                                        }
                                    ]
                                },
                                {
                                    id: 'testes-unitarios-criar-documento-fi-qad-razao',
                                    label: 'Razão',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'QAD',
                                    prompt: 'Quero validar um documento de Razão no ambiente QAD.',
                                    followupText: 'Escolha as informações de teste:',
                                    followupActionsSource: 'children',
                                    children: [
                                        {
                                            id: 'testes-unitarios-criar-documento-fi-qad-razao-default',
                                            label: 'Default',
                                            icon: 'analysis',
                                            processo: 'Testes Unitários',
                                            environment: 'QAD',
                                            subprocesso: 'A3. Criar Documento FI - Razao - QAD.py',
                                            mode: 'default',
                                            prompt: 'Usar informações Default.',
                                            children: []
                                        },
                                        {
                                            id: 'testes-unitarios-criar-documento-fi-qad-razao-manual',
                                            label: 'Manual',
                                            icon: 'analysis',
                                            processo: 'Testes Unitários',
                                            environment: 'QAD',
                                            subprocesso: 'A3. Criar Documento FI - Razao - QAD.py',
                                            mode: 'manual',
                                            prompt: 'Vou informar os campos manualmente.',
                                            children: []
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            id: 'testes-unitarios-criar-documento-fi-prd',
                            label: 'PRD',
                            icon: 'analysis',
                            processo: 'Testes Unitários',
                            environment: 'PRD',
                            prompt: 'Quero trabalhar no ambiente PRD.',
                            followupText: 'Escolha o tipo de documento FI:',
                            followupActionsSource: 'children',
                            children: [
                                {
                                    id: 'testes-unitarios-criar-documento-fi-prd-cliente',
                                    label: 'Cliente',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'PRD',
                                    subprocesso: 'A1. Criar Documento FI - Cliente.py',
                                    prompt: 'Quero criar um documento FI de Cliente.',
                                    children: []
                                },
                                {
                                    id: 'testes-unitarios-criar-documento-fi-prd-fornecedor',
                                    label: 'Fornecedor',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'PRD',
                                    subprocesso: 'A2. Criar Documento FI - Fornecedor.py',
                                    prompt: 'Quero criar um documento FI de Fornecedor.',
                                    children: []
                                },
                                {
                                    id: 'testes-unitarios-criar-documento-fi-prd-razao',
                                    label: 'Razão',
                                    icon: 'analysis',
                                    processo: 'Testes Unitários',
                                    environment: 'PRD',
                                    subprocesso: 'A3. Criar Documento FI - Razao.py',
                                    prompt: 'Quero criar um documento FI de Razão.',
                                    children: []
                                }
                            ]
                        }
                    ]
                },
                {
                    id: 'testes-unitarios-executar-f110',
                    label: 'Executar F110',
                    icon: 'analysis',
                    processo: 'Testes Unitários',
                    prompt: 'Quero abrir o processo de Executar F110.',
                    followupText: 'Escolha o ambiente SAP:',
                    followupActionsSource: 'children',
                    children: [
                        {
                            id: 'testes-unitarios-executar-f110-dev',
                            label: 'DEV',
                            icon: 'analysis',
                            processo: 'Testes Unitários',
                            environment: 'DEV',
                            prompt: 'Quero executar o F110 no ambiente DEV.',
                            followupText: 'Escolha o tipo de execução:',
                            followupActionsSource: 'children',
                            children: asiBuildF110DefaultChildren('DEV')
                        },
                        {
                            id: 'testes-unitarios-executar-f110-qad',
                            label: 'QAD',
                            icon: 'analysis',
                            processo: 'Testes Unitários',
                            environment: 'QAD',
                            prompt: 'Quero executar o F110 no ambiente QAD.',
                            followupText: 'Escolha o tipo de execução:',
                            followupActionsSource: 'children',
                            children: asiBuildF110DefaultChildren('QAD')
                        },
                        {
                            id: 'testes-unitarios-executar-f110-prd',
                            label: 'PRD',
                            icon: 'analysis',
                            processo: 'Testes Unitários',
                            environment: 'PRD',
                            prompt: 'Quero executar o F110 no ambiente PRD.',
                            followupText: 'Escolha o tipo de execução:',
                            followupActionsSource: 'children',
                            children: asiBuildF110DefaultChildren('PRD')
                        }
                    ]
                }
            ]
        }
    ];

    function asiDefaultGreeting() {
        return 'Olá Clayton, em que posso ajudá-lo?';
    }

    function asiMockReply() {
        return 'Recebi a sua mensagem. Em breve esta área estará ligada ao motor do Agente Salsa IT.';
    }

    function asiQuickActionIcon(iconName) {
        const icons = {
            settings: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M9.6 3.25h4.8l.48 2.12c.4.13.78.29 1.15.48l1.94-1.03 3.39 3.39-1.03 1.94c.19.37.35.75.48 1.15l2.12.48v4.8l-2.12.48c-.13.4-.29.78-.48 1.15l1.03 1.94-3.39 3.39-1.94-1.03c-.37.19-.75.35-1.15.48l-.48 2.12H9.6l-.48-2.12a7.9 7.9 0 0 1-1.15-.48l-1.94 1.03-3.39-3.39 1.03-1.94a7.9 7.9 0 0 1-.48-1.15l-2.12-.48v-4.8l2.12-.48c.13-.4.29-.78.48-1.15L2.64 8.21l3.39-3.39 1.94 1.03c.37-.19.75-.35 1.15-.48L9.6 3.25Zm2.4 5.15a3.6 3.6 0 1 0 0 7.2 3.6 3.6 0 0 0 0-7.2Z" fill="currentColor"></path>
                </svg>
            `,
            analysis: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M10.2 4.25a5.95 5.95 0 1 1 0 11.9 5.95 5.95 0 0 1 0-11.9Zm0 1.9a4.05 4.05 0 1 0 0 8.1 4.05 4.05 0 0 0 0-8.1Zm7.56 9.97 4.01 4.01a1 1 0 0 1-1.42 1.41l-4-4v-.02l1.41-1.4Zm-8.85-6.1h2.05c.52 0 .95.42.95.95v2.2a.95.95 0 1 1-1.9 0v-1.25H8.91a.95.95 0 0 1 0-1.9Z" fill="currentColor"></path>
                </svg>
            `,
            tickets: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M5 5.25h14c.97 0 1.75.78 1.75 1.75v3.2c-1.26.18-2.22 1.27-2.22 2.58 0 1.31.96 2.4 2.22 2.58V18A1.75 1.75 0 0 1 19 19.75H5A1.75 1.75 0 0 1 3.25 18v-2.64c1.26-.18 2.22-1.27 2.22-2.58 0-1.31-.96-2.4-2.22-2.58V7c0-.97.78-1.75 1.75-1.75Zm5.15 2.9a.95.95 0 0 0 0 1.9h3.7a.95.95 0 1 0 0-1.9h-3.7Zm0 5.8a.95.95 0 0 0 0 1.9h3.1a.95.95 0 1 0 0-1.9h-3.1Z" fill="currentColor"></path>
                </svg>
            `,
            authorization: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M12 3.25l7 2.58v5.47c0 4.4-2.6 8.44-6.63 10.3L12 21.9l-.37-.16C7.6 19.88 5 15.84 5 11.3V5.83l7-2.58Zm0 2.03-5.1 1.88v4.14c0 3.53 2 6.76 5.1 8.42 3.1-1.66 5.1-4.89 5.1-8.42V7.16L12 5.28Zm-.9 4.12a.95.95 0 0 1 1.9 0v1.05h.45a1.8 1.8 0 0 1 1.8 1.8v2.4a1.8 1.8 0 0 1-1.8 1.8h-2.9a1.8 1.8 0 0 1-1.8-1.8v-2.4a1.8 1.8 0 0 1 1.8-1.8h.55V9.4Zm-.55 2.95v2h2.6v-2h-2.6Z" fill="currentColor"></path>
                </svg>
            `,
            'shield-plus': `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M12 3.2 19 5.8v5.45c0 4.42-2.63 8.48-6.69 10.33L12 21.9l-.31-.13C7.63 19.92 5 15.86 5 11.25V5.8l7-2.6Zm0 2.03-5.08 1.89v4.13c0 3.54 2.02 6.78 5.08 8.43 3.06-1.65 5.08-4.89 5.08-8.43V7.12L12 5.23Zm-.95 3.87a.95.95 0 1 1 1.9 0v1.95h1.95a.95.95 0 1 1 0 1.9h-1.95v1.95a.95.95 0 1 1-1.9 0v-1.95H9.1a.95.95 0 1 1 0-1.9h1.95V9.1Z" fill="currentColor"></path>
                </svg>
            `,
            trash: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M9.1 3.75h5.8c.92 0 1.68.68 1.81 1.56l.09.59H20a.95.95 0 1 1 0 1.9h-.96l-.67 10.16a2.1 2.1 0 0 1-2.1 1.96H7.73a2.1 2.1 0 0 1-2.1-1.96L4.96 7.8H4a.95.95 0 1 1 0-1.9h3.2l.09-.59c.13-.88.89-1.56 1.81-1.56Zm.82 2.15-.02.01H14.1l-.02-.01-.05-.35h-4.08l-.05.35Zm-2.35 3.2a.85.85 0 0 1 .85.85v5.8a.85.85 0 1 1-1.7 0v-5.8a.85.85 0 0 1 .85-.85Zm4.43 0a.85.85 0 0 1 .85.85v5.8a.85.85 0 1 1-1.7 0v-5.8a.85.85 0 0 1 .85-.85Zm4.43 0a.85.85 0 0 1 .85.85v5.8a.85.85 0 1 1-1.7 0v-5.8a.85.85 0 0 1 .85-.85Z" fill="currentColor"></path>
                </svg>
            `,
            'shield-check': `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M12 3.25l7 2.58v5.47c0 4.4-2.6 8.44-6.63 10.3L12 21.9l-.37-.16C7.6 19.88 5 15.84 5 11.3V5.83l7-2.58Zm0 2.03-5.1 1.88v4.14c0 3.53 2 6.76 5.1 8.42 3.1-1.66 5.1-4.89 5.1-8.42V7.16L12 5.28Zm3.4 4.92a.95.95 0 0 1 .12 1.34l-3.2 3.85a.95.95 0 0 1-1.42.07l-1.7-1.7a.95.95 0 0 1 1.34-1.34l.96.96 2.56-3.07a.95.95 0 0 1 1.34-.11Z" fill="currentColor"></path>
                </svg>
            `,
            layers: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M12 3.3 20.15 7.9a.9.9 0 0 1 0 1.57L12 14.07 3.85 9.47a.9.9 0 0 1 0-1.57L12 3.3Zm0 2.02L6.15 8.69 12 11.98l5.85-3.29L12 5.32Zm-8.15 7.2 7.72 4.34a.9.9 0 0 0 .86 0l7.72-4.34a.9.9 0 1 1 .88 1.57l-7.72 4.34a2.7 2.7 0 0 1-2.64 0l-7.72-4.34a.9.9 0 0 1 .9-1.57Zm0 4 7.72 4.34a.9.9 0 0 0 .86 0l7.72-4.34a.9.9 0 1 1 .88 1.57l-7.72 4.34a2.7 2.7 0 0 1-2.64 0l-7.72-4.34a.9.9 0 0 1 .9-1.57Z" fill="currentColor"></path>
                </svg>
            `,
            'user-plus': `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M9 4.3a4.15 4.15 0 1 1 0 8.3 4.15 4.15 0 0 1 0-8.3Zm0 1.9a2.25 2.25 0 1 0 0 4.5 2.25 2.25 0 0 0 0-4.5Zm0 8.45c3.4 0 6.25 1.84 7.18 4.45a.95.95 0 1 1-1.79.64c-.63-1.76-2.73-3.19-5.39-3.19-2.67 0-4.76 1.43-5.39 3.19a.95.95 0 1 1-1.79-.64c.93-2.61 3.77-4.45 7.18-4.45Zm8.25-6.05a.95.95 0 0 1 .95.95v1.7h1.7a.95.95 0 1 1 0 1.9h-1.7v1.7a.95.95 0 1 1-1.9 0v-1.7h-1.7a.95.95 0 0 1 0-1.9h1.7v-1.7a.95.95 0 0 1 .95-.95Z" fill="currentColor"></path>
                </svg>
            `,
            calendar: `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M7.1 3.8a.95.95 0 0 1 .95.95v.85h7.9v-.85a.95.95 0 1 1 1.9 0v.85h.45A2.7 2.7 0 0 1 21 8.3v9A2.7 2.7 0 0 1 18.3 20h-12A2.7 2.7 0 0 1 3.6 17.3v-9a2.7 2.7 0 0 1 2.7-2.7h.45v-.85a.95.95 0 0 1 .95-.95Zm11.3 5.45h-12.8v8.05c0 .44.36.8.8.8h12c.44 0 .8-.36.8-.8V9.25Zm-8.35 2.4h4.2a.95.95 0 1 1 0 1.9h-4.2a.95.95 0 0 1 0-1.9Z" fill="currentColor"></path>
                </svg>
            `,
            'user-minus': `
                <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
                    <path d="M9 4.3a4.15 4.15 0 1 1 0 8.3 4.15 4.15 0 0 1 0-8.3Zm0 1.9a2.25 2.25 0 1 0 0 4.5 2.25 2.25 0 0 0 0-4.5Zm0 8.45c3.4 0 6.25 1.84 7.18 4.45a.95.95 0 1 1-1.79.64c-.63-1.76-2.73-3.19-5.39-3.19-2.67 0-4.76 1.43-5.39 3.19a.95.95 0 1 1-1.79-.64c.93-2.61 3.77-4.45 7.18-4.45Zm5.6-4.35h5.3a.95.95 0 0 1 0 1.9h-5.3a.95.95 0 1 1 0-1.9Z" fill="currentColor"></path>
                </svg>
            `
        };

        return icons[iconName] || icons.analysis;
    }

    function asiFindQuickAction(actionId, actions = salsaAgentActions) {
        for (const action of actions) {
            if (action.id === actionId) return action;
            if (action.children && action.children.length > 0) {
                const childMatch = asiFindQuickAction(actionId, action.children);
                if (childMatch) return childMatch;
            }
        }
        return null;
    }

    function asiFindActionPath(actionId, actions = salsaAgentActions, trail = []) {
        for (const action of actions) {
            const nextTrail = [...trail, action];
            if (action.id === actionId) return nextTrail;
            if (action.children && action.children.length > 0) {
                const childPath = asiFindActionPath(actionId, action.children, nextTrail);
                if (childPath) return childPath;
            }
        }
        return null;
    }

    function asiCreateMessage(role, text, options = {}) {
        asiChatMessageCounter += 1;

        return {
            id: `asi-msg-${asiChatMessageCounter}`,
            role,
            text,
            html: typeof options.html === 'string' ? options.html : '',
            belowBubbleHtml: typeof options.belowBubbleHtml === 'string' ? options.belowBubbleHtml : '',
            bubbleClassName: typeof options.bubbleClassName === 'string' ? options.bubbleClassName : '',
            isProcessing: Boolean(options.isProcessing),
            actions: Array.isArray(options.actions) ? options.actions : [],
            actionLevel: Number(options.actionLevel || 0),
            parentActionId: options.parentActionId || '',
            selectionGroupKey: options.selectionGroupKey || (options.parentActionId || '__root__')
        };
    }

    function asiRenderQuickActionButtons(actions, level = 0, parentActionId = '', selectionGroupKey = '__root__') {
        if (!actions || actions.length === 0) return '';
        const isSubLevel = level > 0;

        return `
            <div class="agent-salsa-quick-actions ${isSubLevel ? 'agent-salsa-quick-actions-sub' : 'agent-salsa-quick-actions-root'}" role="group" aria-label="${isSubLevel ? 'Subopções do agente' : 'Ações rápidas do agente'}">
                ${actions.map(action => {
                    const isSelected = asiSelectedActions[selectionGroupKey] === action.id;
                    const selectedClass = isSelected ? ' selected' : '';
                    const subClass = isSubLevel ? ' agent-salsa-quick-action-sub' : '';

                    return `
                        <button
                            type="button"
                            class="agent-salsa-quick-action${subClass}${selectedClass}"
                            data-agent-action-id="${escapeHtml(action.id)}"
                            data-agent-action-level="${level}"
                            data-agent-parent-action-id="${escapeHtml(parentActionId)}"
                            data-agent-selection-group-key="${escapeHtml(selectionGroupKey)}"
                        >
                            <span class="agent-salsa-quick-action-icon">${asiQuickActionIcon(action.icon)}</span>
                            <span>${escapeHtml(action.label)}</span>
                        </button>
                    `;
                }).join('')}
            </div>
        `;
    }

    function asiRenderQuickActions() {
        return asiRenderQuickActionButtons(salsaAgentActions, 0, '', '__root__');
    }

    function asiPresentMainMenu() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiStopPfcgPolling();
        asiSelectedActions = { '__root__': null };
        asiConversationState = asiDefaultConversationState();
        asiChatHistory = [
            asiCreateMessage('assistant', asiDefaultGreeting(), {
                actions: salsaAgentActions,
                actionLevel: 0,
                selectionGroupKey: '__root__'
            })
        ];
        asiChatInitialized = true;
        asiRenderMessages();
        asiUpdateComposerState();
        const { input } = asiGetElements();
        if (input) input.focus();
    }

    function asiReturnToQuickActionMenu(actionId) {
        const action = asiFindQuickAction(actionId, salsaAgentActions);
        if (!action) return;

        const actionPath = asiFindActionPath(actionId, salsaAgentActions) || [action];
        const nextSelectedActions = { '__root__': null };

        if (actionPath.length > 0) {
            nextSelectedActions['__root__'] = actionPath[0].id;
            for (let i = 1; i < actionPath.length; i += 1) {
                nextSelectedActions[actionPath[i - 1].id] = actionPath[i].id;
            }
        }

        asiSelectedActions = nextSelectedActions;
        asiConversationState = asiDefaultConversationState();

        const actions = Array.isArray(action.children) ? [...action.children] : [];
        if (action.id === 'testes-unitarios') {
            actions.push(ASI_MAIN_MENU_ACTION);
        }

        asiAppendMessage(asiCreateMessage(
            'assistant',
            action.followupText || 'Como deseja prosseguir?',
            {
                actions,
                actionLevel: actionPath.length,
                parentActionId: action.id,
                selectionGroupKey: action.id
            }
        ));
        asiUpdateComposerState();
    }

    function asiGetElements() {
        return {
            messages: document.getElementById('agent-salsa-it-messages'),
            input: document.getElementById('agent-salsa-it-input'),
            send: document.getElementById('agent-salsa-it-send'),
            context: document.getElementById('agent-salsa-fi-context')
        };
    }

    function asiStopPfcgPolling() {
        if (asiPfcgPollingTimer) {
            clearInterval(asiPfcgPollingTimer);
            asiPfcgPollingTimer = null;
        }
        asiPfcgPollingInFlight = false;
    }

    function asiGetComposerPlaceholder() {
        if (asiConversationState.awaitingInput === ASI_PFCG_AWAITING_INPUT) {
            return ASI_PFCG_PLACEHOLDER;
        }
        if (asiConversationState.awaitingInput === ASI_PFCG_COMPOSTA_ROLE_NAME_INPUT) {
            const normalizedRoleName = asiNormalizePfcgRoleName(rawMessage);
            if (!normalizedRoleName) {
                asiAppendMessage(asiCreateMessage('assistant', 'Envie o Nome da Função Composta que vamos criar'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_COMPOSTA_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (normalizedRoleName.length > 30) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    `O nome da Função Composta não pode ultrapassar o tamanho máximo de 30 caracteres (tem ${normalizedRoleName.length} caracteres).\nPor favor, corrija o nome e envie novamente.`
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_COMPOSTA_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (!asiIsValidPfcgRoleName(normalizedRoleName)) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    'O nome da Função Composta contém caracteres inválidos.\nUtilize apenas letras, números, "_", "-", "/" ou ":".\nPor favor, corrija o nome e envie novamente.'
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_COMPOSTA_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreateRoleName: normalizedRoleName,
                pfcgCreateDescription: '',
                awaitingInput: ASI_PFCG_COMPOSTA_DESCRIPTION_INPUT,
                isBusy: false
            };
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Qual é a descrição da Função Composta?'
            ));
            asiUpdateComposerState();
            input.focus();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_COMPOSTA_DESCRIPTION_INPUT) {
            const description = rawMessage.trim();
            if (!description) {
                asiAppendMessage(asiCreateMessage('assistant', 'Informe uma descrição válida para a Função Composta.'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_COMPOSTA_DESCRIPTION_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (description.length > 80) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    `A descrição não pode ultrapassar o tamanho máximo de 80 caracteres (tem ${description.length} caracteres).\nPor favor, corrija e envie novamente.`
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_COMPOSTA_DESCRIPTION_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreateDescription: description,
                awaitingInput: ASI_PFCG_COMPOSTA_CHILDREN_INPUT,
                isBusy: false
            };
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Envie as funções componentes (roles filhas) separadas por vírgula ou espaço (Ex.: Z_ROLE_01, Z_ROLE_02).'
            ));
            asiUpdateComposerState();
            input.focus();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_COMPOSTA_CHILDREN_INPUT) {
            const childRoles = asiNormalizePfcgTcodes(rawMessage);
            if (childRoles.length === 0) {
                asiAppendMessage(asiCreateMessage('assistant', 'Informe pelo menos uma função componente (role filha) válida. Separe por vírgula (ex.: Z_FI_01, Z_FI_02).'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_COMPOSTA_CHILDREN_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreateChildRoles: childRoles,
                awaitingInput: '',
                isBusy: false
            };
            asiUpdateComposerState();
            asiAskPfcgTransportMode();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT) {
            return ASI_PFCG_INDIVIDUAL_ROLE_NAME_PLACEHOLDER;
        }
        if (asiConversationState.awaitingInput === ASI_PFCG_INDIVIDUAL_DESCRIPTION_INPUT) {
            return ASI_PFCG_INDIVIDUAL_DESCRIPTION_PLACEHOLDER;
        }
        if (asiConversationState.awaitingInput === ASI_PFCG_INDIVIDUAL_TCODES_INPUT) {
            return ASI_PFCG_INDIVIDUAL_TCODES_PLACEHOLDER;
        }
        if (asiConversationState.awaitingInput === ASI_PFCG_TRANSPORT_CREATE_DESCRIPTION_INPUT) {
            return 'Ex.: Criação de função PFCG em DEV';
        }
        if (asiConversationState.awaitingInput === ASI_PFCG_DELETE_ROLE_NAME_INPUT) {
            return ASI_PFCG_DELETE_ROLE_NAME_PLACEHOLDER;
        }
        if (asiConversationState.awaitingInput === ASI_PFCG_DELETE_TRANSPORT_CREATE_DESCRIPTION_INPUT) {
            return 'Ex.: Eliminação de função PFCG em DEV';
        }
        return ASI_DEFAULT_PLACEHOLDER;
    }

    function asiNormalizePfcgTcodes(raw) {
        const text = String(raw || '').replace(/\r/g, '\n').replace(/\t/g, ' ');
        const parts = text.split(/[;,\n ]+/);
        const out = [];
        const seen = new Set();
        for (const part of parts) {
            let value = String(part || '').trim().toUpperCase();
            if (!value) continue;
            if (value.startsWith('/N') || value.startsWith('/O')) {
                value = value.slice(2).trim();
            }
            if (!value || seen.has(value)) continue;
            seen.add(value);
            out.push(value);
        }
        return out;
    }

    function asiNormalizePfcgRoleName(value) {
        return String(value || '').trim().toUpperCase();
    }

    function asiIsValidPfcgRoleName(roleName) {
        return Boolean(roleName) && ASI_PFCG_ROLE_PATTERN.test(roleName);
    }

    function asiBuildPfcgProcessingHtml(roleName) {
        return `
            <div style="display:flex;align-items:center;gap:10px;">
                <svg width="14" height="14" viewBox="0 0 50 50" aria-hidden="true" style="flex:0 0 auto;animation: spin 1s linear infinite;color:#94a3b8;">
                    <circle cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" style="stroke-dasharray: 1, 150; stroke-dashoffset: 0; animation: dash 1.5s ease-in-out infinite;"></circle>
                </svg>
                <span style="font-weight:600;">A analisar ${escapeHtml(roleName)} no SAP PRD...</span>
            </div>
        `;
    }

    function asiBuildPfcgDetailRow(label, value) {
        if (value === null || value === undefined || value === '') return '';
        return `
            <div class="asi-pfcg-summary-item">
                <div style="font-size:0.72rem;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:var(--text-secondary);">${escapeHtml(label)}</div>
                <div style="font-size:0.84rem;font-weight:600;color:var(--text-primary);word-break:break-word;">${escapeHtml(String(value))}</div>
            </div>
        `;
    }

    function asiBuildPfcgExcelSummaryItem(label, value) {
        if (value === null || value === undefined || value === '') return '';
        return `
            <div class="asi-pfcg-excel-summary-item">
                <div class="asi-pfcg-excel-summary-label">${escapeHtml(String(label).toUpperCase())}</div>
                <div class="asi-pfcg-excel-summary-value">${escapeHtml(String(value))}</div>
            </div>
        `;
    }

    function asiEnsurePfcgResultStyles() {
        if (document.getElementById('asi-pfcg-result-styles')) return;
        const style = document.createElement('style');
        style.id = 'asi-pfcg-result-styles';
        style.textContent = `
            .asi-pfcg-result-card {
                white-space: normal;
                width: 100%;
            }

            .asi-pfcg-result-shell {
                margin-top: 8px;
                padding-top: 8px;
                border-top: 1px solid rgba(148, 163, 184, 0.18);
            }

            .asi-pfcg-summary-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 12px 18px;
                margin-top: 14px;
                padding-top: 10px;
                border-top: 1px solid rgba(148, 163, 184, 0.2);
            }

            .asi-pfcg-summary-item {
                display: grid;
                grid-template-columns: 92px minmax(0, 1fr);
                gap: 12px;
                align-items: start;
                min-width: 0;
            }

            .asi-pfcg-excel-summary-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                gap: 16px;
                width: 100%;
                align-items: start;
            }

            .asi-pfcg-excel-summary-item {
                display: flex;
                flex-direction: column;
                gap: 4px;
                min-width: 0;
            }

            .asi-pfcg-excel-summary-label {
                font-size: 11px;
                line-height: 1.2;
                letter-spacing: 0.08em;
                text-transform: uppercase;
                color: var(--text-secondary);
                font-weight: 700;
            }

            .asi-pfcg-excel-summary-value {
                min-width: 0;
                overflow-wrap: anywhere;
                word-break: break-word;
                line-height: 1.35;
                color: var(--text-primary);
                font-size: 13px;
                font-weight: 600;
            }

            .asi-pfcg-excel-summary-note {
                margin-top: 10px;
                font-size: 0.82rem;
                color: var(--text-secondary);
                line-height: 1.35;
                overflow-wrap: anywhere;
            }

            .pfcg-create-grid {
                display: grid;
                grid-template-columns: 1fr;
                gap: 12px;
                width: 100%;
                align-items: stretch;
            }

            .asi-pfcg-result-heading-row {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 12px;
                flex-wrap: wrap;
            }

            .asi-pfcg-result-grid {
                display: grid;
                grid-template-columns: repeat(4, minmax(0, 1fr));
                column-gap: 20px;
                row-gap: 14px;
                margin-top: 8px;
                align-items: start;
                width: 100%;
            }

            .asi-pfcg-result-field {
                min-width: 0;
                overflow-wrap: anywhere;
            }

            .asi-pfcg-result-field--span2 {
                grid-column: span 2;
            }

            .asi-pfcg-result-value--nowrap {
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
            }

            .asi-pfcg-result-label {
                display: block;
                margin: 0 0 2px;
                font-size: 0.7rem;
                font-weight: 700;
                letter-spacing: 0.05em;
                text-transform: uppercase;
                color: var(--text-secondary);
                line-height: 1.15;
            }

            .asi-pfcg-result-value {
                display: block;
                margin: 0;
                font-size: 0.84rem;
                font-weight: 600;
                color: var(--text-primary);
                line-height: 1.25;
                word-break: break-word;
                overflow-wrap: anywhere;
            }

            .asi-pfcg-result-pill {
                display: inline-flex;
                align-items: center;
                gap: 6px;
                padding: 2px 7px;
                line-height: 1;
                border-radius: 999px;
                font-size: 0.68rem;
                font-weight: 800;
                letter-spacing: 0.03em;
                text-transform: uppercase;
                background: rgba(148, 163, 184, 0.12);
                color: var(--text-secondary);
                white-space: nowrap;
            }

            .asi-pfcg-result-pill--success {
                background: rgba(22, 163, 74, 0.12);
                color: #16a34a;
            }

            .asi-pfcg-result-pill--warning {
                background: rgba(245, 158, 11, 0.12);
                color: #f59e0b;
            }

            .asi-pfcg-result-heading {
                display: flex;
                align-items: center;
                gap: 8px;
                font-size: 0.86rem;
                font-weight: 700;
                line-height: 1.2;
            }

            .asi-pfcg-result-note {
                margin-top: 8px;
                font-size: 0.78rem;
                line-height: 1.3;
                color: var(--text-secondary);
            }

            @media (max-width: 900px) {
                .asi-pfcg-excel-summary-grid {
                    grid-template-columns: repeat(2, minmax(0, 1fr));
                }
                .asi-pfcg-result-grid {
                    grid-template-columns: repeat(2, minmax(0, 1fr));
                }
                .asi-pfcg-summary-grid {
                    grid-template-columns: 1fr;
                }
            }

            @media (max-width: 550px) {
                .asi-pfcg-excel-summary-grid {
                    grid-template-columns: 1fr;
                }
                .asi-pfcg-result-grid {
                    grid-template-columns: 1fr;
                }
                .asi-pfcg-result-field--span2 {
                    grid-column: span 1;
                }
                .asi-pfcg-summary-grid {
                    grid-template-columns: 1fr;
                }
            }
        `;
        document.head.appendChild(style);
    }

    function asiBuildPfcgResultField(label, value, extraValueClass, isSpan2 = false) {
        if (value === null || value === undefined || value === '') return '';
        const valueClass = extraValueClass ? `asi-pfcg-result-value ${extraValueClass}` : 'asi-pfcg-result-value';
        const spanClass = isSpan2 ? ' asi-pfcg-result-field--span2' : '';
        return `
            <div class="asi-pfcg-result-field${spanClass}">
                <span class="asi-pfcg-result-label">${escapeHtml(label)}</span>
                <span class="${valueClass}">${escapeHtml(String(value))}</span>
            </div>
        `;
    }

    function asiBuildPfcgSuccessHtml(result) {
        asiEnsurePfcgResultStyles();
        const isExisting = result.status === 'EXISTE';
        const heading = isExisting
            ? '✓ Função encontrada em PRD'
            : '○ Função não encontrada em PRD';
        const accent = isExisting ? '#16a34a' : '#f59e0b';
        const roleInExcel = result.role_in_excel || result.role || '';
        const topRightStatus = `<span class="asi-pfcg-result-pill ${isExisting ? 'asi-pfcg-result-pill--success' : 'asi-pfcg-result-pill--warning'}">${escapeHtml(result.status || '')}</span>`;
        const note = result.status === 'NAO_EXISTE'
            ? `<div class="asi-pfcg-result-note">Podemos continuar com a preparação para criação do Perfil de Autorização.</div>`
            : '';

        const rowFields = isExisting
            ? [
                asiBuildPfcgResultField('Função', result.role || roleInExcel, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Descrição', result.description, '', true),
                asiBuildPfcgResultField('Idioma', result.language),
                asiBuildPfcgResultField('Sistema', result.system),
                asiBuildPfcgResultField('Cliente', result.client)
            ].join('')
            : [
                asiBuildPfcgResultField('Função', result.role || roleInExcel, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Sistema', result.system),
                asiBuildPfcgResultField('Cliente', result.client)
            ].join('');

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:${accent};">${heading}</div>
                    ${topRightStatus}
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-grid">
                        ${rowFields}
                    </div>
                    ${note}
                </div>
            </div>
        `;
    }

    function asiBuildPfcgErrorHtml(message, detail = '') {
        const safeDetail = detail ? `<div style="margin-top:10px;font-size:0.78rem;color:var(--text-secondary);">${escapeHtml(detail)}</div>` : '';
        return `
            <div>
                <div style="font-weight:700;color:#dc2626;">${escapeHtml(message)}</div>
                ${safeDetail}
            </div>
        `;
    }

    function asiBuildPfcgGenericProcessingHtml(message) {
        return `
            <div style="display:flex;align-items:center;gap:10px;">
                <svg width="14" height="14" viewBox="0 0 50 50" aria-hidden="true" style="flex:0 0 auto;animation: spin 1s linear infinite;color:#94a3b8;">
                    <circle cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" style="stroke-dasharray: 1, 150; stroke-dashoffset: 0; animation: dash 1.5s ease-in-out infinite;"></circle>
                </svg>
                <span style="font-weight:600;">${escapeHtml(message)}</span>
            </div>
        `;
    }

    function asiBuildThinkingIndicatorHtml() {
        return `
            <div class="agent-salsa-thinking-indicator" role="status" aria-label="A processar">
                <span class="agent-salsa-thinking-dot"></span>
                <span class="agent-salsa-thinking-dot"></span>
                <span class="agent-salsa-thinking-dot"></span>
            </div>
        `;
    }

    function asiPfcgTransportModeLabel(mode) {
        const normalized = String(mode || 'LOCAL').toUpperCase();
        if (normalized === 'CREATE_REQUEST') return 'Nova Request';
        if (normalized === 'EXISTING_REQUEST') return 'Request existente';
        return 'Local (sem transporte)';
    }

    function asiBuildPfcgIndividualPreviewHtml(result) {
        asiEnsurePfcgResultStyles();
        const tcodes = Array.isArray(result.tcodes) ? result.tcodes : [];
        const transport = result.transport || {};
        const transportLabel = asiPfcgTransportModeLabel(transport.transport_mode);
        const transportValue = transport.request_number
            ? `${transportLabel} — ${transport.request_number}`
            : transportLabel;
        const isComposta = result.tipo === 'Função Composta' || Array.isArray(result.child_roles);
        const childRoles = Array.isArray(result.child_roles) ? result.child_roles : [];
        const rowFields = isComposta
            ? [
                asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                asiBuildPfcgResultField('Função Composta', result.role, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Descrição', result.description),
                asiBuildPfcgResultField('Total Componentes', childRoles.length),
                asiBuildPfcgResultField('Funções Componentes', childRoles.join(', '), '', true),
                asiBuildPfcgResultField('Transporte', transportValue, '', true)
            ].join('')
            : [
                asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                asiBuildPfcgResultField('Função', result.role, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Descrição', result.description),
                asiBuildPfcgResultField('Total', result.tcodes_count != null ? result.tcodes_count : tcodes.length),
                asiBuildPfcgResultField('Transações', tcodes.join(', '), '', true),
                asiBuildPfcgResultField('Transporte', transportValue, '', true)
            ].join('');

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:#f59e0b;">⚠ Confirme antes de criar em DEV</div>
                    <span class="asi-pfcg-result-pill asi-pfcg-result-pill--warning">PREVIEW</span>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-grid">
                        ${rowFields}
                    </div>
                    <div class="asi-pfcg-result-note">Esta ação cria uma função PFCG real em DEV via RFC (sem SAP GUI). Reveja os dados antes de confirmar.</div>
                </div>
            </div>
        `;
    }

    function asiBuildPfcgDeletePreviewHtml(result) {
        asiEnsurePfcgResultStyles();
        const tcodes = Array.isArray(result.tcodes) ? result.tcodes : [];
        const transport = result.transport || {};
        const transportLabel = asiPfcgTransportModeLabel(transport.transport_mode);
        const transportValue = transport.request_number
            ? `${transportLabel} — ${transport.request_number}`
            : transportLabel;
        const usersCount = Number(result.users_count || 0);
        const usersValue = usersCount > 0
            ? `<span style="color:#dc2626;font-weight:700;">⚠ ${usersCount} utilizador(es) atribuído(s)</span>`
            : '0 utilizadores';

        const rowFields = [
            asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
            asiBuildPfcgResultField('Função', result.role, 'asi-pfcg-result-value--nowrap'),
            asiBuildPfcgResultField('Descrição', result.description),
            asiBuildPfcgResultField('Utilizadores', usersValue),
            asiBuildPfcgResultField('Transações', tcodes.length ? tcodes.join(', ') : '(Nenhuma)', '', true),
            asiBuildPfcgResultField('Transporte', transportValue, '', true)
        ].join('');

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:#dc2626;">⚠ Confirme a eliminação em DEV</div>
                    <span class="asi-pfcg-result-pill asi-pfcg-result-pill--warning" style="background:rgba(220,38,38,0.12);color:#dc2626;">DELETE PREVIEW</span>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-grid">
                        ${rowFields}
                    </div>
                    <div class="asi-pfcg-result-note" style="color:#dc2626;font-weight:600;">Esta ação ELIMINA permanentemente a função PFCG em DEV via RFC. Reveja com atenção antes de confirmar.</div>
                </div>
            </div>
        `;
    }

    function asiBuildPfcgDeleteResultHtml(result) {
        asiEnsurePfcgResultStyles();
        const ok = result.ok === true;
        const status = String(result.status || (ok ? 'DELETED' : 'ERROR'));
        const accent = ok ? '#16a34a' : '#dc2626';
        const heading = ok ? '✓ Função eliminada em DEV' : '✗ Falha na eliminação em DEV';
        const pillClass = ok ? 'asi-pfcg-result-pill--success' : 'asi-pfcg-result-pill--warning';

        const resultTransportLabel = asiPfcgTransportModeLabel(result.transport_mode);
        const resultTransportValue = result.transport_request
            ? `${resultTransportLabel} — ${result.transport_request}`
            : resultTransportLabel;

        const rowFields = ok
            ? [
                asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                asiBuildPfcgResultField('Função', result.role, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Transporte', resultTransportValue, '', true)
            ].join('')
            : [
                asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                asiBuildPfcgResultField('Função', result.role, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Tipo de erro', result.error_type, '', true)
            ].join('');

        const message = !ok && result.message
            ? `<div class="asi-pfcg-result-note">${escapeHtml(String(result.message))}</div>`
            : '';

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:${accent};">${heading}</div>
                    <span class="asi-pfcg-result-pill ${pillClass}">${escapeHtml(status)}</span>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-grid">
                        ${rowFields}
                    </div>
                    ${message}
                </div>
            </div>
        `;
    }

    function asiBuildPfcgIndividualResultHtml(result) {
        asiEnsurePfcgResultStyles();
        const ok = result.ok === true;
        const status = String(result.status || (ok ? 'CREATED' : 'ERROR'));
        const isPartial = status === 'PARTIAL_FAILURE';
        const accent = ok ? '#16a34a' : (isPartial ? '#f59e0b' : '#dc2626');
        const heading = ok ? '✓ Função criada em DEV' : (isPartial ? '⚠ Criação parcial em DEV' : '✗ Falha na criação em DEV');
        const pillClass = ok ? 'asi-pfcg-result-pill--success' : 'asi-pfcg-result-pill--warning';

        const resultTransportLabel = asiPfcgTransportModeLabel(result.transport_mode);
        const resultTransportValue = result.transport_request
            ? `${resultTransportLabel} — ${result.transport_request}`
            : resultTransportLabel;
        const isCompostaResult = result.tipo === 'Função Composta' || Array.isArray(result.child_roles);
        const childRolesRes = Array.isArray(result.child_roles) ? result.child_roles : [];
        const rowFields = ok
            ? (isCompostaResult
                ? [
                    asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                    asiBuildPfcgResultField('Função Composta', result.role, 'asi-pfcg-result-value--nowrap'),
                    asiBuildPfcgResultField('Descrição', result.description),
                    asiBuildPfcgResultField('Total Componentes', childRolesRes.length),
                    asiBuildPfcgResultField('Funções Componentes', childRolesRes.join(', '), '', true),
                    asiBuildPfcgResultField('Transporte', resultTransportValue, '', true)
                ].join('')
                : [
                    asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                    asiBuildPfcgResultField('Função', result.role, 'asi-pfcg-result-value--nowrap'),
                    asiBuildPfcgResultField('Descrição', result.description),
                    asiBuildPfcgResultField('Perfil gerado', result.profile_generated === true ? 'Sim' : (result.profile_generated === false ? 'Não' : '')),
                    asiBuildPfcgResultField('Transações pedidas', result.tcodes_requested),
                    asiBuildPfcgResultField('Transações criadas', result.tcodes_created),
                    asiBuildPfcgResultField('Transporte', resultTransportValue, '', true)
                ].join(''))
            : [
                asiBuildPfcgResultField('Ambiente', result.environment || 'DEV'),
                asiBuildPfcgResultField('Função', result.role, 'asi-pfcg-result-value--nowrap'),
                asiBuildPfcgResultField('Tipo de erro', result.error_type, '', true)
            ].join('');

        const message = !ok && result.message
            ? `<div class="asi-pfcg-result-note">${escapeHtml(String(result.message))}</div>`
            : '';

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:${accent};">${heading}</div>
                    <span class="asi-pfcg-result-pill ${pillClass}">${escapeHtml(status)}</span>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-grid">
                        ${rowFields}
                    </div>
                    ${message}
                </div>
            </div>
        `;
    }

    function asiEnsurePfcgListStyles() {
        if (document.getElementById('asi-pfcg-list-styles')) return;
        const style = document.createElement('style');
        style.id = 'asi-pfcg-list-styles';
        style.textContent = `
            .asi-pfcg-list-wrap {
                margin-top: 6px;
            }

            .asi-pfcg-list-scroll {
                max-height: 300px;
                overflow-y: auto;
                border: 1px solid rgba(148, 163, 184, 0.18);
                border-radius: 8px;
            }

            .asi-pfcg-list-table {
                width: 100%;
                border-collapse: collapse;
            }

            .asi-pfcg-list-table th,
            .asi-pfcg-list-table td {
                padding: 6px 10px;
                font-size: 0.78rem;
                text-align: left;
                border-bottom: 1px solid rgba(148, 163, 184, 0.12);
                overflow-wrap: anywhere;
            }

            .asi-pfcg-list-table th {
                position: sticky;
                top: 0;
                font-size: 0.68rem;
                font-weight: 700;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                color: var(--text-secondary);
                background: var(--bg-secondary, #f8fafc);
            }

            .asi-pfcg-list-table td {
                color: var(--text-primary);
                font-weight: 600;
            }

            .asi-pfcg-list-table tbody tr:last-child td {
                border-bottom: none;
            }

            .asi-pfcg-list-footer {
                margin-top: 8px;
                font-size: 0.78rem;
                font-weight: 700;
                color: var(--text-secondary);
            }

            .asi-pfcg-list-empty {
                padding: 14px 10px;
                font-size: 0.82rem;
                color: var(--text-secondary);
                text-align: center;
            }

            .asi-pfcg-status-pill {
                display: inline-flex;
                align-items: center;
                padding: 1px 7px;
                border-radius: 999px;
                font-size: 0.66rem;
                font-weight: 800;
                letter-spacing: 0.03em;
                text-transform: uppercase;
                white-space: nowrap;
            }

            .asi-pfcg-status-pill--ativo { background: rgba(22, 163, 74, 0.12); color: #16a34a; }
            .asi-pfcg-status-pill--futuro { background: rgba(37, 99, 235, 0.12); color: #2563eb; }
            .asi-pfcg-status-pill--expirado { background: rgba(220, 38, 38, 0.12); color: #dc2626; }
        `;
        document.head.appendChild(style);
    }

    function asiPfcgStatusPillHtml(status) {
        const normalized = String(status || '').toUpperCase();
        const cls = normalized === 'ATIVO' ? 'asi-pfcg-status-pill--ativo'
            : normalized === 'FUTURO' ? 'asi-pfcg-status-pill--futuro'
            : normalized === 'EXPIRADO' ? 'asi-pfcg-status-pill--expirado'
            : '';
        return `<span class="asi-pfcg-status-pill ${cls}">${escapeHtml(normalized || '-')}</span>`;
    }

    function asiBuildPfcgRoleCompositeNoteHtml(result) {
        if (!result || !result.is_composite) return '';
        const members = Array.isArray(result.composite_members) ? result.composite_members : [];
        return `<div class="asi-pfcg-result-note">Função composta — componentes considerados: ${escapeHtml(members.join(', '))}</div>`;
    }

    function asiBuildPfcgRoleWarningNoteHtml(result) {
        if (!result || !result.warning) return '';
        return `<div class="asi-pfcg-result-note">${escapeHtml(result.warning)}</div>`;
    }

    function asiBuildPfcgTransactionsResultHtml(result, roleName) {
        asiEnsurePfcgResultStyles();
        asiEnsurePfcgListStyles();
        const role = result.role || roleName;
        const transactions = Array.isArray(result.transactions) ? result.transactions : [];
        const count = Number(result.count != null ? result.count : transactions.length);
        const bodyHtml = transactions.length
            ? transactions.map((item) => `
                <tr>
                    <td>${escapeHtml(item.tcode || '')}</td>
                    <td>${escapeHtml(item.description || '-')}</td>
                </tr>
            `).join('')
            : `<tr><td colspan="2" class="asi-pfcg-list-empty">Nenhuma transação encontrada para esta função.</td></tr>`;

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:#16a34a;">✓ Foram encontradas ${count} transações na função.</div>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-field" style="margin-bottom:8px;">
                        <span class="asi-pfcg-result-label">Função</span>
                        <span class="asi-pfcg-result-value asi-pfcg-result-value--nowrap">${escapeHtml(role)}</span>
                    </div>
                    <div class="asi-pfcg-list-wrap">
                        <div class="asi-pfcg-list-scroll">
                            <table class="asi-pfcg-list-table">
                                <thead>
                                    <tr><th>Transação</th><th>Descrição</th></tr>
                                </thead>
                                <tbody>${bodyHtml}</tbody>
                            </table>
                        </div>
                        <div class="asi-pfcg-list-footer">Total: ${count} transações</div>
                    </div>
                    ${asiBuildPfcgRoleCompositeNoteHtml(result)}
                    ${asiBuildPfcgRoleWarningNoteHtml(result)}
                </div>
            </div>
        `;
    }

    function asiBuildPfcgUsersResultHtml(result, roleName) {
        asiEnsurePfcgResultStyles();
        asiEnsurePfcgListStyles();
        const role = result.role || roleName;
        const users = Array.isArray(result.users) ? result.users : [];
        const count = Number(result.count != null ? result.count : users.length);
        const bodyHtml = users.length
            ? users.map((item) => `
                <tr>
                    <td>${escapeHtml(item.username || '')}</td>
                    <td>${escapeHtml(item.valid_from || '-')}</td>
                    <td>${escapeHtml(item.valid_to || '-')}</td>
                    <td>${asiPfcgStatusPillHtml(item.assignment_status)}</td>
                </tr>
            `).join('')
            : `<tr><td colspan="4" class="asi-pfcg-list-empty">Nenhum utilizador atribuído a esta função.</td></tr>`;

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:#16a34a;">✓ Foram encontrados ${count} utilizadores atribuídos à função.</div>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-field" style="margin-bottom:8px;">
                        <span class="asi-pfcg-result-label">Função</span>
                        <span class="asi-pfcg-result-value asi-pfcg-result-value--nowrap">${escapeHtml(role)}</span>
                    </div>
                    <div class="asi-pfcg-list-wrap">
                        <div class="asi-pfcg-list-scroll">
                            <table class="asi-pfcg-list-table">
                                <thead>
                                    <tr><th>Utilizador</th><th>Válido de</th><th>Válido até</th><th>Estado</th></tr>
                                </thead>
                                <tbody>${bodyHtml}</tbody>
                            </table>
                        </div>
                        <div class="asi-pfcg-list-footer">Total: ${count} utilizadores</div>
                    </div>
                    ${asiBuildPfcgRoleCompositeNoteHtml(result)}
                    ${asiBuildPfcgRoleWarningNoteHtml(result)}
                </div>
            </div>
        `;
    }

    function asiBuildPfcgExcelProcessingHtml(fileName = '', message = 'A abrir o seletor de ficheiros Excel...') {
        return `
            <div style="display:flex;align-items:center;gap:10px;">
                <svg width="14" height="14" viewBox="0 0 50 50" aria-hidden="true" style="flex:0 0 auto;animation: spin 1s linear infinite;color:#94a3b8;">
                    <circle cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" style="stroke-dasharray: 1, 150; stroke-dashoffset: 0; animation: dash 1.5s ease-in-out infinite;"></circle>
                </svg>
                <div style="display:flex;flex-direction:column;gap:4px;">
                    <span style="font-weight:600;">${escapeHtml(message)}</span>
                    ${fileName ? `<span style="font-size:0.76rem;color:var(--text-secondary);">Ficheiro: ${escapeHtml(fileName)}</span>` : ''}
                </div>
            </div>
        `;
    }

    function asiBuildPfcgExcelSummaryHtml(summary) {
        if (!summary || typeof summary !== 'object') return '';
        const entries = Object.entries(summary).filter(([, value]) => value !== null && value !== undefined && String(value).trim() !== '');
        if (!entries.length) return '';
        return `
            <div class="asi-pfcg-excel-summary-grid">
                ${entries.map(([key, value]) => asiBuildPfcgExcelSummaryItem(key.replaceAll('_', ' '), value)).join('')}
            </div>
        `;
    }

    function asiBuildPfcgExcelWarningsHtml(items, title) {
        if (!Array.isArray(items) || !items.length) return '';
        return `
            <div style="margin-top:14px;">
                <div style="font-size:0.72rem;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:var(--text-secondary);">${escapeHtml(title)}</div>
                <ul style="margin:8px 0 0 18px;padding:0;color:var(--text-primary);font-size:0.82rem;line-height:1.5;">
                    ${items.map((item) => `<li>${escapeHtml(String(item))}</li>`).join('')}
                </ul>
            </div>
        `;
    }

    function asiBuildPfcgExcelRowsHtml(rows) {
        if (!Array.isArray(rows) || !rows.length) return '';
        const groups = new Map();
        rows.forEach((row) => {
            const agrName = String(row?.AGR_NAME || '').trim() || 'SEM_NOME';
            if (!groups.has(agrName)) {
                groups.set(agrName, {
                    agrName,
                    text: String(row?.TEXT || '').trim(),
                    rows: [],
                    tcodes: []
                });
            }
            const group = groups.get(agrName);
            const tcode = String(row?.TCODE || '').trim();
            const text = String(row?.TEXT || '').trim();
            if (!group.text && text) group.text = text;
            if (tcode) group.tcodes.push(tcode);
            group.rows.push(row);
        });

        const groupCardsHtml = Array.from(groups.values()).map((group) => {
            const uniqueTcodes = Array.from(new Set(group.tcodes.filter(Boolean)));
            const visibleTcodes = uniqueTcodes.slice(0, 6);
            const hiddenCount = Math.max(uniqueTcodes.length - visibleTcodes.length, 0);
            return `
                <div style="height:100%;display:flex;flex-direction:column;padding:12px 12px 10px;border:1px solid rgba(148,163,184,0.18);border-radius:12px;background:linear-gradient(180deg,#ffffff 0%,#fbfdff 100%);box-shadow:0 1px 2px rgba(15,23,42,0.04);box-sizing:border-box;">
                    <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:10px;flex-wrap:wrap;">
                        <div style="min-width:0;flex:1 1 auto;">
                            <div style="font-size:0.68rem;font-weight:800;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-secondary);">Função Simples</div>
                            <div style="margin-top:3px;font-size:0.9rem;font-weight:800;line-height:1.18;color:var(--text-primary);word-break:break-word;">${escapeHtml(group.agrName)}</div>
                            ${group.text ? `<div style="margin-top:4px;font-size:0.76rem;line-height:1.35;color:var(--text-secondary);word-break:break-word;">${escapeHtml(group.text)}</div>` : ''}
                        </div>
                        <div style="text-align:right;min-width:82px;flex:0 0 auto;">
                            <div style="font-size:0.68rem;font-weight:800;letter-spacing:0.06em;text-transform:uppercase;color:var(--text-secondary);">Registos</div>
                            <div style="margin-top:4px;font-size:0.92rem;font-weight:800;color:var(--text-primary);">${group.rows.length}</div>
                        </div>
                    </div>
                    <div style="margin-top:auto;padding-top:10px;display:flex;flex-wrap:wrap;gap:6px;">
                        ${visibleTcodes.map((tcode) => `
                            <span style="display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;background:#eef2ff;border:1px solid rgba(59,130,246,0.18);font-size:0.72rem;font-weight:700;color:#1d4ed8;line-height:1.15;">${escapeHtml(tcode)}</span>
                        `).join('')}
                        ${hiddenCount > 0 ? `<span style="display:inline-flex;align-items:center;padding:4px 8px;border-radius:999px;background:#f8fafc;border:1px dashed rgba(148,163,184,0.55);font-size:0.72rem;font-weight:700;color:var(--text-secondary);line-height:1.15;">+${hiddenCount}</span>` : ''}
                    </div>
                </div>
            `;
        }).join('');

        return `
            <div style="margin-top:14px;padding-top:10px;border-top:1px solid rgba(148,163,184,0.2);">
                <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;">
                    <div style="font-size:0.72rem;font-weight:700;letter-spacing:0.04em;text-transform:uppercase;color:var(--text-secondary);">Linhas preenchidas</div>
                    <div style="font-size:0.78rem;color:var(--text-secondary);font-weight:600;">Total: ${rows.length}</div>
                </div>
                <div class="pfcg-create-grid">
                    ${groupCardsHtml}
                </div>
            </div>
        `;
    }

    function asiBuildPfcgExcelResultHtml(result, fileName) {
        asiEnsurePfcgResultStyles();
        const status = String(result?.status || '').toUpperCase();
        const roleInExcel = result?.role_in_excel || result?.role || '';
        const roleValue = result?.role || (result?.sheet === 'PFCG_CREATE' ? 'PFCG_CREATE' : (roleInExcel || 'N/D'));
        const heading = status === 'VALID'
            ? '✓ Ficheiro analisado com sucesso.'
            : 'Não é possível preparar a criação.';
        const accent = status === 'VALID' ? '#16a34a' : '#dc2626';
        const roleMatches = result?.role && roleInExcel && result.role !== roleInExcel
            ? `<div style="margin-top:10px;font-size:0.82rem;color:#f59e0b;font-weight:600;">O Perfil informado no Excel não corresponde ao Perfil analisado em PRD.</div>`
            : '';
        const description = result?.description
            ? `<div class="asi-pfcg-excel-summary-note"><strong>Descrição:</strong> ${escapeHtml(String(result.description))}</div>`
            : '';
        const sheet = result?.sheet
            ? `<div class="asi-pfcg-excel-summary-note"><strong>Sheet:</strong> ${escapeHtml(String(result.sheet))}</div>`
            : '';
        const summaryHtml = asiBuildPfcgExcelSummaryHtml(result?.summary);
        const warningsHtml = asiBuildPfcgExcelWarningsHtml(result?.warnings || [], 'Avisos');
        const errorsHtml = status === 'VALID'
            ? ''
            : asiBuildPfcgExcelWarningsHtml(result?.errors || [], 'Problemas encontrados');
        const rowsHtml = result?.sheet === 'PFCG_CREATE'
            ? asiBuildPfcgExcelRowsHtml(result?.filled_rows || [])
            : '';
        const finalNote = status === 'VALID'
            ? (result?.sheet === 'PFCG_CREATE'
                ? `<div style="margin-top:14px;font-size:0.84rem;color:var(--text-secondary);font-weight:600;">Os registos da sheet PFCG_CREATE foram carregados para revisão.</div>`
                : `<div style="margin-top:14px;font-size:0.84rem;color:var(--text-secondary);font-weight:600;">O Perfil está pronto para preparação da execução.</div>`)
            : '';

        return `
            <div>
                <div style="display:flex;align-items:center;gap:8px;font-weight:700;color:${accent};">${heading}</div>
                <div style="margin-top:12px;padding-top:10px;border-top:1px solid rgba(148,163,184,0.2);">
                ${description}
                <div class="asi-pfcg-excel-summary-grid">
                    ${asiBuildPfcgExcelSummaryItem('Perfil', roleValue)}
                    ${asiBuildPfcgExcelSummaryItem('Ficheiro', fileName)}
                    ${asiBuildPfcgExcelSummaryItem('Sistema', result?.system)}
                        ${asiBuildPfcgExcelSummaryItem('Cliente', result?.client)}
                    </div>
                    ${sheet}
                </div>
                ${roleMatches}
                ${summaryHtml}
                ${warningsHtml}
                ${errorsHtml}
                ${finalNote}
                ${rowsHtml}
            </div>
        `;
    }

    function asiUpdateMessage(messageId, updates = {}) {
        const messageIndex = asiChatHistory.findIndex((msg) => msg.id === messageId);
        if (messageIndex === -1) return null;
        asiChatHistory[messageIndex] = {
            ...asiChatHistory[messageIndex],
            ...updates
        };
        asiRenderMessages();
        return asiChatHistory[messageIndex];
    }

    function asiScrollToBottom() {
        const { messages } = asiGetElements();
        if (!messages) return;
        if (typeof messages.scrollTo === 'function') {
            messages.scrollTo({
                top: messages.scrollHeight,
                behavior: 'smooth'
            });
            return;
        }
        messages.scrollTop = messages.scrollHeight;
    }

    function asiScrollToBottomAfterPaint() {
        if (typeof window.requestAnimationFrame === 'function') {
            window.requestAnimationFrame(() => {
                asiScrollToBottom();
                window.requestAnimationFrame(() => asiScrollToBottom());
            });
            return;
        }

        setTimeout(() => asiScrollToBottom(), 0);
    }

    function asiAutoResizeInput() {
        const { input } = asiGetElements();
        if (!input) return;
        const maxHeight = 160;
        input.style.height = 'auto';
        const nextHeight = Math.min(input.scrollHeight, maxHeight);
        input.style.height = `${nextHeight}px`;
        input.style.overflowY = input.scrollHeight > maxHeight ? 'auto' : 'hidden';
    }

    function asiUpdateComposerState() {
        const { input, send } = asiGetElements();
        if (!input || !send) return;
        input.placeholder = asiGetComposerPlaceholder();
        input.disabled = Boolean(asiConversationState.isBusy);
        send.disabled = Boolean(asiConversationState.isBusy) || !input.value.trim();
        asiAutoResizeInput();
    }

    function asiGetFiFlowLabel(workflow) {
        return String(workflow || '').trim().toLowerCase() === 'f110_default_document'
            ? 'Execução F110'
            : 'Documento FI';
    }

    function asiGetFiFlowTone(workflow) {
        return String(workflow || '').trim().toLowerCase() === 'f110_default_document'
            ? 'rgba(16,185,129,0.20)'
            : 'rgba(59,130,246,0.18)';
    }

    function asiGetFiContextModel() {
        const selectedWorkflow = String(asiConversationState.selectedFiWorkflow || '').trim().toLowerCase();
        const lastWorkflow = String(asiConversationState.lastFiDocumentWorkflow || '').trim().toLowerCase();
        const workflow = selectedWorkflow || lastWorkflow;
        const flowLabel = asiGetFiFlowLabel(workflow);
        const environment = String(asiConversationState.selectedFiEnvironment || asiConversationState.lastFiDocumentEnvironment || '').trim().toUpperCase();
        const branch = asiNormalizeFiBranch(asiConversationState.selectedFiBranch || asiConversationState.lastFiDocumentBranch || '');
        const branchLabel = branch === 'fornecedor'
            ? 'Fornecedor'
            : branch === 'razao'
                ? 'Razão'
                : branch === 'cliente'
                    ? 'Cliente'
                    : '';
        const documentNumber = String(asiConversationState.lastFiDocumentNumber || '').trim();
        const hasContext = Boolean(workflow || environment || branchLabel || documentNumber);
        return {
            workflow,
            flowLabel,
            environment,
            branchLabel,
            documentNumber,
            hasContext,
            tone: asiGetFiFlowTone(workflow),
        };
    }

    function asiRenderMessages() {
        const { messages } = asiGetElements();
        if (!messages) return;

        messages.innerHTML = asiChatHistory.map((msg) => {
            const isUser = msg.role === 'user';
            const wrapperClass = isUser
                ? 'agent-salsa-message agent-salsa-message-user'
                : `agent-salsa-message agent-salsa-message-assistant${msg.wide ? ' agent-salsa-message--wide' : ''}`;
            const bubbleClass = isUser
                ? 'chat-msg-bubble chat-msg-user'
                : `chat-msg-bubble chat-msg-bot${msg.wide ? ' chat-msg-bubble--pfcg' : ''}${msg.bubbleClassName ? ` ${msg.bubbleClassName}` : ''}`;
            const label = isUser ? 'Utilizador' : 'Assistente';
            const bubbleContent = msg.html
                ? msg.html
                : escapeHtml(msg.text).replace(/\n/g, '<br>');
            const belowBubbleHtml = !isUser && msg.isProcessing && typeof msg.belowBubbleHtml === 'string'
                ? msg.belowBubbleHtml
                : '';
            const actionsHtml = !isUser && Array.isArray(msg.actions) && msg.actions.length > 0
                ? `<div class="agent-salsa-quick-actions-stack">${asiRenderQuickActionButtons(msg.actions, msg.actionLevel || 0, msg.parentActionId || '', msg.selectionGroupKey || '__root__')}</div>`
                : '';

            return `
                <div class="${wrapperClass}">
                    <p class="agent-salsa-message-label">${label}</p>
                    <div class="${bubbleClass}">${bubbleContent}</div>
                    ${belowBubbleHtml}
                    ${actionsHtml}
                </div>
            `;
        }).join('');

        asiScrollToBottomAfterPaint();
    }

    function asiAppendMessage(message) {
        asiChatHistory.push(message);
        asiRenderMessages();
    }

    function asiResetPfcgInteraction(options = {}) {
        asiStopPfcgPolling();
        asiConversationState = {
            ...asiConversationState,
            awaitingInput: options.awaitingInput || '',
            pendingJobId: '',
            pendingRoleName: '',
            pendingMessageId: '',
            pendingExcelSelectionJobId: '',
            pendingExcelAnalyzeJobId: '',
            pendingExcelMessageId: '',
            pendingExcelFileName: '',
            isBusy: false
        };
        asiUpdateComposerState();
    }

    async function asiStartPfcgPolling(jobId, roleName, messageId) {
        const startedAt = Date.now();

        asiStopPfcgPolling();
        asiPfcgPollingTimer = setInterval(async () => {
            if (asiPfcgPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'A análise está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A análise está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                return;
            }

            asiPfcgPollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/analyze/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json'
                    }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a análise do Perfil de Autorização em PRD.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível concluir a análise do Perfil de Autorização em PRD.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const safeDetail = result && typeof result.message === 'string' ? result.message : '';
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a análise do Perfil de Autorização em PRD.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível concluir a análise do Perfil de Autorização em PRD.',
                            safeDetail
                        ),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    return;
                }

                const normalizedResult = {
                    ...result,
                    role: result.role || roleName
                };
                asiLastAnalyzedPfcgRole = normalizedResult.role;
                const isExisting = normalizedResult.status === 'EXISTE';
                asiUpdateMessage(messageId, {
                    text: isExisting
                        ? 'A função já existe em PRD.'
                        : 'A função não existe em PRD.',
                    html: asiBuildPfcgSuccessHtml(normalizedResult),
                    isProcessing: false,
                    wide: true,
                    actions: isExisting ? ASI_PFCG_ROLE_RESULT_ACTIONS : [],
                    actionLevel: 0,
                    parentActionId: '',
                    selectionGroupKey: '__pfcg_role_result__'
                });
                if (isExisting) {
                    asiPfcgRoleState = {
                        role: normalizedResult.role,
                        description: normalizedResult.description,
                        language: normalizedResult.language,
                        system: normalizedResult.system,
                        client: normalizedResult.client
                    };
                } else {
                    asiPfcgRoleState = null;
                }
                asiResetPfcgInteraction();
            } catch (error) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a análise do Perfil de Autorização em PRD.',
                    html: asiBuildPfcgErrorHtml(
                        'Não foi possível concluir a análise do Perfil de Autorização em PRD.'
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
            } finally {
                asiPfcgPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiStartPfcgAnalysis(roleName) {
        const { input } = asiGetElements();
        const pendingExcelSelectionJobId = String(asiConversationState.pendingExcelSelectionJobId || '').trim();
        const pendingExcelFileName = String(asiConversationState.pendingExcelFileName || '').trim();
        const pendingExcelMessageId = String(asiConversationState.pendingExcelMessageId || '').trim();

        if (pendingExcelSelectionJobId && pendingExcelFileName && pendingExcelMessageId) {
            asiConversationState = {
                ...asiConversationState,
                awaitingInput: '',
                pendingRoleName: roleName,
                lastPfcgRoleName: roleName,
                isBusy: true
            };
            asiUpdateComposerState();
            await asiStartPfcgCreateExcelAnalysis(
                roleName,
                pendingExcelSelectionJobId,
                pendingExcelFileName,
                pendingExcelMessageId
            );
            return;
        }

        const processingMessage = asiCreateMessage('assistant', `A analisar ${roleName} no SAP PRD...`, {
            html: asiBuildPfcgProcessingHtml(roleName),
            isProcessing: true
        });

        asiAppendMessage(processingMessage);
        asiConversationState = {
            ...asiConversationState,
            awaitingInput: '',
            pendingRoleName: roleName,
            pendingMessageId: processingMessage.id,
            pendingJobId: '',
            lastPfcgRoleName: roleName,
            isBusy: true
        };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/analyze', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({
                    role_name: roleName
                })
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                const detail = data && typeof data.detail === 'string' ? data.detail : '';
                if (response.status === 400) {
                    asiUpdateMessage(processingMessage.id, {
                        text: ASI_PFCG_INVALID_MESSAGE,
                        html: asiBuildPfcgErrorHtml(
                            'O nome do Perfil de Autorização contém caracteres inválidos.',
                            'Utilize apenas letras, números, "_", "-", "/" ou ":".'
                        ),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction({
                        awaitingInput: ASI_PFCG_AWAITING_INPUT
                    });
                    if (input) input.focus();
                    return;
                }
                throw new Error(detail || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            asiConversationState = {
                ...asiConversationState,
                pendingJobId: jobId,
                pendingRoleName: roleName,
                pendingMessageId: processingMessage.id,
                lastPfcgRoleName: roleName,
                isBusy: true
            };
            asiUpdateComposerState();
            await asiStartPfcgPolling(jobId, roleName, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível concluir a análise do Perfil de Autorização em PRD.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível concluir a análise do Perfil de Autorização em PRD.'
                ),
                isProcessing: false
            });
            asiResetPfcgInteraction();
            if (input) input.focus();
        }
    }

    function asiBuildPfcgTransactionRolesResultHtml(result) {
        asiEnsurePfcgResultStyles();
        asiEnsurePfcgListStyles();
        const tcode = String(result.tcode || '');
        const tcodeDesc = result.tcode_description ? ` — ${escapeHtml(result.tcode_description)}` : '';
        const roles = Array.isArray(result.roles) ? result.roles : [];
        const count = Number(result.count != null ? result.count : roles.length);
        const heading = count > 0
            ? `✓ A transação ${escapeHtml(tcode)} está em ${count} função(ões) Z* em PRD.`
            : `A transação ${escapeHtml(tcode)} não está em nenhuma função Z* em PRD.`;
        const bodyHtml = roles.length
            ? roles.map((item) => {
                const parents = Array.isArray(item.composite_parents) ? item.composite_parents : [];
                const via = parents.length
                    ? `Composta: ${parents.map((p) => escapeHtml(p)).join(', ')}`
                    : 'Direta';
                return `
                    <tr>
                        <td>${escapeHtml(item.role || '')}</td>
                        <td>${escapeHtml(item.description || '-')}</td>
                        <td>${via}</td>
                    </tr>
                `;
            }).join('')
            : `<tr><td colspan="3" class="asi-pfcg-list-empty">Nenhuma função encontrada.</td></tr>`;

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:${count > 0 ? '#16a34a' : '#b45309'};">${heading}</div>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-field" style="margin-bottom:8px;">
                        <span class="asi-pfcg-result-label">Transação</span>
                        <span class="asi-pfcg-result-value asi-pfcg-result-value--nowrap">${escapeHtml(tcode)}${tcodeDesc}</span>
                    </div>
                    <div class="asi-pfcg-list-wrap">
                        <div class="asi-pfcg-list-scroll">
                            <table class="asi-pfcg-list-table">
                                <thead>
                                    <tr><th>Função</th><th>Descrição</th><th>Atribuição</th></tr>
                                </thead>
                                <tbody>${bodyHtml}</tbody>
                            </table>
                        </div>
                        <div class="asi-pfcg-list-footer">Total: ${count} função(ões)</div>
                    </div>
                    ${asiBuildPfcgRoleWarningNoteHtml(result)}
                </div>
            </div>
        `;
    }

    async function asiPollPfcgTransactionRoles(jobId, tcode, messageId) {
        const startedAt = Date.now();
        asiStopPfcgPolling();
        asiPfcgPollingTimer = setInterval(async () => {
            if (asiPfcgPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'A análise da transação está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A análise da transação está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                return;
            }

            asiPfcgPollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/transaction/roles/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));
                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }
                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a análise da transação em PRD.',
                        html: asiBuildPfcgErrorHtml('Não foi possível concluir a análise da transação em PRD.', data.message || ''),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const detail = result && typeof result.message === 'string' ? result.message : '';
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a análise da transação em PRD.',
                        html: asiBuildPfcgErrorHtml('Não foi possível concluir a análise da transação em PRD.', detail),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    return;
                }

                asiUpdateMessage(messageId, {
                    text: `Funções da transação ${tcode}.`,
                    html: asiBuildPfcgTransactionRolesResultHtml(result),
                    isProcessing: false,
                    wide: true,
                    actions: asiPfcgRootMenuActions(),
                    ...ASI_PFCG_ROOT_MENU_META
                });
                asiResetPfcgInteraction();
            } catch (error) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a análise da transação em PRD.',
                    html: asiBuildPfcgErrorHtml('Não foi possível concluir a análise da transação em PRD.'),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
            } finally {
                asiPfcgPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiStartPfcgTransactionRoles(tcode) {
        const { input } = asiGetElements();
        const processingText = `A procurar as funções com a transação ${tcode} em SAP PRD...`;
        const processingMessage = asiCreateMessage('assistant', processingText, {
            html: asiBuildPfcgGenericProcessingHtml(processingText),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, awaitingInput: '', isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/transaction/roles', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                body: JSON.stringify({ tcode })
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }
            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }
            await asiPollPfcgTransactionRoles(jobId, tcode, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível iniciar a análise da transação.',
                html: asiBuildPfcgErrorHtml('Não foi possível iniciar a análise da transação.', error.message || ''),
                isProcessing: false
            });
            asiResetPfcgInteraction();
            if (input) input.focus();
        }
    }

    function asiBuildPfcgObjectRolesResultHtml(result) {
        asiEnsurePfcgResultStyles();
        asiEnsurePfcgListStyles();
        const obj = String(result.auth_object || '');
        const objDesc = result.auth_object_text ? ` — ${escapeHtml(result.auth_object_text)}` : '';
        const roles = Array.isArray(result.roles) ? result.roles : [];
        const count = Number(result.count != null ? result.count : roles.length);
        const heading = count > 0
            ? `✓ O objeto ${escapeHtml(obj)} está em ${count} função(ões) Z* em PRD.`
            : `O objeto ${escapeHtml(obj)} não está em nenhuma função Z* em PRD.`;
        const bodyHtml = roles.length
            ? roles.map((item) => {
                const parents = Array.isArray(item.composite_parents) ? item.composite_parents : [];
                const via = parents.length
                    ? `Composta: ${parents.map((p) => escapeHtml(p)).join(', ')}`
                    : 'Direta';
                return `
                    <tr>
                        <td>${escapeHtml(item.role || '')}</td>
                        <td>${escapeHtml(item.description || '-')}</td>
                        <td>${via}</td>
                    </tr>
                `;
            }).join('')
            : `<tr><td colspan="3" class="asi-pfcg-list-empty">Nenhuma função encontrada.</td></tr>`;

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:${count > 0 ? '#16a34a' : '#b45309'};">${heading}</div>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-result-field" style="margin-bottom:8px;">
                        <span class="asi-pfcg-result-label">Objeto</span>
                        <span class="asi-pfcg-result-value asi-pfcg-result-value--nowrap">${escapeHtml(obj)}${objDesc}</span>
                    </div>
                    <div class="asi-pfcg-list-wrap">
                        <div class="asi-pfcg-list-scroll">
                            <table class="asi-pfcg-list-table">
                                <thead>
                                    <tr><th>Função</th><th>Descrição</th><th>Atribuição</th></tr>
                                </thead>
                                <tbody>${bodyHtml}</tbody>
                            </table>
                        </div>
                        <div class="asi-pfcg-list-footer">Total: ${count} função(ões)</div>
                    </div>
                    ${asiBuildPfcgRoleWarningNoteHtml(result)}
                </div>
            </div>
        `;
    }

    async function asiPollPfcgObjectRoles(jobId, authObject, messageId) {
        const startedAt = Date.now();
        asiStopPfcgPolling();
        asiPfcgPollingTimer = setInterval(async () => {
            if (asiPfcgPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'A análise do objeto de autorização está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A análise do objeto de autorização está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                return;
            }

            asiPfcgPollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/object/roles/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));
                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }
                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a análise do objeto de autorização em PRD.',
                        html: asiBuildPfcgErrorHtml('Não foi possível concluir a análise do objeto de autorização em PRD.', data.message || ''),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const detail = result && typeof result.message === 'string' ? result.message : '';
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a análise do objeto de autorização em PRD.',
                        html: asiBuildPfcgErrorHtml('Não foi possível concluir a análise do objeto de autorização em PRD.', detail),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    return;
                }

                asiUpdateMessage(messageId, {
                    text: `Funções com o objeto ${authObject}.`,
                    html: asiBuildPfcgObjectRolesResultHtml(result),
                    isProcessing: false,
                    wide: true,
                    actions: asiPfcgRootMenuActions(),
                    ...ASI_PFCG_ROOT_MENU_META
                });
                asiResetPfcgInteraction();
            } catch (error) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a análise do objeto de autorização em PRD.',
                    html: asiBuildPfcgErrorHtml('Não foi possível concluir a análise do objeto de autorização em PRD.'),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
            } finally {
                asiPfcgPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiStartPfcgObjectRoles(authObject) {
        const { input } = asiGetElements();
        const processingText = `A procurar as funções com o objeto ${authObject} em SAP PRD...`;
        const processingMessage = asiCreateMessage('assistant', processingText, {
            html: asiBuildPfcgGenericProcessingHtml(processingText),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, awaitingInput: '', isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/object/roles', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
                body: JSON.stringify({ auth_object: authObject })
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }
            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }
            await asiPollPfcgObjectRoles(jobId, authObject, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível iniciar a análise do objeto de autorização.',
                html: asiBuildPfcgErrorHtml('Não foi possível iniciar a análise do objeto de autorização.', error.message || ''),
                isProcessing: false
            });
            asiResetPfcgInteraction();
            if (input) input.focus();
        }
    }

    async function asiStartPfcgCreateExcelSelection() {
        const { input } = asiGetElements();
        const roleName = String(asiConversationState.lastPfcgRoleName || asiConversationState.pendingRoleName || '').trim();

        const processingMessage = asiCreateMessage('assistant', 'A abrir o seletor de ficheiros Excel...', {
            html: asiBuildPfcgExcelProcessingHtml('', 'A abrir o seletor de ficheiros Excel...'),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = {
            ...asiConversationState,
            pendingRoleName: roleName,
            lastPfcgRoleName: roleName,
            pendingExcelMessageId: processingMessage.id,
            pendingExcelSelectionJobId: '',
            pendingExcelAnalyzeJobId: '',
            pendingExcelFileName: '',
            isBusy: true
        };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/create/select-excel', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({})
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            asiConversationState = {
                ...asiConversationState,
                pendingExcelSelectionJobId: jobId,
                pendingExcelMessageId: processingMessage.id,
                isBusy: true
            };
            asiUpdateComposerState();
            await asiStartPfcgCreateExcelSelectionPolling(jobId, roleName, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível iniciar a preparação do Excel.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível iniciar a preparação do Excel.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiResetPfcgInteraction();
            asiConversationState = {
                ...asiConversationState,
                lastPfcgRoleName: roleName
            };
            asiUpdateComposerState();
            if (input) input.focus();
        }
    }

    async function asiStartPfcgCreateExcelSelectionPolling(jobId, roleName, messageId) {
        const { input } = asiGetElements();
        const startedAt = Date.now();

        asiStopPfcgPolling();
        asiPfcgPollingTimer = setInterval(async () => {
            if (asiPfcgPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'A seleção do ficheiro Excel está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A seleção do ficheiro Excel está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                asiConversationState = {
                    ...asiConversationState,
                    lastPfcgRoleName: roleName
                };
                asiUpdateComposerState();
                return;
            }

            asiPfcgPollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/create/select-excel/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json'
                    }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível selecionar o ficheiro Excel.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível selecionar o ficheiro Excel.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    asiConversationState = {
                        ...asiConversationState,
                        lastPfcgRoleName: roleName
                    };
                    asiUpdateComposerState();
                    return;
                }

                const fileName = String(data.file_name || '').trim();
                const selectionId = String(data.selection_id || jobId).trim();
                if (!fileName || !selectionId) {
                    throw new Error('Seleção de Excel concluída sem dados válidos.');
                }

                asiConversationState = {
                    ...asiConversationState,
                    pendingExcelSelectionJobId: selectionId,
                    pendingExcelFileName: fileName,
                    pendingExcelMessageId: messageId,
                    lastPfcgRoleName: roleName,
                    awaitingInput: '',
                    isBusy: true
                };
                asiUpdateMessage(messageId, {
                    text: `Ficheiro ${fileName} selecionado. A ler registos da sheet PFCG_CREATE...`,
                    html: asiBuildPfcgExcelProcessingHtml(fileName, `Ficheiro ${fileName} selecionado. A ler registos da sheet PFCG_CREATE...`),
                    isProcessing: true
                });
                asiUpdateComposerState();
                asiStartPfcgCreateExcelAnalysis(roleName, selectionId, fileName, messageId).catch((error) => {
                    console.error(error);
                });
            } catch (error) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível selecionar o ficheiro Excel.',
                    html: asiBuildPfcgErrorHtml(
                        'Não foi possível selecionar o ficheiro Excel.',
                        error.message || ''
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                asiConversationState = {
                    ...asiConversationState,
                    lastPfcgRoleName: roleName
                };
                asiUpdateComposerState();
            } finally {
                asiPfcgPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiStartPfcgCreateExcelAnalysis(roleName, selectionId, fileName, messageId) {
        try {
            const fallbackRoleName = String(roleName || 'PFCG_CREATE').trim().toUpperCase() || 'PFCG_CREATE';
            const response = await fetch('/api/salsa-it-agent/pfcg/create/analyze', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({
                    role_name: fallbackRoleName,
                    selection_id: selectionId
                })
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            asiConversationState = {
                ...asiConversationState,
                pendingExcelAnalyzeJobId: jobId,
                pendingExcelFileName: fileName,
                pendingExcelMessageId: messageId,
                lastPfcgRoleName: roleName,
                isBusy: true
            };
            asiUpdateComposerState();
            await asiStartPfcgCreateExcelAnalysisPolling(jobId, roleName, fileName, messageId);
        } catch (error) {
            asiUpdateMessage(messageId, {
                text: 'Não foi possível analisar o ficheiro Excel.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível analisar o ficheiro Excel.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiResetPfcgInteraction();
            asiConversationState = {
                ...asiConversationState,
                lastPfcgRoleName: roleName
            };
            asiUpdateComposerState();
        }
    }

    async function asiStartPfcgCreateExcelAnalysisPolling(jobId, roleName, fileName, messageId) {
        const { input } = asiGetElements();
        const startedAt = Date.now();

        asiStopPfcgPolling();
        asiPfcgPollingTimer = setInterval(async () => {
            if (asiPfcgPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'A análise do ficheiro Excel está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A análise do ficheiro Excel está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                asiConversationState = {
                    ...asiConversationState,
                    lastPfcgRoleName: roleName
                };
                asiUpdateComposerState();
                return;
            }

            asiPfcgPollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/create/analyze/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json'
                    }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não é possível preparar a criação.',
                        html: asiBuildPfcgErrorHtml(
                            'Não é possível preparar a criação.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    asiConversationState = {
                        ...asiConversationState,
                        lastPfcgRoleName: roleName
                    };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const safeDetail = result && Array.isArray(result.errors) && result.errors.length
                        ? result.errors.join('\n')
                        : (result && typeof result.message === 'string' ? result.message : '');
                    asiUpdateMessage(messageId, {
                        text: 'Não é possível preparar a criação.',
                        html: asiBuildPfcgErrorHtml(
                            'Não é possível preparar a criação.',
                            safeDetail
                        ),
                        isProcessing: false
                    });
                    asiResetPfcgInteraction();
                    asiConversationState = {
                        ...asiConversationState,
                        lastPfcgRoleName: roleName
                    };
                    asiUpdateComposerState();
                    return;
                }

                const successText = result?.sheet === 'PFCG_CREATE'
                    ? 'Registos da sheet PFCG_CREATE carregados com sucesso.'
                    : 'O Perfil está pronto para preparação da execução.';

                if (String(result.status || '').toUpperCase() === 'VALID') {
                    asiUpdateMessage(messageId, {
                        text: successText,
                        html: asiBuildPfcgExcelResultHtml(result, fileName),
                        isProcessing: false,
                        wide: true
                    });
                } else {
                    asiUpdateMessage(messageId, {
                        text: 'Não é possível preparar a criação.',
                        html: asiBuildPfcgExcelResultHtml(result, fileName),
                        isProcessing: false,
                        wide: true
                    });
                }
                asiResetPfcgInteraction();
                asiConversationState = {
                    ...asiConversationState,
                    lastPfcgRoleName: roleName
                };
                asiUpdateComposerState();
            } catch (error) {
                asiStopPfcgPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não é possível preparar a criação.',
                    html: asiBuildPfcgErrorHtml(
                        'Não é possível preparar a criação.',
                        error.message || ''
                    ),
                    isProcessing: false
                });
                asiResetPfcgInteraction();
                asiConversationState = {
                    ...asiConversationState,
                    lastPfcgRoleName: roleName
                };
                asiUpdateComposerState();
            } finally {
                asiPfcgPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    function asiStopPfcgSubPolling() {
        if (asiPfcgSubPollingTimer) {
            clearInterval(asiPfcgSubPollingTimer);
            asiPfcgSubPollingTimer = null;
        }
        asiPfcgSubPollingInFlight = false;
    }

    function asiPfcgSubAnalysisFailMessage(kind) {
        return kind === 'transactions'
            ? 'Não foi possível concluir a análise de transações da função em PRD.'
            : 'Não foi possível concluir a análise de utilizadores da função em PRD.';
    }

    async function asiPollPfcgSubAnalysis(jobId, roleName, messageId, kind) {
        const startedAt = Date.now();
        const pollUrl = kind === 'transactions'
            ? `/api/salsa-it-agent/pfcg/transactions/analyze/${encodeURIComponent(jobId)}`
            : `/api/salsa-it-agent/pfcg/users/analyze/${encodeURIComponent(jobId)}`;
        const failMessage = asiPfcgSubAnalysisFailMessage(kind);

        asiStopPfcgSubPolling();
        asiPfcgSubPollingTimer = setInterval(async () => {
            if (asiPfcgSubPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_SUB_POLL_TIMEOUT_MS) {
                asiStopPfcgSubPolling();
                asiUpdateMessage(messageId, {
                    text: 'A análise está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A análise está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
                return;
            }

            asiPfcgSubPollingInFlight = true;
            try {
                const response = await fetch(pollUrl, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgSubPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: failMessage,
                        html: asiBuildPfcgErrorHtml(failMessage, data.message || ''),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const safeDetail = result && typeof result.message === 'string' ? result.message : '';
                    asiUpdateMessage(messageId, {
                        text: failMessage,
                        html: asiBuildPfcgErrorHtml(failMessage, safeDetail),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                if (result.status === 'NAO_EXISTE') {
                    asiUpdateMessage(messageId, {
                        text: 'A função não existe em PRD.',
                        html: asiBuildPfcgErrorHtml(
                            'A função não existe em PRD.',
                            'Não é possível consultar transações/utilizadores de uma função inexistente.'
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const count = Number(result.count || 0);
                const html = kind === 'transactions'
                    ? asiBuildPfcgTransactionsResultHtml(result, roleName)
                    : asiBuildPfcgUsersResultHtml(result, roleName);
                const successText = kind === 'transactions'
                    ? `✓ Foram encontradas ${count} transações na função.`
                    : `✓ Foram encontrados ${count} utilizadores atribuídos à função.`;
                const footerActions = kind === 'transactions'
                    ? [ASI_PFCG_ROLE_BACK_TO_CARD_ACTION, ASI_PFCG_ROLE_ANALYZE_USERS_ACTION]
                    : [ASI_PFCG_ROLE_BACK_TO_CARD_ACTION, ASI_PFCG_ROLE_ANALYZE_TRANSACTIONS_ACTION];

                asiUpdateMessage(messageId, {
                    text: successText,
                    html,
                    isProcessing: false,
                    wide: true,
                    actions: footerActions,
                    actionLevel: 0,
                    parentActionId: '',
                    selectionGroupKey: `__pfcg_role_${kind}__`
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } catch (error) {
                asiStopPfcgSubPolling();
                asiUpdateMessage(messageId, {
                    text: failMessage,
                    html: asiBuildPfcgErrorHtml(failMessage),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } finally {
                asiPfcgSubPollingInFlight = false;
            }
        }, ASI_PFCG_SUB_POLL_INTERVAL_MS);
    }

    async function asiStartPfcgRoleSubAnalysis(kind) {
        const roleName = asiPfcgRoleState && asiPfcgRoleState.role
            ? String(asiPfcgRoleState.role).trim()
            : '';
        if (!roleName) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Não foi possível identificar a função analisada em PRD. Volte a analisar a função e tente novamente.'
            ));
            return;
        }

        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        const userText = kind === 'transactions'
            ? `Quero analisar as transações da função ${roleName}.`
            : `Quero analisar os utilizadores atribuídos à função ${roleName}.`;
        const processingText = kind === 'transactions'
            ? 'A consultar as transações atribuídas à função no SAP PRD...'
            : 'A consultar os utilizadores atribuídos à função no SAP PRD...';
        const endpoint = kind === 'transactions'
            ? '/api/salsa-it-agent/pfcg/transactions/analyze'
            : '/api/salsa-it-agent/pfcg/users/analyze';

        asiAppendMessage(asiCreateMessage('user', userText));
        const processingMessage = asiCreateMessage('assistant', processingText, {
            html: asiBuildPfcgGenericProcessingHtml(processingText),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch(endpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({ role_name: roleName })
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                const detail = data && typeof data.detail === 'string' ? data.detail : '';
                throw new Error(detail || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            await asiPollPfcgSubAnalysis(jobId, roleName, processingMessage.id, kind);
        } catch (error) {
            const failMessage = asiPfcgSubAnalysisFailMessage(kind);
            asiUpdateMessage(processingMessage.id, {
                text: failMessage,
                html: asiBuildPfcgErrorHtml(failMessage),
                isProcessing: false
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
        }
    }

    function asiStartPfcgRoleTransactionsAnalysis() {
        return asiStartPfcgRoleSubAnalysis('transactions');
    }

    function asiStartPfcgRoleUsersAnalysis() {
        return asiStartPfcgRoleSubAnalysis('users');
    }

    function asiHandlePfcgRoleBack() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiStopPfcgSubPolling();

        const isComposta = (asiConversationState.subprocesso && asiConversationState.subprocesso.includes('PFCG_COMPOSTA')) || asiConversationState.actionId === 'pfcg-composta-analyze';
        const parentId = isComposta ? 'pfcg-composta' : 'pfcg-create';
        const labelBack = isComposta ? 'Quero voltar às opções da Função Composta.' : 'Quero voltar às opções do Perfil de Autorização.';
        asiAppendMessage(asiCreateMessage('user', labelBack));

        const parentNode = asiFindQuickAction(parentId);
        const children = parentNode && Array.isArray(parentNode.children) ? parentNode.children : [];
        asiAppendMessage(asiCreateMessage('assistant', 'Como deseja prosseguir?', {
            actions: children,
            actionLevel: 2,
            parentActionId: parentId,
            selectionGroupKey: parentId
        }));

        asiConversationState = {
            ...asiConversationState,
            awaitingInput: '',
            pendingJobId: '',
            pendingRoleName: '',
            pendingMessageId: '',
            isBusy: false
        };
        asiUpdateComposerState();
    }

    function asiPresentConfiguracoesMenu() {
        const configAction = asiFindQuickAction('configuracoes', salsaAgentActions);
        const text = (configAction && configAction.followupText) || 'Escolha uma opção de Configurações:';
        const actions = (configAction && Array.isArray(configAction.children)) ? configAction.children : [];
        asiAppendMessage(asiCreateMessage('assistant', text, {
            actions: actions,
            actionLevel: 1,
            parentActionId: 'configuracoes',
            selectionGroupKey: 'configuracoes'
        }));
    }

    function asiHandlePfcgRoleBackToCard() {
        if (!asiPfcgRoleState || !asiPfcgRoleState.role) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Não foi possível recuperar os dados da função analisada. Volte a analisar a função.'
            ));
            return;
        }

        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        asiAppendMessage(asiCreateMessage('assistant', 'A função já existe em PRD.', {
            html: asiBuildPfcgSuccessHtml({ status: 'EXISTE', ...asiPfcgRoleState }),
            wide: true,
            actions: ASI_PFCG_ROLE_RESULT_ACTIONS,
            actionLevel: 0,
            parentActionId: '',
            selectionGroupKey: '__pfcg_role_result__'
        }));
    }

    function asiStopPfcgIndividualPolling() {
        if (asiPfcgIndividualPollingTimer) {
            clearInterval(asiPfcgIndividualPollingTimer);
            asiPfcgIndividualPollingTimer = null;
        }
        asiPfcgIndividualPollingInFlight = false;
    }

    function asiStartPfcgCompostaIndividualCreate() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        asiConversationState = {
            ...asiConversationState,
            pfcgCreateIsComposta: true,
            pfcgCreateRoleName: '',
            pfcgCreateDescription: '',
            pfcgCreateChildRoles: [],
            pfcgCreateTransportMode: '',
            pfcgCreateTransportRequestNumber: '',
            pfcgCreateTransportRequestDescription: '',
            pfcgCreatePreviewJobId: '',
            awaitingInput: ASI_PFCG_COMPOSTA_ROLE_NAME_INPUT,
            isBusy: false
        };
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Envie o Nome da Função Composta que vamos criar'
        ));
        asiUpdateComposerState();
        const { input } = asiGetElements();
        if (input) input.focus();
    }

    function asiStartPfcgIndividualCreate() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        asiConversationState = {
            ...asiConversationState,
            pfcgCreateIsComposta: false,
            pfcgCreateRoleName: '',
            pfcgCreateDescription: '',
            pfcgCreateTcodes: [],
            pfcgCreateTransportMode: '',
            pfcgCreateTransportRequestNumber: '',
            pfcgCreateTransportRequestDescription: '',
            pfcgCreatePreviewJobId: '',
            awaitingInput: ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT,
            isBusy: false
        };
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Envie o Nome da Função Simples que vamos criar'
        ));
        asiUpdateComposerState();
        const { input } = asiGetElements();
        if (input) input.focus();
    }

    async function asiPollPfcgIndividualPreview(jobId, messageId) {
        const startedAt = Date.now();

        asiStopPfcgIndividualPolling();
        asiPfcgIndividualPollingTimer = setInterval(async () => {
            if (asiPfcgIndividualPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgIndividualPolling();
                asiUpdateMessage(messageId, {
                    text: 'A pré-visualização está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A pré-visualização está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false, awaitingInput: '' };
                asiUpdateComposerState();
                return;
            }

            asiPfcgIndividualPollingInFlight = true;
            try {
                const pollPreviewUrl = asiConversationState.pfcgCreateIsComposta
                    ? `/api/salsa-it-agent/pfcg/composta/preview/${encodeURIComponent(jobId)}`
                    : `/api/salsa-it-agent/pfcg/create/rfc/preview/${encodeURIComponent(jobId)}`;
                const response = await fetch(pollPreviewUrl, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgIndividualPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a pré-visualização da criação em DEV.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível concluir a pré-visualização da criação em DEV.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false, awaitingInput: '' };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const safeDetail = result && typeof result.message === 'string' ? result.message : '';
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível preparar a criação em DEV.',
                        html: asiBuildPfcgErrorHtml('Não foi possível preparar a criação em DEV.', safeDetail),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false, awaitingInput: '' };
                    asiUpdateComposerState();
                    return;
                }

                asiUpdateMessage(messageId, {
                    text: 'Confirme os dados antes de criar a função em DEV.',
                    html: asiBuildPfcgIndividualPreviewHtml(result),
                    isProcessing: false,
                    wide: true,
                    actions: [ASI_PFCG_INDIVIDUAL_BACK_ACTION, ASI_PFCG_INDIVIDUAL_CONFIRM_ACTION],
                    actionLevel: 0,
                    parentActionId: '',
                    selectionGroupKey: '__pfcg_individual_preview__'
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } catch (error) {
                asiStopPfcgIndividualPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a pré-visualização da criação em DEV.',
                    html: asiBuildPfcgErrorHtml('Não foi possível concluir a pré-visualização da criação em DEV.'),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false, awaitingInput: '' };
                asiUpdateComposerState();
            } finally {
                asiPfcgIndividualPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    function asiStopPfcgTransportSearchPolling() {
        if (asiPfcgTransportSearchPollingTimer) {
            clearInterval(asiPfcgTransportSearchPollingTimer);
            asiPfcgTransportSearchPollingTimer = null;
        }
        asiPfcgTransportSearchPollingInFlight = false;
    }

    function asiAskPfcgTransportMode() {
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Deseja utilizar ordem de transporte em DEV?',
            {
                actions: [ASI_PFCG_TRANSPORT_LOCAL_ACTION, ASI_PFCG_TRANSPORT_CREATE_ACTION, ASI_PFCG_TRANSPORT_EXISTING_ACTION],
                actionLevel: 0,
                parentActionId: '',
                selectionGroupKey: '__pfcg_transport_mode__'
            }
        ));
        asiConversationState = { ...asiConversationState, isBusy: false };
        asiUpdateComposerState();
    }

    function asiSelectPfcgTransportLocal() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiAppendMessage(asiCreateMessage('user', 'Sem transporte (Local)'));
        asiConversationState = {
            ...asiConversationState,
            pfcgCreateTransportMode: 'LOCAL',
            pfcgCreateTransportRequestNumber: '',
            pfcgCreateTransportRequestDescription: '',
            isBusy: true
        };
        asiUpdateComposerState();
        asiStartPfcgIndividualPreview();
    }

    function asiAskPfcgTransportCreateDescription() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiAppendMessage(asiCreateMessage('user', 'Criar nova Request'));
        asiConversationState = {
            ...asiConversationState,
            awaitingInput: ASI_PFCG_TRANSPORT_CREATE_DESCRIPTION_INPUT,
            isBusy: false
        };
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Envie a descrição/nome para a nova Request de transporte a criar em DEV:'
        ));
        asiUpdateComposerState();
        const { input } = asiGetElements();
        if (input) input.focus();
    }

    function asiBuildPfcgTransportListHtml(result) {
        asiEnsurePfcgResultStyles();
        asiEnsurePfcgListStyles();
        const requests = Array.isArray(result.requests) ? result.requests : [];
        const count = Number(result.requests_count != null ? result.requests_count : requests.length);
        const bodyHtml = requests.length
            ? requests.map((item) => {
                const requestNumber = String(item.request || '');
                const onclickArg = escapeHtml(JSON.stringify(requestNumber));
                return `
                    <tr>
                        <td>${escapeHtml(requestNumber)}</td>
                        <td>${escapeHtml(item.description || '-')}</td>
                        <td>${escapeHtml(item.target_system || '-')}</td>
                        <td><button type="button" class="btn btn-primary" style="padding:3px 10px;font-size:11px;border-radius:6px;" onclick="asiSelectPfcgTransportRequest(${onclickArg})">Selecionar</button></td>
                    </tr>
                `;
            }).join('')
            : `<tr><td colspan="4" class="asi-pfcg-list-empty">Não foram encontradas Requests abertas para este utilizador em DEV.</td></tr>`;

        return `
            <div class="asi-pfcg-result-card">
                <div class="asi-pfcg-result-heading-row">
                    <div class="asi-pfcg-result-heading" style="color:#16a34a;">Requests abertas em DEV (${count})</div>
                </div>
                <div class="asi-pfcg-result-shell">
                    <div class="asi-pfcg-list-wrap">
                        <div class="asi-pfcg-list-scroll">
                            <table class="asi-pfcg-list-table">
                                <thead>
                                    <tr><th>Request</th><th>Descrição</th><th>Sistema</th><th></th></tr>
                                </thead>
                                <tbody>${bodyHtml}</tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        `;
    }

    function asiSelectPfcgTransportRequest(requestNumber) {
        const normalized = String(requestNumber || '').trim();
        if (!normalized) return;
        asiStopPfcgTransportSearchPolling();
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiAppendMessage(asiCreateMessage('user', `Usar a Request ${normalized}`));

        if (asiConversationState.pfcgDeleteRoleName) {
            asiConversationState = {
                ...asiConversationState,
                pfcgDeleteTransportMode: 'EXISTING_REQUEST',
                pfcgDeleteTransportRequestNumber: normalized,
                pfcgDeleteTransportRequestDescription: '',
                isBusy: true
            };
            asiUpdateComposerState();
            asiStartPfcgDeletePreview();
            return;
        }

        asiConversationState = {
            ...asiConversationState,
            pfcgCreateTransportMode: 'EXISTING_REQUEST',
            pfcgCreateTransportRequestNumber: normalized,
            pfcgCreateTransportRequestDescription: '',
            isBusy: true
        };
        asiUpdateComposerState();
        asiStartPfcgIndividualPreview();
    }

    async function asiPollPfcgTransportSearch(jobId, messageId) {
        const startedAt = Date.now();

        asiStopPfcgTransportSearchPolling();
        asiPfcgTransportSearchPollingTimer = setInterval(async () => {
            if (asiPfcgTransportSearchPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgTransportSearchPolling();
                asiUpdateMessage(messageId, {
                    text: 'A procura de Requests está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A procura de Requests está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo e tente novamente.'
                    ),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
                return;
            }

            asiPfcgTransportSearchPollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/transport/search/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgTransportSearchPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a procura de Requests em DEV.',
                        html: asiBuildPfcgErrorHtml('Não foi possível concluir a procura de Requests em DEV.', data.message || ''),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : null;
                if (!result || result.ok !== true) {
                    const safeDetail = result && typeof result.message === 'string' ? result.message : '';
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível obter as Requests de transporte em DEV.',
                        html: asiBuildPfcgErrorHtml('Não foi possível obter as Requests de transporte em DEV.', safeDetail),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                asiUpdateMessage(messageId, {
                    text: 'Selecione a Request de transporte a utilizar.',
                    html: asiBuildPfcgTransportListHtml(result),
                    isProcessing: false,
                    wide: true,
                    actions: [ASI_PFCG_TRANSPORT_BACK_ACTION],
                    actionLevel: 0,
                    parentActionId: '',
                    selectionGroupKey: '__pfcg_transport_list__'
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } catch (error) {
                asiStopPfcgTransportSearchPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a procura de Requests em DEV.',
                    html: asiBuildPfcgErrorHtml('Não foi possível concluir a procura de Requests em DEV.'),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } finally {
                asiPfcgTransportSearchPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiStartPfcgTransportSearch() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiAppendMessage(asiCreateMessage('user', 'Usar Request existente'));

        const processingMessage = asiCreateMessage('assistant', 'A procurar Requests de transporte abertas em DEV...', {
            html: asiBuildPfcgGenericProcessingHtml('A procurar Requests de transporte abertas em DEV...'),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/transport/search', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({})
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            await asiPollPfcgTransportSearch(jobId, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível procurar as Requests de transporte em DEV.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível procurar as Requests de transporte em DEV.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            const { input } = asiGetElements();
            if (input) input.focus();
        }
    }

    async function asiStartPfcgIndividualPreview() {
        const { input } = asiGetElements();
        const roleName = String(asiConversationState.pfcgCreateRoleName || '').trim();
        const description = String(asiConversationState.pfcgCreateDescription || '').trim();
        const tcodes = Array.isArray(asiConversationState.pfcgCreateTcodes) ? asiConversationState.pfcgCreateTcodes : [];

        const processingMessage = asiCreateMessage('assistant', 'A validar os dados em DEV...', {
            html: asiBuildPfcgGenericProcessingHtml('A validar função, descrição e transações em DEV...'),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = {
            ...asiConversationState,
            pfcgCreateMessageId: processingMessage.id,
            isBusy: true
        };
        asiUpdateComposerState();

        try {
            const previewEndpoint = asiConversationState.pfcgCreateIsComposta ? '/api/salsa-it-agent/pfcg/composta/preview' : '/api/salsa-it-agent/pfcg/create/rfc/preview';
            const previewBody = asiConversationState.pfcgCreateIsComposta
                ? {
                    role_name: roleName,
                    description,
                    child_roles: Array.isArray(asiConversationState.pfcgCreateChildRoles) ? asiConversationState.pfcgCreateChildRoles : [],
                    transport_mode: String(asiConversationState.pfcgCreateTransportMode || 'LOCAL').trim().toUpperCase(),
                    request_number: String(asiConversationState.pfcgCreateTransportRequestNumber || '').trim(),
                    request_description: String(asiConversationState.pfcgCreateTransportRequestDescription || '').trim()
                }
                : {
                    role_name: roleName,
                    description,
                    tcodes,
                    transport_mode: String(asiConversationState.pfcgCreateTransportMode || 'LOCAL').trim().toUpperCase(),
                    request_number: String(asiConversationState.pfcgCreateTransportRequestNumber || '').trim(),
                    request_description: String(asiConversationState.pfcgCreateTransportRequestDescription || '').trim()
                };
            const response = await fetch(previewEndpoint, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify(previewBody)
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreatePreviewJobId: jobId,
                isBusy: true
            };
            asiUpdateComposerState();
            await asiPollPfcgIndividualPreview(jobId, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível preparar a pré-visualização da criação em DEV.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível preparar a pré-visualização da criação em DEV.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            if (input) input.focus();
        }
    }

    function asiHandlePfcgIndividualBack() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiStopPfcgIndividualPolling();
        asiAppendMessage(asiCreateMessage('user', 'Quero rever os dados da criação individual.'));
        asiStartPfcgIndividualCreate();
    }

    async function asiPollPfcgIndividualConfirm(jobId, messageId) {
        const startedAt = Date.now();

        asiStopPfcgIndividualPolling();
        asiPfcgIndividualPollingTimer = setInterval(async () => {
            if (asiPfcgIndividualPollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgIndividualPolling();
                asiUpdateMessage(messageId, {
                    text: 'A criação está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A criação está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo antes de repetir a operação.'
                    ),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
                return;
            }

            asiPfcgIndividualPollingInFlight = true;
            try {
                const pollConfirmUrl = asiConversationState.pfcgCreateIsComposta
                    ? `/api/salsa-it-agent/pfcg/composta/confirm/${encodeURIComponent(jobId)}`
                    : `/api/salsa-it-agent/pfcg/create/rfc/confirm/${encodeURIComponent(jobId)}`;
                const response = await fetch(pollConfirmUrl, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgIndividualPolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a criação da função em DEV.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível concluir a criação da função em DEV.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : {};
                const ok = result.ok === true;

                asiUpdateMessage(messageId, {
                    text: ok
                        ? `Função ${result.role} criada com sucesso em DEV.`
                        : 'Não foi possível concluir a criação da função em DEV.',
                    html: asiBuildPfcgIndividualResultHtml(result),
                    isProcessing: false,
                    wide: true
                });
                asiConversationState = { ...asiConversationState, isBusy: false, awaitingInput: '' };
                asiUpdateComposerState();
                setTimeout(() => {
                    asiPresentConfiguracoesMenu();
                }, 400);
            } catch (error) {
                asiStopPfcgIndividualPolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a criação da função em DEV.',
                    html: asiBuildPfcgErrorHtml('Não foi possível concluir a criação da função em DEV.'),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } finally {
                asiPfcgIndividualPollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiHandlePfcgIndividualConfirm() {
        const { input } = asiGetElements();
        const previewJobId = String(asiConversationState.pfcgCreatePreviewJobId || '').trim();
        if (!previewJobId) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Não foi possível localizar a pré-visualização. Repita a preparação da criação individual.'
            ));
            return;
        }

        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        asiAppendMessage(asiCreateMessage('user', 'Confirmar criação'));
        const processingMessage = asiCreateMessage('assistant', 'A criar a função em DEV via RFC...', {
            html: asiBuildPfcgGenericProcessingHtml('A criar a função em DEV via RFC...'),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, isBusy: true };
        asiUpdateComposerState();

        try {
            const confirmUrl = asiConversationState.pfcgCreateIsComposta ? '/api/salsa-it-agent/pfcg/composta/confirm' : '/api/salsa-it-agent/pfcg/create/rfc/confirm';
            const response = await fetch(confirmUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({ preview_job_id: previewJobId })
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            await asiPollPfcgIndividualConfirm(jobId, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível concluir a criação da função em DEV.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível concluir a criação da função em DEV.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            if (input) input.focus();
        }
    }

    let asiPfcgDeletePollingTimer = null;
    let asiPfcgDeletePollingInFlight = false;

    function asiStopPfcgDeletePolling() {
        if (asiPfcgDeletePollingTimer) {
            clearInterval(asiPfcgDeletePollingTimer);
            asiPfcgDeletePollingTimer = null;
        }
        asiPfcgDeletePollingInFlight = false;
    }

    function asiStartPfcgIndividualDelete() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        asiConversationState = {
            ...asiConversationState,
            pfcgDeleteRoleName: '',
            pfcgDeleteTransportMode: '',
            pfcgDeleteTransportRequestNumber: '',
            pfcgDeleteTransportRequestDescription: '',
            pfcgDeletePreviewJobId: '',
            awaitingInput: ASI_PFCG_DELETE_ROLE_NAME_INPUT,
            isBusy: false
        };
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Envie o Nome do perfil que vamos eliminar'
        ));
        asiUpdateComposerState();
        const { input } = asiGetElements();
        if (input) input.focus();
    }

    function asiAskPfcgDeleteTransportMode() {
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Deseja utilizar ordem de transporte em DEV?',
            {
                actions: [ASI_PFCG_DELETE_TRANSPORT_LOCAL_ACTION, ASI_PFCG_DELETE_TRANSPORT_CREATE_ACTION, ASI_PFCG_DELETE_TRANSPORT_EXISTING_ACTION],
                actionLevel: 0,
                parentActionId: '',
                selectionGroupKey: '__pfcg_delete_transport_mode__'
            }
        ));
        asiConversationState = { ...asiConversationState, isBusy: false };
        asiUpdateComposerState();
    }

    function asiSelectPfcgDeleteTransportLocal() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiAppendMessage(asiCreateMessage('user', 'Sem transporte (Local)'));
        asiConversationState = {
            ...asiConversationState,
            pfcgDeleteTransportMode: 'LOCAL',
            pfcgDeleteTransportRequestNumber: '',
            pfcgDeleteTransportRequestDescription: '',
            isBusy: true
        };
        asiUpdateComposerState();
        asiStartPfcgDeletePreview();
    }

    function asiAskPfcgDeleteTransportCreateDescription() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiAppendMessage(asiCreateMessage('user', 'Criar nova Request'));
        asiConversationState = {
            ...asiConversationState,
            awaitingInput: ASI_PFCG_DELETE_TRANSPORT_CREATE_DESCRIPTION_INPUT,
            isBusy: false
        };
        asiAppendMessage(asiCreateMessage(
            'assistant',
            'Envie a descrição/nome para a nova Request de transporte a criar em DEV:'
        ));
        asiUpdateComposerState();
        const { input } = asiGetElements();
        if (input) input.focus();
    }

    async function asiStartPfcgDeletePreview() {
        const { input } = asiGetElements();
        const roleName = String(asiConversationState.pfcgDeleteRoleName || '').trim();

        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        const processingMessage = asiCreateMessage('assistant', 'A validar dados para eliminação da função em DEV...', {
            html: asiBuildPfcgGenericProcessingHtml('A validar dados para eliminação da função em DEV...'),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/delete/rfc/preview', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({
                    role_name: roleName,
                    transport_mode: String(asiConversationState.pfcgDeleteTransportMode || 'LOCAL').trim().toUpperCase(),
                    request_number: String(asiConversationState.pfcgDeleteTransportRequestNumber || '').trim(),
                    request_description: String(asiConversationState.pfcgDeleteTransportRequestDescription || '').trim()
                })
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgDeletePreviewJobId: jobId,
                isBusy: true
            };
            asiUpdateComposerState();
            await asiPollPfcgDeletePreview(jobId, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível preparar a pré-visualização da eliminação em DEV.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível preparar a pré-visualização da eliminação em DEV.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            if (input) input.focus();
        }
    }

    async function asiPollPfcgDeletePreview(jobId, messageId) {
        const startedAt = Date.now();

        asiStopPfcgDeletePolling();
        asiPfcgDeletePollingTimer = setInterval(async () => {
            if (asiPfcgDeletePollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgDeletePolling();
                asiUpdateMessage(messageId, {
                    text: 'A validação está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A validação está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo antes de repetir a operação.'
                    ),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
                return;
            }

            asiPfcgDeletePollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/delete/rfc/preview/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgDeletePolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível validar os dados da função em DEV.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível validar os dados da função em DEV.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : {};
                if (!result.ok && result.status === 'NOT_FOUND') {
                    asiUpdateMessage(messageId, {
                        text: result.message || `A função não existe no ambiente DEV.`,
                        html: asiBuildPfcgErrorHtml(
                            'Função não encontrada em DEV',
                            result.message || 'A função informada não existe no ambiente DEV.'
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    setTimeout(() => {
                        asiPresentConfiguracoesMenu();
                    }, 400);
                    return;
                }

                if (!result.ok) {
                    asiUpdateMessage(messageId, {
                        text: result.message || 'A validação da função em DEV devolveu erro.',
                        html: asiBuildPfcgErrorHtml(
                            'Não é possível eliminar a função em DEV',
                            result.message || 'Verifique as mensagens de erro antes de prosseguir.'
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                asiUpdateMessage(messageId, {
                    text: 'Confirme os dados antes de eliminar a função em DEV.',
                    html: asiBuildPfcgDeletePreviewHtml(result),
                    isProcessing: false,
                    wide: true,
                    actions: [ASI_PFCG_DELETE_BACK_ACTION, ASI_PFCG_DELETE_CONFIRM_ACTION],
                    actionLevel: 0,
                    parentActionId: '',
                    selectionGroupKey: '__pfcg_delete_preview__'
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } catch (error) {
                asiStopPfcgDeletePolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível validar os dados da função em DEV.',
                    html: asiBuildPfcgErrorHtml('Não foi possível validar os dados da função em DEV.'),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } finally {
                asiPfcgDeletePollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    async function asiHandlePfcgDeleteConfirm() {
        const { input } = asiGetElements();
        const previewJobId = String(asiConversationState.pfcgDeletePreviewJobId || '').trim();
        if (!previewJobId) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Não foi possível localizar a pré-visualização. Repita a preparação da eliminação individual.'
            ));
            return;
        }

        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        asiAppendMessage(asiCreateMessage('user', 'Confirmar eliminação'));
        const processingMessage = asiCreateMessage('assistant', 'A eliminar a função em DEV via RFC...', {
            html: asiBuildPfcgGenericProcessingHtml('A eliminar a função em DEV via RFC...'),
            isProcessing: true
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/salsa-it-agent/pfcg/delete/rfc/confirm', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                },
                body: JSON.stringify({ preview_job_id: previewJobId })
            });
            const data = await response.json().catch(() => ({}));

            if (!response.ok) {
                throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
            }

            const jobId = data && typeof data.job_id === 'string' ? data.job_id.trim() : '';
            if (!jobId) {
                throw new Error('Resposta do backend sem job_id.');
            }

            await asiPollPfcgDeleteConfirm(jobId, processingMessage.id);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: 'Não foi possível concluir a eliminação da função em DEV.',
                html: asiBuildPfcgErrorHtml(
                    'Não foi possível concluir a eliminação da função em DEV.',
                    error.message || ''
                ),
                isProcessing: false
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            if (input) input.focus();
        }
    }

    async function asiPollPfcgDeleteConfirm(jobId, messageId) {
        const startedAt = Date.now();

        asiStopPfcgDeletePolling();
        asiPfcgDeletePollingTimer = setInterval(async () => {
            if (asiPfcgDeletePollingInFlight) return;
            if ((Date.now() - startedAt) >= ASI_PFCG_POLL_TIMEOUT_MS) {
                asiStopPfcgDeletePolling();
                asiUpdateMessage(messageId, {
                    text: 'A eliminação está a demorar mais do que o esperado.',
                    html: asiBuildPfcgErrorHtml(
                        'A eliminação está a demorar mais do que o esperado.',
                        'Verifique se o worker Windows está ativo antes de repetir a operação.'
                    ),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
                return;
            }

            asiPfcgDeletePollingInFlight = true;
            try {
                const response = await fetch(`/api/salsa-it-agent/pfcg/delete/rfc/confirm/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const data = await response.json().catch(() => ({}));

                if (!response.ok) {
                    throw new Error((data && data.detail) || `Erro HTTP ${response.status}`);
                }

                if (data.state === 'pending' || data.state === 'running') {
                    return;
                }

                asiStopPfcgDeletePolling();

                if (data.state === 'failed') {
                    asiUpdateMessage(messageId, {
                        text: 'Não foi possível concluir a eliminação da função em DEV.',
                        html: asiBuildPfcgErrorHtml(
                            'Não foi possível concluir a eliminação da função em DEV.',
                            data.message || ''
                        ),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const result = data && typeof data.result === 'object' ? data.result : {};
                const ok = result.ok === true;

                asiUpdateMessage(messageId, {
                    text: ok
                        ? `Função ${result.role} eliminada com sucesso em DEV.`
                        : 'Não foi possível concluir a eliminação da função em DEV.',
                    html: asiBuildPfcgDeleteResultHtml(result),
                    isProcessing: false,
                    wide: true
                });
                asiConversationState = { ...asiConversationState, isBusy: false, awaitingInput: '' };
                asiUpdateComposerState();
                setTimeout(() => {
                    asiPresentConfiguracoesMenu();
                }, 400);
            } catch (error) {
                asiStopPfcgDeletePolling();
                asiUpdateMessage(messageId, {
                    text: 'Não foi possível concluir a eliminação da função em DEV.',
                    html: asiBuildPfcgErrorHtml('Não foi possível concluir a eliminação da função em DEV.'),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } finally {
                asiPfcgDeletePollingInFlight = false;
            }
        }, ASI_PFCG_POLL_INTERVAL_MS);
    }

    function asiHandlePfcgDeleteBack() {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }
        asiStopPfcgDeletePolling();
        asiAppendMessage(asiCreateMessage('user', 'Quero rever os dados da eliminação individual.'));
        asiStartPfcgIndividualDelete();
    }

    function asiHandlePfcgRoleDynamicAction(actionId) {
        if (actionId === 'pfcg-delete-individual-back') {
            asiHandlePfcgDeleteBack();
            return;
        }
        if (actionId === 'pfcg-delete-individual-confirm') {
            asiHandlePfcgDeleteConfirm();
            return;
        }
        if (actionId === 'pfcg-delete-transport-local') {
            asiSelectPfcgDeleteTransportLocal();
            return;
        }
        if (actionId === 'pfcg-delete-transport-create') {
            asiAskPfcgDeleteTransportCreateDescription();
            return;
        }
        if (actionId === 'pfcg-delete-transport-existing') {
            asiStartPfcgTransportSearch();
            return;
        }
        if (actionId === 'pfcg-role-back') {
            asiHandlePfcgRoleBack();
            return;
        }
        if (actionId === 'pfcg-role-back-to-card') {
            asiHandlePfcgRoleBackToCard();
            return;
        }
        if (actionId === 'pfcg-role-analyze-transactions') {
            asiStartPfcgRoleTransactionsAnalysis();
            return;
        }
        if (actionId === 'pfcg-role-analyze-users') {
            asiStartPfcgRoleUsersAnalysis();
            return;
        }
        if (actionId === 'pfcg-create-individual-back') {
            asiHandlePfcgIndividualBack();
            return;
        }
        if (actionId === 'pfcg-create-individual-confirm') {
            asiHandlePfcgIndividualConfirm();
            return;
        }
        if (actionId === 'pfcg-transport-local') {
            asiSelectPfcgTransportLocal();
            return;
        }
        if (actionId === 'pfcg-transport-create') {
            asiAskPfcgTransportCreateDescription();
            return;
        }
        if (actionId === 'pfcg-transport-existing') {
            asiStartPfcgTransportSearch();
            return;
        }
        if (actionId === 'pfcg-transport-back') {
            asiAskPfcgTransportMode();
            return;
        }
    }

    async function asiSendMessage(presetText = null, options = {}) {
        const { input, send } = asiGetElements();
        if (!input || !send) return;

        const rawMessage = typeof presetText === 'string' ? String(presetText).trim() : input.value.trim();
        if (!rawMessage) {
            asiUpdateComposerState();
            return;
        }

        input.value = '';
        const awaitingRoleName = asiConversationState.awaitingInput === ASI_PFCG_AWAITING_INPUT;
        const normalizedRoleName = awaitingRoleName ? asiNormalizePfcgRoleName(rawMessage) : rawMessage;
        const isValidRoleName = awaitingRoleName ? asiIsValidPfcgRoleName(normalizedRoleName) : true;
        const displayMessage = awaitingRoleName && isValidRoleName ? normalizedRoleName : rawMessage;

        asiAppendMessage(asiCreateMessage('user', displayMessage));
        asiUpdateComposerState();
        input.focus();

        if (awaitingRoleName) {
            if (!isValidRoleName) {
                asiAppendMessage(asiCreateMessage('assistant', ASI_PFCG_INVALID_MESSAGE));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_AWAITING_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            await asiStartPfcgAnalysis(normalizedRoleName);
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_TCODE_INPUT) {
            const tcode = rawMessage.toUpperCase().trim();
            if (!ASI_PFCG_TCODE_PATTERN.test(tcode)) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    'Código de transação inválido. Use apenas letras, números, "_", "/", "-", "+", "." ou "$" (máx. 40).'
                ));
                asiConversationState = { ...asiConversationState, awaitingInput: ASI_PFCG_TCODE_INPUT, isBusy: false };
                asiUpdateComposerState();
                input.focus();
                return;
            }
            await asiStartPfcgTransactionRoles(tcode);
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_AUTHOBJ_INPUT) {
            const authObject = rawMessage.toUpperCase().trim();
            if (!ASI_PFCG_AUTHOBJ_PATTERN.test(authObject)) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    'Objeto de autorização inválido. Use apenas letras, números, "_" ou "/" (máx. 40).'
                ));
                asiConversationState = { ...asiConversationState, awaitingInput: ASI_PFCG_AUTHOBJ_INPUT, isBusy: false };
                asiUpdateComposerState();
                input.focus();
                return;
            }
            await asiStartPfcgObjectRoles(authObject);
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_DELETE_ROLE_NAME_INPUT) {
            const normalizedRoleName = asiNormalizePfcgRoleName(rawMessage);
            if (!normalizedRoleName) {
                asiAppendMessage(asiCreateMessage('assistant', 'Envie o Nome do perfil que vamos eliminar'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_DELETE_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (normalizedRoleName.length > 30) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    `O nome do Perfil de Autorização não pode ultrapassar o tamanho máximo de 30 caracteres (tem ${normalizedRoleName.length} caracteres).\nPor favor, corrija o nome e envie novamente.`
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_DELETE_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (!asiIsValidPfcgRoleName(normalizedRoleName)) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    'O nome do Perfil de Autorização contém caracteres inválidos.\nUtilize apenas letras, números, "_", "-", "/" ou ":".\nPor favor, corrija o nome e envie novamente.'
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_DELETE_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgDeleteRoleName: normalizedRoleName,
                awaitingInput: '',
                isBusy: false
            };
            asiUpdateComposerState();
            asiAskPfcgDeleteTransportMode();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_DELETE_TRANSPORT_CREATE_DESCRIPTION_INPUT) {
            const requestDescription = rawMessage.trim();
            if (!requestDescription) {
                asiAppendMessage(asiCreateMessage('assistant', 'Informe uma descrição válida para a nova Request de transporte.'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_DELETE_TRANSPORT_CREATE_DESCRIPTION_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgDeleteTransportMode: 'CREATE_REQUEST',
                pfcgDeleteTransportRequestDescription: requestDescription,
                pfcgDeleteTransportRequestNumber: '',
                awaitingInput: '',
                isBusy: true
            };
            asiUpdateComposerState();
            await asiStartPfcgDeletePreview();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT) {
            const normalizedRoleName = asiNormalizePfcgRoleName(rawMessage);
            if (!normalizedRoleName) {
                asiAppendMessage(asiCreateMessage('assistant', 'Envie o Nome da Função Simples que vamos criar'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (normalizedRoleName.length > 30) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    `O nome do Perfil de Autorização não pode ultrapassar o tamanho máximo de 30 caracteres (tem ${normalizedRoleName.length} caracteres).\nPor favor, corrija o nome e envie novamente.`
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (!asiIsValidPfcgRoleName(normalizedRoleName)) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    'O nome do Perfil de Autorização contém caracteres inválidos.\nUtilize apenas letras, números, "_", "-", "/" ou ":".\nPor favor, corrija o nome e envie novamente.'
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_INDIVIDUAL_ROLE_NAME_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiLastAnalyzedPfcgRole = normalizedRoleName;
            asiConversationState = {
                ...asiConversationState,
                pfcgCreateRoleName: normalizedRoleName,
                pfcgCreateDescription: '',
                awaitingInput: ASI_PFCG_INDIVIDUAL_DESCRIPTION_INPUT,
                isBusy: false
            };
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Qual é a descrição do Perfil de Autorização?'
            ));
            asiUpdateComposerState();
            input.focus();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_INDIVIDUAL_DESCRIPTION_INPUT) {
            const description = rawMessage.trim();
            if (!description) {
                asiAppendMessage(asiCreateMessage('assistant', 'Informe uma descrição válida para o Perfil de Autorização.'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_INDIVIDUAL_DESCRIPTION_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            if (description.length > 80) {
                asiAppendMessage(asiCreateMessage(
                    'assistant',
                    `A descrição não pode ultrapassar o tamanho máximo de 80 caracteres (tem ${description.length} caracteres).\nPor favor, corrija e envie novamente.`
                ));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_INDIVIDUAL_DESCRIPTION_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreateDescription: description,
                awaitingInput: ASI_PFCG_INDIVIDUAL_TCODES_INPUT,
                isBusy: false
            };
            asiAppendMessage(asiCreateMessage(
                'assistant',
                'Quais as transações a atribuir à função? Se for mais de uma transação, separe por vírgula (ex.: FB01, VL03N).'
            ));
            asiUpdateComposerState();
            input.focus();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_INDIVIDUAL_TCODES_INPUT) {
            const tcodes = asiNormalizePfcgTcodes(rawMessage);
            if (tcodes.length === 0) {
                asiAppendMessage(asiCreateMessage('assistant', 'Informe pelo menos uma transação válida. Se for mais de uma transação, separe por vírgula (ex.: FB01, VL03N).'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_INDIVIDUAL_TCODES_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreateTcodes: tcodes,
                awaitingInput: '',
                isBusy: false
            };
            asiUpdateComposerState();
            asiAskPfcgTransportMode();
            return;
        }

        if (asiConversationState.awaitingInput === ASI_PFCG_TRANSPORT_CREATE_DESCRIPTION_INPUT) {
            const requestDescription = rawMessage.trim();
            if (!requestDescription) {
                asiAppendMessage(asiCreateMessage('assistant', 'Informe uma descrição válida para a nova Request de transporte.'));
                asiConversationState = {
                    ...asiConversationState,
                    awaitingInput: ASI_PFCG_TRANSPORT_CREATE_DESCRIPTION_INPUT,
                    isBusy: false
                };
                asiUpdateComposerState();
                input.focus();
                return;
            }

            asiConversationState = {
                ...asiConversationState,
                pfcgCreateTransportMode: 'CREATE_REQUEST',
                pfcgCreateTransportRequestDescription: requestDescription,
                pfcgCreateTransportRequestNumber: '',
                awaitingInput: '',
                isBusy: true
            };
            asiUpdateComposerState();
            await asiStartPfcgIndividualPreview();
            return;
        }

        if (options.skipAssistantReply) return;

        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
        }

        const assistantReplyText = options.assistantText || asiMockReply();
        const assistantReplyActions = Array.isArray(options.assistantActions) ? options.assistantActions : [];
        const assistantActionLevel = Number(options.assistantActionLevel || 0);
        const assistantParentActionId = options.assistantParentActionId || '';
        const assistantSelectionGroupKey = options.assistantSelectionGroupKey || (assistantParentActionId || '__root__');

        asiChatMockTimer = setTimeout(() => {
            asiAppendMessage(
                asiCreateMessage('assistant', assistantReplyText, {
                    actions: assistantReplyActions,
                    actionLevel: assistantActionLevel,
                    parentActionId: assistantParentActionId,
                    selectionGroupKey: assistantSelectionGroupKey
                })
            );
            asiChatMockTimer = null;
        }, 420);
    }

    function asiSendQuickMessage(prompt, options = {}) {
        asiSendMessage(prompt, options);
    }

    function asiIsFiDefaultQuickAction(action) {
        const processo = String(action && action.processo ? action.processo : '').trim().toLowerCase();
        const subprocesso = String(action && action.subprocesso ? action.subprocesso : '').trim().toLowerCase();
        const mode = String(action && action.mode ? action.mode : '').trim().toLowerCase();
        const environment = String(action && action.environment ? action.environment : '').trim().toUpperCase();
        const workflow = String(action && action.workflow ? action.workflow : '').trim().toLowerCase();

        return (
            mode === 'default' &&
            environment === 'QAD' &&
            workflow !== 'f110_default_document' &&
            processo === 'testes unitários' &&
            subprocesso.includes('criar documento fi')
        );
    }

    function asiExtractJobStatusText(job) {
        const rawStatus = String(job && job.status ? job.status : '').trim();
        if (!rawStatus) return '';
        try {
            const parsed = JSON.parse(rawStatus);
            if (parsed && typeof parsed === 'object') {
                return String(parsed.message || parsed.status || rawStatus).trim();
            }
        } catch (_) {
            // Texto simples.
        }
        return rawStatus;
    }

    async function asiPollGenericJob(jobId, messageId, successTitle, failureTitle) {
        const startedAt = Date.now();
        const timeoutMs = 120000;
        const pollIntervalMs = 2000;

        const pollTimer = setInterval(async () => {
            if ((Date.now() - startedAt) >= timeoutMs) {
                clearInterval(pollTimer);
                asiUpdateMessage(messageId, {
                    text: failureTitle,
                    html: escapeHtml(failureTitle),
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
                return;
            }

            try {
                const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`, {
                    method: 'GET',
                    headers: { 'Accept': 'application/json' }
                });
                const job = await response.json().catch(() => ({}));
                if (!response.ok) {
                    throw new Error((job && job.detail) || `Erro HTTP ${response.status}`);
                }

                const state = String(job.state || '').trim().toLowerCase();
                if (state === 'pending' || state === 'running') {
                    return;
                }

                clearInterval(pollTimer);

                if (state === 'failed') {
                    const failDetail = String(job.status || job.log || '').trim();
                    const failText = failDetail
                        ? `${failureTitle}\n${failDetail}`
                        : failureTitle;
                    asiUpdateMessage(messageId, {
                        text: failText,
                        html: failDetail
                            ? `${escapeHtml(failureTitle)}<br><span style="color: var(--danger);">${escapeHtml(failDetail).replace(/\n/g, '<br>')}</span>`
                            : escapeHtml(failureTitle),
                        isProcessing: false
                    });
                    asiConversationState = { ...asiConversationState, isBusy: false };
                    asiUpdateComposerState();
                    return;
                }

                const jobStatusText = asiExtractJobStatusText(job);
                const successText = jobStatusText
                    ? `${successTitle}\n${jobStatusText}`
                    : successTitle;
                asiUpdateMessage(messageId, {
                    text: successText,
                    html: jobStatusText
                        ? `${escapeHtml(successTitle)}<br><span style="color: var(--text-secondary);">${escapeHtml(jobStatusText).replace(/\n/g, '<br>')}</span>`
                        : escapeHtml(successTitle),
                    isProcessing: false,
                    wide: true
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            } catch (error) {
                clearInterval(pollTimer);
                const failText = `${failureTitle}\n${error.message || 'Falha desconhecida.'}`;
                asiUpdateMessage(messageId, {
                    text: failText,
                    html: `${escapeHtml(failureTitle)}<br><span style="color: var(--danger);">${escapeHtml(error.message || 'Falha desconhecida.').replace(/\n/g, '<br>')}</span>`,
                    isProcessing: false
                });
                asiConversationState = { ...asiConversationState, isBusy: false };
                asiUpdateComposerState();
            }
        }, pollIntervalMs);
    }

    // asiStartFiDefaultJob: definicao unica mais abaixo, com assinatura
    // (action, options = {}). A copia anterior, que estava aqui e era codigo
    // morto (sombreada pela definicao seguinte), foi removida na Fase 1.

    function asiGetFiActionContext(action) {
        const actionId = String(action?.id || '').toLowerCase();
        const environment = String(
            action?.environment
            || (actionId.includes('-dev-') ? 'DEV' : '')
            || (actionId.includes('-prd-') ? 'PRD' : '')
            || (actionId.includes('-qad-') ? 'QAD' : 'QAD')
        ).trim().toUpperCase();
        const branch = asiNormalizeFiBranch(
            action?.branch
            || (actionId.includes('fornecedor') ? 'fornecedor' : actionId.includes('razao') ? 'razao' : 'cliente')
        );
        return { environment, branch };
    }

    function asiBuildFiDefaultPayload(action) {
        const { environment, branch } = asiGetFiActionContext(action);
        const defaults = FI_DEFAULTS || {};
        const common = defaults.common || {};
        const branchDefaults = (defaults.branches || {})[branch] || {};
        const payload = {
            environment,
            branch,
            payload: {
                environment,
                branch,
                data_mode: 'default',
                company_code: common.company_code || '',
                posting_date: common.posting_date || '',
                document_date: common.document_date || '',
                currency: common.currency || 'EUR',
                header_text: common.header_text || '',
                reference: '',
                username: '',
                amount: common.amount || '',
                tax_code: common.tax_code || '',
                tax_amount: common.tax_amount || '0',
                tax_rate: common.tax_rate || '',
                tax_gl_account: common.tax_gl_account || '',
                item_text: common.item_text || common.header_text || '',
                customer_account: '',
                revenue_gl_account: '',
                vendor_account: '',
                expense_gl_account: '',
                debit_gl_account: '',
                credit_gl_account: '',
                tax_direction: common.tax_direction || 'credit',
            }
        };

        if (branch === 'cliente') {
            payload.payload.customer_account = branchDefaults.account || '';
            payload.payload.revenue_gl_account = branchDefaults.counterparty || '';
        } else if (branch === 'fornecedor') {
            payload.payload.vendor_account = branchDefaults.account || '';
            payload.payload.expense_gl_account = branchDefaults.counterparty || '';
        } else if (branch === 'razao') {
            payload.payload.debit_gl_account = branchDefaults.debit_gl_account || '';
            payload.payload.credit_gl_account = branchDefaults.credit_gl_account || '';
        }

        return payload;
    }

    function asiGetF110ProposalConfig(action, fiPayload, documentNumber) {
        const { environment, branch } = asiGetFiActionContext(action);
        const defaults = FI_DEFAULTS || {};
        const common = defaults.common || {};
        const branchDefaults = (defaults.branches || {})[branch] || {};
        const payloadRoot = fiPayload && typeof fiPayload.payload === 'object' ? fiPayload.payload : {};
        const companyCode = String(payloadRoot.company_code || common.company_code || '').trim().toUpperCase();
        const postingDate = String(payloadRoot.posting_date || common.posting_date || payloadRoot.document_date || common.document_date || '').trim();
        const paymentMethod = String(
            branchDefaults.payment_method
            || common.payment_method
            || (branch === 'cliente' ? 'Q' : branch === 'fornecedor' ? 'S' : '')
            || ''
        ).trim().toUpperCase();
        const accountNumber = branch === 'cliente'
            ? String(branchDefaults.account || payloadRoot.customer_account || '').trim().toUpperCase()
            : branch === 'fornecedor'
                ? String(branchDefaults.account || payloadRoot.vendor_account || '').trim().toUpperCase()
                : '';

        return {
            environment,
            branch,
            branchLabel: branch === 'fornecedor' ? 'Fornecedor' : branch === 'razao' ? 'Razão' : 'Cliente',
            operationType: branch === 'fornecedor'
                ? 'pagamento'
                : branch === 'cliente'
                    ? 'cobranca'
                    : '',
            payload: {
                environment,
                company_code: companyCode,
                payment_method: paymentMethod,
                account_number: accountNumber,
                posting_date: postingDate,
                next_due_date: '',
                document_number: String(documentNumber || '').trim(),
            }
        };
    }

    function asiBuildF110CompactDocumentHtml(documentNumber) {
        return `
            <div class="asi-f110-chain-summary">
                <span class="asi-f110-chain-summary-label">Documento</span>
                <span class="asi-f110-chain-summary-value">✓ ${escapeHtml(String(documentNumber || ''))}</span>
            </div>
        `;
    }

    function asiStringifyForLog(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'string') return value;
        if (typeof value === 'number' || typeof value === 'boolean') return String(value);
        try {
            return JSON.stringify(value);
        } catch {
            return String(value);
        }
    }

    function asiNormalizeErrorText(error, fallback = 'Falha desconhecida.') {
        if (typeof error === 'string') {
            const text = error.trim();
            return text || fallback;
        }

        if (error && typeof error === 'object') {
            const candidates = [
                error.message,
                error.detail,
                error.error,
                error.reason,
                error.title,
            ];

            for (const candidate of candidates) {
                const text = asiStringifyForLog(candidate).trim();
                if (text && text !== '[object Object]') {
                    return text;
                }
            }

            const objectText = asiStringifyForLog(error).trim();
            if (objectText && objectText !== '[object Object]') {
                return objectText;
            }
        }

        const fallbackText = asiStringifyForLog(error).trim();
        return fallbackText && fallbackText !== '[object Object]' ? fallbackText : fallback;
    }

    function asiLogF110ProposalIssue(environment, branchLabel, missingFields, payload) {
        const normalizedFields = Array.isArray(missingFields) ? missingFields.filter(Boolean) : [];
        const payloadSnapshot = {
            environment: payload?.environment || '',
            company_code: payload?.company_code || '',
            payment_method: payload?.payment_method || '',
            account_number: payload?.account_number || '',
            posting_date: payload?.posting_date || '',
            next_due_date: payload?.next_due_date || '',
            document_number: payload?.document_number || '',
        };

        console.warn('[ASI][F110] Proposta bloqueada por configuração em falta.', {
            environment: environment || '',
            branch: branchLabel || '',
            missingFields: normalizedFields,
            payload: payloadSnapshot,
        });
    }

    function asiBuildF110ProposalCompactHtml(result, documentNumber) {
        const runId = String(result?.run_id || '').trim();
        const runDate = String(result?.run_date || '').trim();
        const displayDate = runDate && /^\d{8}$/.test(runDate)
            ? `${runDate.slice(6, 8)}${runDate.slice(4, 6)}${runDate.slice(0, 4)}`
            : runDate;
        const summary = `Proposta F110 ${displayDate || '-'}${runId ? `/${runId}` : ''}`;

        return `
            <div class="asi-f110-proposal-summary">
                <div class="asi-f110-proposal-title">${escapeHtml(summary)}</div>
            </div>
        `;
    }

    function asiFormatF110ShortLabel(result) {
        const runId = String(result?.run_id || '').trim();
        const runDate = String(result?.run_date || '').trim();
        const displayDate = runDate && /^\d{8}$/.test(runDate)
            ? `${runDate.slice(6, 8)}${runDate.slice(4, 6)}${runDate.slice(0, 4)}`
            : runDate;
        return `Proposta F110 ${displayDate || '-'}${runId ? `/${runId}` : ''}`;
    }

    function asiStartF110DefaultWorkflow(action) {
        return asiStartFiDefaultJob(action, {
            chainF110Proposal: true,
            returnActionId: 'testes-unitarios'
        });
    }

    async function asiStartFiDefaultJob(action, options = {}) {
        if (asiChatMockTimer) {
            clearTimeout(asiChatMockTimer);
            asiChatMockTimer = null;
        }

        const { environment, branch } = asiGetFiActionContext(action);
        const workflow = String(action && action.workflow ? action.workflow : '').trim().toLowerCase();
        const isF110DefaultWorkflow = workflow === 'f110_default_document';
        const chainF110Proposal = Boolean(options.chainF110Proposal && isF110DefaultWorkflow);
        const returnActionId = options.returnActionId || 'testes-unitarios';
        const branchLabel = branch === 'cliente' ? 'Cliente' : branch === 'fornecedor' ? 'Fornecedor' : 'Razão';
        const userPrompt = action?.prompt || 'Usar informações Default.';
        asiAppendMessage(asiCreateMessage('user', userPrompt));

        const processingLabel = isF110DefaultWorkflow
            ? `A criar o documento FI de apoio à Execução F110 em ${environment} com os dados Default...`
            : `A criar o documento FI padrão de ${branchLabel} em ${environment} via RFC com os dados Default...`;
        const fiPayload = asiBuildFiDefaultPayload(action);
        const processingMessage = asiCreateMessage('assistant', processingLabel, {
            belowBubbleHtml: asiBuildThinkingIndicatorHtml(),
            isProcessing: true,
        });
        asiAppendMessage(processingMessage);
        asiConversationState = { ...asiConversationState, isBusy: true };
        asiUpdateComposerState();

        try {
            const response = await fetch('/api/fi/default-document', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(fiPayload),
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.ok === false || data.status === 'ERRO') {
                throw new Error(data.message || data.detail || `Falha HTTP ${response.status}`);
            }

            const documentNumber = String(data.document_number || '').trim();
            asiConversationState = {
                ...asiConversationState,
                isBusy: false,
                selectedFiEnvironment: environment,
                selectedFiBranch: branch,
                selectedFiWorkflow: workflow,
                lastFiDocumentNumber: documentNumber,
                lastFiDocumentEnvironment: environment,
                lastFiDocumentBranch: branch,
                lastFiDocumentWorkflow: isF110DefaultWorkflow ? 'f110_default_document' : 'fi_default',
            };
            if (chainF110Proposal) {
                asiUpdateMessage(processingMessage.id, {
                    text: documentNumber
                        ? `Documento ${documentNumber}`
                        : 'Documento FI criado.',
                    html: asiBuildF110CompactDocumentHtml(documentNumber),
                    belowBubbleHtml: '',
                    isProcessing: false,
                    wide: false,
                    bubbleClassName: 'asi-f110-chain-document-bubble',
                });
                asiUpdateComposerState();
                await asiStartF110ProposalWorkflow(action, fiPayload, documentNumber, returnActionId);
                return;
            }
            const successFields = [
                {
                    label: 'Fluxo',
                    value: isF110DefaultWorkflow ? 'Execução F110' : 'Documento FI',
                },
                {
                    label: 'Ambiente',
                    value: environment,
                },
                {
                    label: 'Tipo',
                    value: branchLabel,
                },
                {
                    label: 'Estado',
                    value: String(data.status || 'SUCESSO'),
                },
                documentNumber ? {
                    label: 'Documento',
                    value: documentNumber,
                } : null,
                isF110DefaultWorkflow ? {
                    label: 'Uso',
                    value: 'Guardado para a próxima Execução F110.',
                } : null,
            ].filter(Boolean).map((field) => `
                <div class="asi-fi-success-field">
                    <div class="asi-fi-success-label${field.label === 'Estado' ? ' asi-fi-success-label--status' : ''}">${escapeHtml(String(field.label))}</div>
                    <div class="asi-fi-success-value${field.label === 'Estado' ? ' asi-fi-success-value--status' : ''}">${escapeHtml(String(field.value))}</div>
                </div>
            `).join('');

            const successHtml = `
                <div class="asi-fi-success-title">${escapeHtml(isF110DefaultWorkflow ? 'Documento FI de apoio à Execução F110' : 'Documento FI padrão')} executado via RFC</div>
                <div class="asi-fi-success-grid">
                    ${successFields}
                </div>`;

            asiUpdateMessage(processingMessage.id, {
                text: isF110DefaultWorkflow
                    ? `Documento FI de apoio à Execução F110 guardado em ${environment} (${branchLabel}).${documentNumber ? ` Documento ${documentNumber}.` : ''}`
                    : `Documento FI padrão concluído em ${environment} (${branchLabel}).${documentNumber ? ` Documento ${documentNumber}.` : ''}`,
                html: successHtml,
                belowBubbleHtml: '',
                isProcessing: false,
                wide: true,
                bubbleClassName: 'asi-fi-success-bubble',
            });
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
            return;
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: `Não foi possível executar o documento FI em ${environment}.`,
                html: `${escapeHtml(`Não foi possível executar o documento FI em ${environment}.`)}<br><span style="color: var(--danger);">${escapeHtml(error.message || 'Falha desconhecida.').replace(/\n/g, '<br>')}</span>`,
                belowBubbleHtml: '',
                isProcessing: false,
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
        }
    }

    async function asiStartF110ProposalWorkflow(action, fiPayload, documentNumber, returnActionId = 'testes-unitarios') {
        const config = asiGetF110ProposalConfig(action, fiPayload, documentNumber);
        const { environment, branch, branchLabel, operationType, payload } = config;
        const normalizedDocumentNumber = String(documentNumber || '').trim();

        if (!normalizedDocumentNumber) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                `Não foi possível iniciar a proposta F110 em ${environment}.`,
                {
                    html: `
                        <div style="font-weight:700;color:var(--danger);">Não foi possível iniciar a proposta F110 em ${environment}.</div>
                        <div style="margin-top:6px;color:var(--text-secondary);">O documento FI não devolveu um número válido.</div>
                    `
                }
            ));
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
            return;
        }

        if (branch === 'razao' || !operationType) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                `A proposta F110 automática não está disponível para ${branchLabel}.`,
                {
                    html: `
                        <div style="font-weight:700;color:var(--danger);">A proposta F110 automática não está disponível para ${branchLabel}.</div>
                        <div style="margin-top:6px;color:var(--text-secondary);">Este fluxo requer uma especificação funcional adicional para conta de Razão.</div>
                    `
                }
            ));
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
            return;
        }

        const missingFields = [];
        if (!payload.company_code) missingFields.push('empresa');
        if (!payload.payment_method) missingFields.push('forma de pagamento/cobrança');
        if (!payload.account_number) missingFields.push('conta');
        if (!payload.posting_date) missingFields.push('data de lançamento');

        if (missingFields.length > 0) {
            asiLogF110ProposalIssue(environment, branchLabel, missingFields, payload);
            asiAppendMessage(asiCreateMessage(
                'assistant',
                `Não foi possível iniciar a proposta F110 em ${environment}.`,
                {
                    html: `
                        <div style="font-weight:700;color:var(--danger);">Não foi possível iniciar a proposta F110 em ${environment}.</div>
                        <div style="margin-top:6px;color:var(--text-secondary);">Configuração em falta: ${escapeHtml(missingFields.join(', '))}.</div>
                    `
                }
            ));
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
            return;
        }

        const processingMessage = asiCreateMessage('assistant', `A criar a proposta F110 em ${environment} para o documento ${documentNumber}...`, {
            belowBubbleHtml: asiBuildThinkingIndicatorHtml(),
            isProcessing: true,
        });
        asiAppendMessage(processingMessage);

        try {
            const response = await fetch('/api/f110/proposal', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    environment,
                    operation_type: operationType,
                    company_code: payload.company_code,
                    payment_method: payload.payment_method,
                    account_number: payload.account_number,
                    posting_date: payload.posting_date,
                    next_due_date: payload.next_due_date || '',
                    document_number: normalizedDocumentNumber,
                }),
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.ok === false || data.status === 'ERRO') {
                throw new Error(asiNormalizeErrorText(data.message || data.detail || `Falha HTTP ${response.status}`));
            }

            asiConversationState = {
                ...asiConversationState,
                lastF110ProposalPayload: payload,
                lastF110ProposalResult: data,
                lastFiDocumentNumber: normalizedDocumentNumber,
                lastFiDocumentEnvironment: environment,
                lastFiDocumentBranch: branch,
                lastFiDocumentWorkflow: 'f110_default_document',
            };
            asiUpdateMessage(processingMessage.id, {
                text: asiFormatF110ShortLabel(data),
                html: asiBuildF110ProposalCompactHtml(data, normalizedDocumentNumber),
                belowBubbleHtml: '',
                isProcessing: false,
                wide: true,
                bubbleClassName: 'asi-f110-proposal-bubble',
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: `Erro ao executar a proposta F110 para ${normalizedDocumentNumber}.`,
                html: `
                    <div class="asi-f110-proposal-summary">
                        <div class="asi-f110-proposal-title" style="color:var(--danger);">Erro ao executar a proposta F110</div>
                        <div class="asi-f110-proposal-message">${escapeHtml(asiNormalizeErrorText(error, 'Falha desconhecida.')).replace(/\n/g, '<br>')}</div>
                    </div>
                `,
                belowBubbleHtml: '',
                isProcessing: false,
                wide: true,
                bubbleClassName: 'asi-f110-proposal-bubble',
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
        }
    }

    async function asiStartF110PaymentWorkflowFromLastProposal(returnActionId = 'testes-unitarios') {
        const proposalPayload = asiConversationState.lastF110ProposalPayload || null;
        const proposalResult = asiConversationState.lastF110ProposalResult || null;
        const documentNumber = String(
            asiConversationState.lastFiDocumentNumber
            || proposalResult?.document_number
            || proposalPayload?.document_number
            || ''
        ).trim();
        const environment = String(
            asiConversationState.lastFiDocumentEnvironment
            || proposalPayload?.environment
            || 'QAD'
        ).trim().toUpperCase();
        const branch = String(asiConversationState.lastFiDocumentBranch || 'fornecedor').trim().toLowerCase();

        if (!proposalPayload || !documentNumber) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                `N�o foi poss�vel iniciar o pagamento F110 em ${environment}.`,
                {
                    html: `
                        <div style="font-weight:700;color:var(--danger);">N�o foi poss�vel iniciar o pagamento F110 em ${environment}.</div>
                        <div style="margin-top:6px;color:var(--text-secondary);">É necess�rio executar primeiro a proposta e manter o mesmo contexto.</div>
                    `
                }
            ));
            return;
        }

        const paymentPayload = {
            environment,
            operation_type: 'pagamento',
            company_code: String(proposalPayload.company_code || '').trim().toUpperCase(),
            payment_method: String(proposalPayload.payment_method || '').trim().toUpperCase(),
            account_number: String(proposalPayload.account_number || '').trim().toUpperCase(),
            posting_date: String(proposalPayload.posting_date || '').trim(),
            next_due_date: String(proposalPayload.next_due_date || '').trim(),
            document_number: documentNumber,
            source_payload: proposalPayload,
        };

        if (!paymentPayload.company_code || !paymentPayload.payment_method || !paymentPayload.account_number || !paymentPayload.posting_date) {
            asiAppendMessage(asiCreateMessage(
                'assistant',
                `N�o foi poss�vel iniciar o pagamento F110 em ${environment}.`,
                {
                    html: `
                        <div style="font-weight:700;color:var(--danger);">N�o foi poss�vel iniciar o pagamento F110 em ${environment}.</div>
                        <div style="margin-top:6px;color:var(--text-secondary);">O contexto da proposta n�o contém empresa, conta, forma de pagamento ou data de lan�amento.</div>
                    `
                }
            ));
            return;
        }

        const processingMessage = asiCreateMessage('assistant', `A executar o ciclo F110 em ${environment} com os dados da proposta ${documentNumber}...`, {
            belowBubbleHtml: asiBuildThinkingIndicatorHtml(),
            isProcessing: true,
        });
        asiAppendMessage(processingMessage);

        try {
            const response = await fetch('/api/f110/payment', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(paymentPayload),
            });
            const data = await response.json().catch(() => ({}));
            if (!response.ok || data.ok === false || data.status === 'ERRO') {
                throw new Error(asiNormalizeErrorText(data.message || data.detail || `Falha HTTP ${response.status}`));
            }

            asiConversationState = {
                ...asiConversationState,
                lastF110ProposalResult: data,
                lastFiDocumentNumber: documentNumber,
                lastFiDocumentEnvironment: environment,
                lastFiDocumentBranch: branch,
                lastFiDocumentWorkflow: 'f110_default_document',
            };
            asiUpdateMessage(processingMessage.id, {
                text: `Ciclo F110 executado para ${documentNumber}.`,
                html: asiBuildF110ProposalCompactHtml(data, documentNumber),
                belowBubbleHtml: '',
                isProcessing: false,
                wide: true,
                bubbleClassName: 'asi-f110-proposal-bubble',
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
            asiReturnToQuickActionMenu(returnActionId);
        } catch (error) {
            asiUpdateMessage(processingMessage.id, {
                text: `Erro ao executar o ciclo F110 para ${documentNumber}.`,
                html: `
                    <div class="asi-f110-proposal-summary">
                        <div class="asi-f110-proposal-title" style="color:var(--danger);">Erro ao executar o ciclo F110</div>
                        <div class="asi-f110-proposal-message">${escapeHtml(asiNormalizeErrorText(error, 'Falha desconhecida.')).replace(/\n/g, '<br>')}</div>
                    </div>
                `,
                belowBubbleHtml: '',
                isProcessing: false,
                wide: true,
                bubbleClassName: 'asi-f110-proposal-bubble',
            });
            asiConversationState = { ...asiConversationState, isBusy: false };
            asiUpdateComposerState();
        }
    }

    function asiHandleQuickActionSelection(actionId, level, parentActionId = '', selectionGroupKey = '__root__') {
        if (ASI_PFCG_DYNAMIC_ACTION_IDS.has(actionId)) {
            asiHandlePfcgRoleDynamicAction(actionId);
            return;
        }

        if (actionId === ASI_MAIN_MENU_ACTION.id) {
            asiPresentMainMenu();
            return;
        }

        const numericLevel = Number(level);
        const action = asiFindQuickAction(actionId, salsaAgentActions);
        if (!action) return;

        const actionPath = asiFindActionPath(actionId, salsaAgentActions) || [];
        const nextSelectedActions = { '__root__': null };

        if (actionPath.length > 0) {
            nextSelectedActions['__root__'] = actionPath[0].id;
            for (let i = 1; i < actionPath.length; i += 1) {
                nextSelectedActions[actionPath[i - 1].id] = actionPath[i].id;
            }
        }

        if (!actionPath.length) {
            nextSelectedActions[selectionGroupKey || '__root__'] = action.id;
        }

        asiSelectedActions = nextSelectedActions;

        const nextConversationState = asiDefaultConversationState();
        if (action.processo || action.subprocesso || action.mode || action.environment || action.branch || action.workflow) {
            Object.assign(nextConversationState, {
                processo: action.processo || '',
                subprocesso: action.subprocesso || '',
                actionId: action.id,
                mode: action.mode || '',
                selectedFiEnvironment: action.environment || '',
                selectedFiBranch: action.branch || '',
                selectedFiWorkflow: action.workflow || ''
            });
        }
        asiConversationState = nextConversationState;

        asiRenderMessages();

        if (String(action.workflow || '').trim().toLowerCase() === 'f110_default_document') {
            asiStartF110DefaultWorkflow(action);
            return;
        }

        if (asiIsFiDefaultQuickAction(action)) {
            asiStartFiDefaultJob(action);
            return;
        }

        if (action.id === 'pfcg-role-analyze-transacao') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            asiAppendMessage(asiCreateMessage('assistant', 'Qual é o código de transação? (ex.: FB01)'));
            asiConversationState = {
                ...asiConversationState,
                awaitingInput: ASI_PFCG_TCODE_INPUT
            };
            asiUpdateComposerState();
            const { input } = asiGetElements();
            if (input) input.focus();
            return;
        }

        if (action.id === 'pfcg-role-analyze-objeto') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            asiAppendMessage(asiCreateMessage('assistant', 'Qual é o objeto de autorização? (ex.: S_TCODE)'));
            asiConversationState = {
                ...asiConversationState,
                awaitingInput: ASI_PFCG_AUTHOBJ_INPUT
            };
            asiUpdateComposerState();
            const { input } = asiGetElements();
            if (input) input.focus();
            return;
        }

        if (action.id === 'pfcg-composta-analyze' || action.id === 'pfcg-role-analyze-funcao') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            const assistantPrompt = action.followupText || (action.id === 'pfcg-composta-analyze' ? 'Qual é o nome da Função Composta que deseja analisar em PRD?' : 'Qual é o nome do Perfil de Autorização que deseja analisar em PRD?');
            asiAppendMessage(asiCreateMessage('assistant', assistantPrompt));
            asiConversationState = {
                ...asiConversationState,
                awaitingInput: ASI_PFCG_AWAITING_INPUT
            };
            asiUpdateComposerState();
            const { input } = asiGetElements();
            if (input) input.focus();
            return;
        }

        if (action.id === 'pfcg-create-select-excel' || action.id === 'pfcg-composta-select-excel') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            asiStartPfcgCreateExcelSelection();
            return;
        }

        if (action.id === 'pfcg-create-individual') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            asiStartPfcgIndividualCreate();
            return;
        }

        if (action.id === 'pfcg-composta-individual') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            asiStartPfcgCompostaIndividualCreate();
            return;
        }

        if (action.id === 'pfcg-delete-individual') {
            if (asiChatMockTimer) {
                clearTimeout(asiChatMockTimer);
                asiChatMockTimer = null;
            }

            asiAppendMessage(asiCreateMessage('user', action.prompt));
            asiStartPfcgIndividualDelete();
            return;
        }

        const sendOptions = {};
        if (action.followupText) {
            sendOptions.assistantText = action.followupText;
        }
        if (action.followupActionsSource === 'children' && Array.isArray(action.children) && action.children.length > 0) {
            sendOptions.assistantActions = action.children;
            sendOptions.assistantActionLevel = numericLevel + 1;
            sendOptions.assistantParentActionId = action.id;
            sendOptions.assistantSelectionGroupKey = action.id;
        }

        asiSendQuickMessage(action.prompt, sendOptions);
    }

    function asiBindChat() {
        const { input, send, messages } = asiGetElements();
        if (!input || !send || !messages) return;

        if (!input.dataset.bound) {
            input.addEventListener('input', asiUpdateComposerState);
            input.addEventListener('keydown', function(event) {
                if (event.key === 'Enter' && !event.shiftKey) {
                    event.preventDefault();
                    asiSendMessage();
                }
            });
            input.dataset.bound = 'true';
        }

        if (!send.dataset.bound) {
            send.addEventListener('click', asiSendMessage);
            send.dataset.bound = 'true';
        }

        if (!messages.dataset.bound) {
            messages.addEventListener('click', function(event) {
                const quickAction = event.target.closest('[data-agent-action-id]');
                if (!quickAction) return;
                const actionId = quickAction.getAttribute('data-agent-action-id') || '';
                const level = quickAction.getAttribute('data-agent-action-level') || '0';
                const parentActionId = quickAction.getAttribute('data-agent-parent-action-id') || '';
                const selectionGroupKey = quickAction.getAttribute('data-agent-selection-group-key') || '__root__';
                asiHandleQuickActionSelection(actionId, level, parentActionId, selectionGroupKey);
            });
            messages.dataset.bound = 'true';
        }
    }

    function asiInitChat() {
        const { input } = asiGetElements();
        asiBindChat();

        if (!asiChatInitialized) {
            asiStopPfcgPolling();
            asiConversationState = asiDefaultConversationState();
            asiChatHistory = [
                asiCreateMessage('assistant', asiDefaultGreeting(), {
                    actions: salsaAgentActions,
                    actionLevel: 0,
                    selectionGroupKey: '__root__'
                })
            ];
            asiChatInitialized = true;
        }

        asiRenderMessages();
        asiUpdateComposerState();

        if (input) {
            input.focus();
        }
    }

    function renderKpiModalTable(filtered, tbody) {
        tbody.innerHTML = '';

        if (filtered.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" style="padding: 32px 16px; text-align: center; color: var(--text-secondary);">
                        Nenhum job encontrado para esta categoria.
                    </td>
                </tr>
            `;
        } else {
            filtered.forEach(job => {
                const tr = document.createElement('tr');
                tr.className = 'kpi-table-row';
                
                const shortId = job.id.substring(0, 8);
                const dateStr = new Date(job.created_at).toLocaleString('pt-PT', { 
                    day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit'
                });
                
                const proc = job.params ? (job.params.subprocesso || job.params.processo || job.task) : job.task;
                const env = job.ambiente || job.params?.ambiente || 'DEV';
                
                let badgeClass = '';
                let badgeText = job.state;
                if (job.state === 'running') { badgeClass = 'running'; badgeText = 'executando'; }
                else if (job.state === 'pending') { badgeClass = 'pending'; badgeText = 'pendente'; }
                else if (job.state === 'succeeded') { badgeClass = 'success'; badgeText = 'sucesso'; }
                else if (job.state === 'failed') { badgeClass = 'failed'; badgeText = 'erro'; }
                
                let archivedBadge = '';
                if (job.is_archived) {
                    archivedBadge = ` <span style="font-size: 10px; opacity: 0.7; color: var(--text-secondary); font-style: italic;">(arquivado)</span>`;
                }

                tr.innerHTML = `
                    <td style="padding: 12px 16px; font-family: monospace; font-weight: bold; color: var(--text-secondary);">#${shortId}</td>
                    <td style="padding: 12px 16px; white-space: nowrap;">${dateStr}</td>
                    <td style="padding: 12px 16px; font-weight: 600; color: var(--text-primary); max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(proc)}">${escapeHtml(proc)}</td>
                    <td style="padding: 12px 16px;">
                        <span style="font-size: 11px; font-weight: bold; padding: 2px 6px; background: rgba(0,0,0,0.03); border: 1px solid var(--border-color); border-radius: 4px;">${escapeHtml(env)}</span>
                    </td>
                    <td style="padding: 12px 16px;">
                        <span class="badge ${badgeClass}">${badgeText}</span>${archivedBadge}
                    </td>
                    <td style="padding: 12px 16px; text-align: center;">
                        <button class="btn btn-primary" style="padding: 4px 10px; font-size: 11px; border-radius: 6px;" onclick="selectAndFocusJob('${job.id}')">🔍 Focar</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }
    }

    function openKpiJobsModal(type) {
        const modalKpi = document.getElementById('modal-kpi-jobs');
        const titleEl = document.getElementById('kpi-jobs-modal-title');
        const subtitleEl = document.getElementById('kpi-jobs-modal-subtitle');
        const tbody = document.getElementById('kpi-jobs-modal-table-body');
        
        let filtered = [];
        let title = '';
        let subtitle = '';
        const now = new Date();

        if (type === 'running') {
            filtered = allJobs.filter(j => j.state === 'running');
            title = '⚡ Jobs em Execução';
            subtitle = 'Lista de todos os scripts que estão a ser executados no SAP neste momento.';
            renderKpiModalTable(filtered, tbody);
        } else if (type === 'pending') {
            filtered = allJobs.filter(j => j.state === 'pending');
            title = '⏳ Jobs Pendentes na Fila';
            subtitle = 'Lista de pedidos agendados a aguardar a libertação de um worker Windows.';
            renderKpiModalTable(filtered, tbody);
        } else if (type === 'success') {
            filtered = allJobs.filter(j => j.state === 'succeeded' && new Date(j.created_at).toDateString() === now.toDateString());
            title = '🟢 Jobs Concluídos Hoje';
            subtitle = 'Lista de rotinas terminadas com sucesso total no dia de hoje.';
            renderKpiModalTable(filtered, tbody);
        } else if (type === 'failed') {
            filtered = allJobs.filter(j => j.state === 'failed' && new Date(j.created_at).toDateString() === now.toDateString());
            title = '🔴 Jobs com Erro Hoje';
            subtitle = 'Lista de pedidos que falharam ou foram cancelados hoje. Requerem verificação do log.';
            renderKpiModalTable(filtered, tbody);
        } else if (type === 'archived') {
            title = '📦 Histórico de Jobs Arquivados';
            subtitle = 'Lista de jobs arquivados que foram ocultados do painel principal.';
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" style="padding: 32px 16px; text-align: center; color: var(--text-secondary);">
                        <span style="display:inline-block; animation: spin 0.7s linear infinite; margin-right: 6px;">↻</span> A carregar histórico arquivado...
                    </td>
                </tr>
            `;
            fetch('/api/jobs?include_archived=true&limit=100')
                .then(res => res.json())
                .then(data => {
                    const archived = (data.jobs || []).filter(j => j.is_archived);
                    renderKpiModalTable(archived, tbody);
                })
                .catch(err => {
                    tbody.innerHTML = `
                        <tr>
                            <td colspan="6" style="padding: 32px 16px; text-align: center; color: var(--danger);">
                                Erro ao carregar histórico: ${err.message}
                            </td>
                        </tr>
                    `;
                });
        }

        titleEl.textContent = title;
        subtitleEl.textContent = subtitle;
        modalKpi.classList.add('active');
    }

    async function selectAndFocusJob(jobId) {
        let job = allJobs.find(j => j.id === jobId);
        if (!job) {
            try {
                const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`);
                if (res.ok) {
                    const data = await res.json();
                    if (data) {
                        allJobs.push(data);
                    }
                }
            } catch (err) {
                console.error("Erro ao carregar detalhes do job arquivado:", err);
            }
        }
        activeJobId = jobId;
        document.getElementById('modal-kpi-jobs').classList.remove('active');
        renderQueue();
    }

    let pollInterval = null;
    let lastPollWasActive = null;

    function startPolling() {
        // Adaptive polling: 2.5s when jobs are running/pending, 8s when idle
        const hasActiveJobs = allJobs.some(j => j.state === 'running' || j.state === 'pending');
        const targetInterval = hasActiveJobs ? 2500 : 8000;

        if (lastPollWasActive !== hasActiveJobs) {
            lastPollWasActive = hasActiveJobs;
            if (pollInterval) clearInterval(pollInterval);
            pollInterval = setInterval(async () => {
                await loadJobs();
                startPolling(); // re-evaluate interval after each load
            }, targetInterval);
        }
    }

    // --- JIRA View and Synchronization Logic ---
    let jiraTickets = [];
    let jiraSortColumn = null;
    let jiraSortAscending = true;
    let activeKpiCardFilter = null;

    const TEAM_MEMBERS = {
        "Core Systems": ["Clayton Lopes", "Rita Rodrigues", "Filipe Galego", "Paula Silva", "José Pereira"],
        "Helpdesk": ["Filipe Abreu", "Miguel Ribeiro", "Alexandre Rodrigues"],
        "Retail Systems": ["Vitor.Pereira", "Marisa Moreira", "Sandra Gomes"],
        "Digital": ["Sandra Gomes", "Vitor Silva", "Diogo Oliveira"],
        "Systems Administration and Network": ["Alexandre Rodrigues"],
        "Business Intelligence": ["Mariana Pinto"],
        "Development": ["Joao.Pinheiro", "Pedro Silva"]
    };

    const JIRA_SUPPLIERS = [
        "ABACO", "ABAP", "ADYEN", "Axians", "BIT", "Bizdirect", "Bloomreach", "Canon", "Cegid", "Centric", "Claranet", "Cycleon", "DCP", "Decunify", "Deepidoo", "Deloitte", "Devscope", "DHL", "DILAX", "DIMEP", "DPD", "DSA", "Evolutive", "Google", "INASE", "Indra", "INETUM", "Inside", "Konica Minolta", "Lenovo", "MCube", "Microsoft", "Milestone", "Millenium", "Mirame", "Movistar", "Neogrid", "NOS", "OMS", "Orbcom", "OSF", "Outsystems", "Pamafe", "Paypal", "Planet", "Redicom", "RHPro", "S21", "Sales Force", "SAP", "Saphety", "SBX", "SEUR", "SIBS", "SISQUAL", "Splio", "Suporte Fashion", "Suporte Losan", "Suporte Salsa", "Tlantic", "Valantic", "Vodafone", "Winprovit Field", "Zetes"
    ];

    function setKpiCardFilter(filter) {
        if (filter === 'total') {
            activeKpiCardFilter = null;
        } else if (activeKpiCardFilter === filter) {
            activeKpiCardFilter = null;
        } else {
            activeKpiCardFilter = filter;
        }
        updateKpiCardHighlight();
        filterAndRenderJiraTickets();
    }

    function resetJiraFilters(rerender = false) {
        const jiraFilters = [
            'jira-filter-team',
            'jira-filter-stream',
            'jira-filter-ticket-type',
            'jira-filter-priority',
            'jira-filter-assignee',
            'jira-filter-creator',
            'jira-filter-process',
            'jira-filter-supplier'
        ];
        jiraFilters.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = '';
        });

        const filterKey = document.getElementById('jira-filter-key');
        if (filterKey) filterKey.value = '';

        const searchInput = document.getElementById('jira-search-input');
        if (searchInput) searchInput.value = '';

        jiraSortColumn = null;
        jiraSortAscending = true;
        activeKpiCardFilter = null;
        updateKpiCardHighlight();

        if (rerender) {
            filterAndRenderJiraTickets();
        }
    }

    function updateKpiCardHighlight() {
        const cardIds = {
            'total': 'jira-card-total',
            'today': 'jira-card-today',
            'review': 'jira-card-review',
            'wip': 'jira-card-wip',
            'waiting-customer': 'jira-card-waiting-customer',
            'waiting-support': 'jira-card-waiting-support',
            'incidents': 'jira-card-incidents',
            'unassigned': 'jira-card-unassigned',
            'sla-breached': 'jira-card-sla-breached'
        };

        for (const [filter, id] of Object.entries(cardIds)) {
            const el = document.getElementById(id);
            if (el) {
                if ((filter === 'total' && activeKpiCardFilter === null) || activeKpiCardFilter === filter) {
                    el.classList.add('active-filter');
                } else {
                    el.classList.remove('active-filter');
                }
            }
        }
    }

    async function updateTicketAssignee(key, newAssignee) {
        try {
            const res = await fetch(`/api/jira/tickets/${encodeURIComponent(key)}/assign`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ assignee: newAssignee })
            });
            if (res.ok) {
                const data = await res.json();
                
                // Update local list
                const ticket = jiraTickets.find(t => t.key === key);
                if (ticket) {
                    ticket.assignee = newAssignee;
                }
                
                // Show notification toast
                if (data.jira_updated) {
                    showToast(`Responsável do ticket ${key} atualizado no JIRA para "${newAssignee || 'Sem responsável'}"!`, 'success');
                } else {
                    showToast(`Responsável do ticket ${key} atualizado localmente para "${newAssignee || 'Sem responsável'}".`, 'info');
                }
                
                // Refresh JIRA KPI metrics and rendering
                filterAndRenderJiraTickets();
            } else {
                showToast('Erro ao atualizar o responsável.', 'error');
            }
        } catch (err) {
            console.error(err);
            showToast('Erro ao comunicar com o servidor.', 'error');
        }
    }

    async function updateTicketType(key, newType, selectEl) {
        try {
            const res = await fetch(`/api/jira/tickets/${encodeURIComponent(key)}/type`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ticket_type: newType })
            });
            if (res.ok) {
                const data = await res.json();
                
                // Update local list
                const ticket = jiraTickets.find(t => t.key === key);
                if (ticket) {
                    ticket.ticket_type = newType;
                }
                
                // Show notification toast
                if (data.jira_updated) {
                    showToast(`Tipo do ticket ${key} atualizado no JIRA para "${newType || '-'}"!`, 'success');
                } else {
                    showToast(`Tipo do ticket ${key} atualizado localmente para "${newType || '-'}".`, 'info');
                }
                
                // Update select element background/style dynamically
                if (newType.toLowerCase() === 'incident') {
                    selectEl.classList.add('is-incident');
                } else {
                    selectEl.classList.remove('is-incident');
                }

                // Refresh JIRA KPI metrics and rendering
                filterAndRenderJiraTickets();
            } else {
                showToast('Erro ao atualizar o tipo do ticket.', 'error');
            }
        } catch (err) {
            console.error(err);
            showToast('Erro ao comunicar com o servidor.', 'error');
        }
    }

    async function updateTicketSupplier(key, newSupplier) {
        try {
            const res = await fetch(`/api/jira/tickets/${encodeURIComponent(key)}/supplier`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ supplier: newSupplier })
            });
            if (res.ok) {
                const data = await res.json();
                
                // Update local list
                const ticket = jiraTickets.find(t => t.key === key);
                if (ticket) {
                    ticket.supplier = newSupplier;
                }
                
                // Show notification toast
                if (data.jira_updated) {
                    showToast(`Supplier do ticket ${key} atualizado no JIRA para "${newSupplier || 'Sem supplier'}"!`, 'success');
                } else {
                    showToast(`Supplier do ticket ${key} atualizado localmente para "${newSupplier || 'Sem supplier'}".`, 'info');
                }
                
                // Refresh JIRA KPI metrics and rendering
                filterAndRenderJiraTickets();
            } else {
                showToast('Erro ao atualizar o supplier.', 'error');
            }
        } catch (err) {
            console.error(err);
            showToast('Erro ao comunicar com o servidor.', 'error');
        }
    }

    async function loadTransitionsForSelect(ticketKey, selectEl) {
        if (selectEl.dataset.loaded === 'true' || selectEl.dataset.loading === 'true') {
            return;
        }
        
        selectEl.dataset.loading = 'true';
        const loadingOpt = document.createElement('option');
        loadingOpt.value = "";
        loadingOpt.disabled = true;
        loadingOpt.text = "A carregar transições...";
        selectEl.appendChild(loadingOpt);
        
        try {
            const res = await fetch(`/api/jira/tickets/${encodeURIComponent(ticketKey)}/transitions`);
            if (!res.ok) throw new Error();
            const data = await res.json();
            const transitions = data.transitions || [];
            
            const currentText = selectEl.options[0]?.text || selectEl.value;
            selectEl.innerHTML = '';
            
            const currentOpt = document.createElement('option');
            currentOpt.value = "";
            currentOpt.text = currentText;
            currentOpt.selected = true;
            selectEl.appendChild(currentOpt);
            
            transitions.forEach(t => {
                const opt = document.createElement('option');
                opt.value = `${t.id}|${t.name}`;
                opt.text = t.name;
                opt.style.background = "#ffffff";
                opt.style.color = "#334155";
                opt.style.textTransform = "uppercase";
                opt.style.fontWeight = "bold";
                selectEl.appendChild(opt);
            });
            
            selectEl.dataset.loaded = 'true';
        } catch (err) {
            console.error('Error loading transitions:', err);
            if (loadingOpt.parentNode) {
                selectEl.removeChild(loadingOpt);
            }
        } finally {
            selectEl.dataset.loading = 'false';
        }
    }

    // Works both via event delegation (no args) and direct call (with args)
    function toggleReplyExpand(ticketKey, expandId) {
        const panel = document.getElementById(expandId);
        const link  = document.getElementById('summary-link-' + ticketKey.replace(/[^a-z0-9]/gi, '-'));
        if (!panel) return;
        const isOpen = panel.classList.toggle('open');
        if (link) link.classList.toggle('expanded', isOpen);
    }

    async function saveReplyComment(ticketKey, expandId) {
        const textarea = document.getElementById('reply-text-' + expandId);
        const saveBtn  = document.querySelector(`.js-reply-save[data-expand-id="${expandId}"]`);
        const commentText = textarea ? textarea.value.trim() : '';

        if (!commentText) {
            if (textarea) {
                textarea.classList.add('field-error');
                textarea.focus();
                textarea.addEventListener('input', function onInput() {
                    textarea.classList.remove('field-error');
                    textarea.removeEventListener('input', onInput);
                }, { once: true });
            }
            showToast('⚠️ Escreva um comentário antes de guardar.', 'warning');
            return;
        }

        if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = 'A guardar...'; }

        try {
            const res = await fetch('/api/jira/tickets/' + encodeURIComponent(ticketKey) + '/comment', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ comment: commentText })
            });
            if (res.ok) {
                const data = await res.json();
                if (data.jira_updated) {
                    showToast('💬 Comentário guardado no JIRA com sucesso!', 'success');
                } else {
                    showToast('💬 Comentário guardado localmente (JIRA não sincronizado).', 'info');
                }
                // Limpar e fechar
                if (textarea) textarea.value = '';
                toggleReplyExpand(ticketKey, expandId);
            } else {
                const err = await res.json().catch(() => ({}));
                showToast('❌ Erro ao guardar comentário: ' + (err.detail || res.statusText), 'error');
            }
        } catch (e) {
            console.error('saveReplyComment error:', e);
            showToast('❌ Erro ao comunicar com o servidor.', 'error');
        } finally {
            if (saveBtn) {
                saveBtn.disabled = false;
                saveBtn.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Guardar comentário';
            }
        }
    }

    async function handleStatusChange(ticketKey, selectEl) {
        const val = selectEl.value;
        if (!val) return;
        
        const [transitionId, statusName] = val.split('|');
        if (!transitionId || !statusName) return;

        // ── Mandatory comment when resolving ─────────────────────
        const isResolved = statusName.toLowerCase().includes('resolv') ||
                           statusName.toLowerCase().includes('resolved') ||
                           statusName.toLowerCase() === 'done';
        if (isResolved) {
            const safeKey   = ticketKey.replace(/[^a-z0-9]/gi, '-');
            const expandId  = `reply-expand-${safeKey}`;
            const textarea  = document.getElementById(`reply-text-${expandId}`);
            const errorMsg  = document.getElementById(`reply-error-${expandId}`);
            const badge     = document.getElementById(`reply-required-badge-${expandId}`);
            const panel     = document.getElementById(expandId);
            const link      = document.getElementById(`summary-link-${safeKey}`);

            if (!textarea || !textarea.value.trim()) {
                // Open expand panel & highlight error
                if (panel && !panel.classList.contains('open')) {
                    panel.classList.add('open');
                    if (link) link.classList.add('expanded');
                }
                if (textarea) {
                    textarea.classList.add('field-error');
                    textarea.focus();
                    textarea.addEventListener('input', function onInput() {
                        if (textarea.value.trim()) {
                            textarea.classList.remove('field-error');
                            if (errorMsg) errorMsg.classList.remove('visible');
                            textarea.removeEventListener('input', onInput);
                        }
                    }, { once: false });
                }
                if (errorMsg) errorMsg.classList.add('visible');
                if (badge)    badge.style.display = '';
                // Reset select back to current status
                selectEl.value = '';
                showToast('⚠️ Preencha o campo "Reply to customer" antes de resolver o ticket.', 'warning');
                return;
            }
        }
        // ─────────────────────────────────────────────────────────
        
        selectEl.disabled = true;
        
        try {
            const res = await fetch(`/api/jira/tickets/${encodeURIComponent(ticketKey)}/transition`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ transition_id: transitionId, status_name: statusName })
            });
            if (res.ok) {
                const data = await res.json();
                
                const ticket = jiraTickets.find(t => t.key === ticketKey);
                if (ticket) {
                    ticket.status = statusName;
                }
                
                if (data.jira_updated) {
                    showToast(`Estado do ticket ${ticketKey} atualizado no JIRA para "${statusName}"!`, 'success');
                } else {
                    showToast(`Estado do ticket ${ticketKey} atualizado localmente para "${statusName}".`, 'info');
                }

                // ── Guardar comentário Reply to customer se preenchido ──
                const safeKeyComment   = ticketKey.replace(/[^a-z0-9]/gi, '-');
                const expandIdComment  = `reply-expand-${safeKeyComment}`;
                const textareaComment  = document.getElementById(`reply-text-${expandIdComment}`);
                const panelComment     = document.getElementById(expandIdComment);
                const linkComment      = document.getElementById(`summary-link-${safeKeyComment}`);
                const commentText      = textareaComment ? textareaComment.value.trim() : '';

                if (commentText) {
                    try {
                        const commentRes = await fetch(`/api/jira/tickets/${encodeURIComponent(ticketKey)}/comment`, {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ comment: commentText })
                        });
                        if (commentRes.ok) {
                            const commentData = await commentRes.json();
                            if (commentData.jira_updated) {
                                showToast(`💬 Comentário "Reply to customer" guardado no JIRA para ${ticketKey}!`, 'success');
                            } else {
                                showToast(`💬 Comentário guardado localmente (JIRA não sincronizado).`, 'info');
                            }
                        } else {
                            showToast(`⚠️ Erro ao guardar comentário no JIRA.`, 'warning');
                        }
                    } catch (commentErr) {
                        console.error('Erro ao guardar comentário:', commentErr);
                        showToast(`⚠️ Erro ao comunicar comentário com o servidor.`, 'warning');
                    }

                    // Limpar textarea e fechar painel
                    if (textareaComment) textareaComment.value = '';
                    if (panelComment && panelComment.classList.contains('open')) {
                        panelComment.classList.remove('open');
                        if (linkComment) linkComment.classList.remove('expanded');
                    }
                }
                // ─────────────────────────────────────────────────────

                filterAndRenderJiraTickets();
            } else {
                showToast('Erro ao transicionar o estado do ticket.', 'error');
                filterAndRenderJiraTickets();
            }
        } catch (err) {
            console.error(err);
            showToast('Erro ao comunicar com o servidor.', 'error');
            filterAndRenderJiraTickets();
        }
    }

    function showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.style.position = 'fixed';
        toast.style.bottom = '24px';
        toast.style.right = '24px';
        toast.style.padding = '12px 24px';
        toast.style.borderRadius = '12px';
        toast.style.fontSize = '13px';
        toast.style.fontWeight = '600';
        toast.style.zIndex = '9999';
        toast.style.boxShadow = '0 10px 25px rgba(0,0,0,0.15)';
        toast.style.display = 'flex';
        toast.style.alignItems = 'center';
        toast.style.gap = '8px';
        toast.style.transition = 'all 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
        toast.style.transform = 'translateY(100px)';
        toast.style.opacity = '0';
        
        let bg = 'var(--primary)';
        let icon = 'ℹ️';
        if (type === 'success') {
            bg = 'var(--success)';
            icon = '✅';
        } else if (type === 'error') {
            bg = 'var(--danger)';
            icon = '❌';
        } else if (type === 'warning') {
            bg = 'var(--warning)';
            icon = '⚠️';
        }
        
        toast.style.backgroundColor = bg;
        toast.style.color = 'white';
        toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
        
        document.body.appendChild(toast);
        
        setTimeout(() => {
            toast.style.transform = 'translateY(0)';
            toast.style.opacity = '1';
        }, 10);
        
        setTimeout(() => {
            toast.style.transform = 'translateY(100px)';
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 4000);
    }

    function setJiraSort(column) {
        if (jiraSortColumn === column) {
            jiraSortAscending = !jiraSortAscending;
        } else {
            jiraSortColumn = column;
            jiraSortAscending = true;
        }
        
        const columns = ['key', 'summary', 'status', 'project', 'team', 'stream', 'ticket_type', 'priority', 'time_to_resolution', 'created_at', 'updated_at', 'creator', 'assignee', 'process', 'supplier'];
        columns.forEach(col => {
            const iconEl = document.getElementById(`sort-icon-${col}`);
            if (iconEl) {
                if (col === jiraSortColumn) {
                    iconEl.innerHTML = jiraSortAscending ? '▲' : '▼';
                    iconEl.style.opacity = '1';
                    iconEl.style.color = 'var(--primary)';
                } else {
                    iconEl.innerHTML = '↕';
                    iconEl.style.opacity = '0.5';
                    iconEl.style.color = '';
                }
            }
        });
        
        filterAndRenderJiraTickets();
    }

    // --- Histórico View and Search Logic ---
    let historyJobs = [];

    async function loadHistoryJobs() {
        const tbody = document.getElementById('history-jobs-table-body');
        if (!tbody) return;

        tbody.innerHTML = `
            <tr>
                <td colspan="6" style="padding: 40px; text-align: center; color: var(--text-secondary);">
                    <div style="font-size: 24px; animation: pulse-opacity 1.5s infinite; display: inline-block;">⏳</div>
                    <p style="margin-top: 8px;">A carregar histórico...</p>
                </td>
            </tr>
        `;

        try {
            const res = await fetch('/api/jobs?include_archived=true&limit=100');
            if (!res.ok) throw new Error('Falha ao carregar histórico');
            const data = await res.json();
            historyJobs = (data.jobs || []).filter(j => j.is_archived);
            filterAndRenderHistoryJobs();
        } catch (err) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" style="padding: 40px; text-align: center; color: var(--danger);">
                        ❌ Erro ao carregar histórico: ${err.message}
                    </td>
                </tr>
            `;
        }
    }

    function filterAndRenderHistoryJobs() {
        const tbody = document.getElementById('history-jobs-table-body');
        const query = document.getElementById('history-search-input').value.toLowerCase().trim();
        if (!tbody) return;

        const filtered = historyJobs.filter(job => {
            const proc = job.params ? (job.params.subprocesso || job.params.processo || job.task) : job.task;
            const env = job.ambiente || job.params?.ambiente || 'DEV';
            return job.id.toLowerCase().includes(query) || proc.toLowerCase().includes(query) || env.toLowerCase().includes(query);
        });

        if (filtered.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" style="padding: 40px; text-align: center; color: var(--text-secondary);">
                        Nenhum job arquivado encontrado.
                    </td>
                </tr>
            `;
            return;
        }

        tbody.innerHTML = filtered.map(job => {
            const shortId = job.id.substring(0, 8);
            const dateStr = new Date(job.created_at).toLocaleString('pt-PT', { 
                day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit'
            });
            
            const proc = job.params ? (job.params.subprocesso || job.params.processo || job.task) : job.task;
            const env = job.ambiente || job.params?.ambiente || 'DEV';
            
            let badgeClass = '';
            let badgeText = job.state;
            if (job.state === 'running') { badgeClass = 'running'; badgeText = 'executando'; }
            else if (job.state === 'pending') { badgeClass = 'pending'; badgeText = 'pendente'; }
            else if (job.state === 'succeeded') { badgeClass = 'success'; badgeText = 'sucesso'; }
            else if (job.state === 'failed') { badgeClass = 'failed'; badgeText = 'erro'; }

            return `
                <tr class="kpi-table-row">
                    <td style="padding: 14px 16px; font-family: monospace; font-weight: bold; color: var(--text-secondary);">#${shortId}</td>
                    <td style="padding: 14px 16px; white-space: nowrap;">${dateStr}</td>
                    <td style="padding: 14px 16px; font-weight: 600; color: var(--text-primary); max-width: 250px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${escapeHtml(proc)}">${escapeHtml(proc)}</td>
                    <td style="padding: 14px 16px;">
                        <span style="font-size: 11px; font-weight: bold; padding: 2px 6px; background: rgba(0,0,0,0.03); border: 1px solid var(--border-color); border-radius: 4px;">${escapeHtml(env)}</span>
                    </td>
                    <td style="padding: 14px 16px;">
                        <span class="badge ${badgeClass}">${badgeText}</span>
                    </td>
                    <td style="padding: 14px 16px; text-align: center; white-space: nowrap;">
                        <button class="btn btn-primary" style="padding: 4px 10px; font-size: 11px; border-radius: 6px;" onclick="focusArchivedJobAndSwitch('${job.id}')">🔍 Focar</button>
                        <button class="btn btn-danger" style="padding: 4px 10px; font-size: 11px; border-radius: 6px; margin-left: 6px;" onclick="deleteJobFromDb('${job.id}')">🗑️ Eliminar</button>
                    </td>
                </tr>
            `;
        }).join('');
    }

    async function focusArchivedJobAndSwitch(jobId) {
        let job = allJobs.find(j => j.id === jobId);
        if (!job) {
            try {
                const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`);
                if (res.ok) {
                    const data = await res.json();
                    if (data) {
                        allJobs.push(data);
                    }
                }
            } catch (err) {
                console.error("Erro ao carregar detalhes do job arquivado:", err);
            }
        }
        activeJobId = jobId;
        switchView('visao-geral');
        renderQueue();
    }

    async function deleteJobFromDb(jobId) {
        if (!confirm('Deseja eliminar permanentemente este job da base de dados? Esta ação não pode ser desfeita.')) {
            return;
        }
        try {
            const res = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`, { method: 'DELETE' });
            if (res.ok) {
                allJobs = allJobs.filter(j => j.id !== jobId);
                if (activeJobId === jobId) {
                    activeJobId = null;
                }
                await loadHistoryJobs();
            } else {
                alert('Erro ao eliminar o job.');
            }
        } catch(e) {
            console.error('Erro de rede ao eliminar:', e);
            alert('Erro de rede ao eliminar o job.');
        }
    }

    function switchView(viewName) {
        // 1. Remove active class from all nav items
        document.querySelectorAll('.nav-menu .nav-item').forEach(item => {
            item.classList.remove('active');
        });

        // 2. Set active class and toggle views
        const visaoGeralView = document.getElementById('view-visao-geral');
        const agentSalsaItView = document.getElementById('view-agent-salsa-it');
        const jiraView = document.getElementById('view-jira');
        const historicoView = document.getElementById('view-historico');
        const headerTitle = document.querySelector('.top-header .header-titles h2');
        const headerSub = document.querySelector('.top-header .header-titles p');
        const jiraSyncWrapper = document.getElementById('jira-sync-wrapper');
        const refreshBtn = document.getElementById('refresh-button');
        const newJobBtn = document.getElementById('new-job-btn');

        if (viewName === 'visao-geral') {
            document.getElementById('nav-item-visao-geral').classList.add('active');
            visaoGeralView.style.display = 'flex';
            jiraView.style.display = 'none';
            if (agentSalsaItView) agentSalsaItView.style.display = 'none';
            if (historicoView) historicoView.style.display = 'none';
            if (headerTitle) headerTitle.textContent = 'Cockpit SAP Script';
            if (headerSub) headerSub.textContent = 'Monitorização moderna dos workers, jobs e logs SAP GUI Scripting.';
            if (jiraSyncWrapper) jiraSyncWrapper.style.display = 'none';
            if (refreshBtn) refreshBtn.style.display = 'inline-flex';
            if (newJobBtn) newJobBtn.style.display = 'inline-flex';
        } else if (viewName === 'jira') {
            document.getElementById('nav-item-jira').classList.add('active');
            visaoGeralView.style.display = 'none';
            jiraView.style.display = 'flex';
            const dashView = document.getElementById('view-jira-dashboard');
            if (dashView) dashView.style.display = 'none';
            const twrapper = document.getElementById('view-jira-tickets-wrapper');
            if (twrapper) twrapper.style.display = 'flex';
            if (agentSalsaItView) agentSalsaItView.style.display = 'none';
            if (historicoView) historicoView.style.display = 'none';
            if (headerTitle) headerTitle.textContent = 'Fila de Tickets';
            if (headerSub) headerSub.textContent = 'Tickets em aberto e sincronização automática.';
            if (jiraSyncWrapper) jiraSyncWrapper.style.display = 'inline-flex';
            if (refreshBtn) refreshBtn.style.display = 'none';
            if (newJobBtn) newJobBtn.style.display = 'none';
            resetJiraFilters(false);
            loadJiraTickets();
        } else if (viewName === 'jira-dashboard') {
            const navItem = document.getElementById('nav-item-jira-dashboard');
            if (navItem) navItem.classList.add('active');
            visaoGeralView.style.display = 'none';
            jiraView.style.display = 'flex';
            const dashView = document.getElementById('view-jira-dashboard');
            if (dashView) dashView.style.display = 'flex';
            if (agentSalsaItView) agentSalsaItView.style.display = 'none';
            if (historicoView) historicoView.style.display = 'none';
            if (headerTitle) headerTitle.textContent = 'Dashboard de Tickets';
            if (headerSub) headerSub.textContent = 'Visão gráfica e analítica dos tickets de suporte.';
            if (jiraSyncWrapper) jiraSyncWrapper.style.display = 'none';
            if (refreshBtn) refreshBtn.style.display = 'none';
            if (newJobBtn) newJobBtn.style.display = 'none';
            // Refresh the backing ticket list so the dashboard stays in sync.
            loadJiraTickets(true).then(() => applyDashFilters());
        } else if (viewName === 'historico') {
            const navItem = document.getElementById('nav-item-historico');
            if (navItem) navItem.classList.add('active');
            visaoGeralView.style.display = 'none';
            jiraView.style.display = 'none';
            if (agentSalsaItView) agentSalsaItView.style.display = 'none';
            const dashView = document.getElementById('view-jira-dashboard');
            if (dashView) dashView.style.display = 'none';
            if (historicoView) historicoView.style.display = 'flex';
            if (headerTitle) headerTitle.textContent = 'Histórico de Jobs Arquivados';
            if (headerSub) headerSub.textContent = 'Visualização e gestão de rotinas antigas e arquivadas.';
            if (jiraSyncWrapper) jiraSyncWrapper.style.display = 'none';
            if (refreshBtn) refreshBtn.style.display = 'none';
            if (newJobBtn) newJobBtn.style.display = 'none';
            loadHistoryJobs();
        } else if (viewName === 'definicoes') {
            const navItem = document.getElementById('nav-item-definicoes');
            if (navItem) navItem.classList.add('active');
            visaoGeralView.style.display = 'none';
            jiraView.style.display = 'flex';
            const dashView = document.getElementById('view-jira-dashboard');
            if (dashView) dashView.style.display = 'none';
            const twrapper = document.getElementById('view-jira-tickets-wrapper');
            if (twrapper) twrapper.style.display = 'none';
            if (agentSalsaItView) agentSalsaItView.style.display = 'none';
            if (historicoView) historicoView.style.display = 'none';
            const defView = document.getElementById('view-definicoes');
            if (defView) defView.style.display = 'flex';
            if (headerTitle) headerTitle.textContent = 'Definições';
            if (headerSub) headerSub.textContent = 'Parâmetros de contexto para o Agente SAP.';
            if (jiraSyncWrapper) jiraSyncWrapper.style.display = 'none';
            if (refreshBtn) refreshBtn.style.display = 'none';
            if (newJobBtn) newJobBtn.style.display = 'none';
            defLoadRules();
        } else if (viewName === 'agent-salsa-it') {
            const navItem = document.getElementById('nav-item-agent-salsa-it');
            if (navItem) navItem.classList.add('active');
            visaoGeralView.style.display = 'none';
            jiraView.style.display = 'none';
            if (agentSalsaItView) agentSalsaItView.style.display = 'flex';
            if (historicoView) historicoView.style.display = 'none';
            const dashView = document.getElementById('view-jira-dashboard');
            if (dashView) dashView.style.display = 'none';
            const twrapper = document.getElementById('view-jira-tickets-wrapper');
            if (twrapper) twrapper.style.display = 'none';
            const defView = document.getElementById('view-definicoes');
            if (defView) defView.style.display = 'none';
            if (headerTitle) headerTitle.textContent = 'Agente Salsa IT';
            if (headerSub) headerSub.textContent = 'Assistente operacional para processos SAP.';
            if (jiraSyncWrapper) jiraSyncWrapper.style.display = 'none';
            if (refreshBtn) refreshBtn.style.display = 'none';
            if (newJobBtn) newJobBtn.style.display = 'none';
            asiInitChat();
        }
    }

    const initialSelectOptions = {};
    const staticFilterIds = [
        'jira-filter-team',
        'jira-filter-stream',
        'jira-filter-ticket-type',
        'jira-filter-priority',
        'jira-filter-process'
    ];
    staticFilterIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            initialSelectOptions[id] = Array.from(el.options).map(opt => ({
                value: opt.value,
                text: opt.textContent
            }));
        }
    });

    function populateAllDropdownsFaceted(allTickets) {
        const fields = {
            'jira-filter-team': { key: 'team', defaultText: 'Todas as equipas', isStatic: true },
            'jira-filter-stream': { key: 'stream', defaultText: 'Todos os streams', isStatic: true },
            'jira-filter-ticket-type': { key: 'ticket_type', defaultText: 'Todos os tipos', isStatic: true },
            'jira-filter-priority': { key: 'priority', defaultText: 'Todas as prioridades', isStatic: true },
            'jira-filter-process': { key: 'process', defaultText: 'Todos os processos', isStatic: true },
            'jira-filter-supplier': { key: 'supplier', defaultText: 'Todos os suppliers' },
            'jira-filter-assignee': { key: 'assignee', defaultText: 'Todos os responsáveis' },
            'jira-filter-creator': { key: 'creator', defaultText: 'Todos os criadores' }
        };

        // Read current filter values
        const currentValues = {};
        for (const id of Object.keys(fields)) {
            currentValues[id] = document.getElementById(id)?.value || '';
        }
        const valKey = (document.getElementById('jira-filter-key')?.value || '').toLowerCase().trim();
        const query = (document.getElementById('jira-search-input')?.value || '').toLowerCase().trim();

        // For each filter dropdown
        for (const [id, config] of Object.entries(fields)) {
            const selectEl = document.getElementById(id);
            if (!selectEl) continue;

            const prevValue = selectEl.value;

            // Filter tickets using all other criteria EXCEPT current select field
            const filteredTickets = allTickets.filter(t => {
                if (valKey && (!t.key || !t.key.toLowerCase().includes(valKey))) return false;
                if (query && (!t.summary || !t.summary.toLowerCase().includes(query))) return false;

                for (const [otherId, otherConfig] of Object.entries(fields)) {
                    if (otherId === id) continue; // skip current field
                    const val = currentValues[otherId];
                    if (val && t[otherConfig.key] !== val) return false;
                }
                return true;
            });

            // Extract unique values present in the filtered subset
            let presentValues = new Set(
                filteredTickets.map(t => t[config.key]).filter(v => v && String(v).trim() !== "")
            );
            if (config.key === 'assignee') {
                const activeTeamFilter = currentValues['jira-filter-team'];
                if (activeTeamFilter && TEAM_MEMBERS[activeTeamFilter]) {
                    // Keep present assignees for this team, and add all predefined members of this specific team
                    const teamMembers = TEAM_MEMBERS[activeTeamFilter];
                    teamMembers.forEach(val => presentValues.add(val));
                } else {
                    // No team filter is active, show all present assignees, and make sure Clayton Lopes is there
                    presentValues.add("Clayton Lopes");
                    // Add other known team members from TEAM_MEMBERS to fallback list
                    for (const members of Object.values(TEAM_MEMBERS)) {
                        members.forEach(val => presentValues.add(val));
                    }
                }
            }


            let html = ``;

            if (config.isStatic) {
                // For static select, read from stored initial options and show all options
                const initialOpts = initialSelectOptions[id] || [];
                initialOpts.forEach(opt => {
                    html += `<option value="${escapeHtml(opt.value)}" style="background: #ffffff; color: #334155;">${escapeHtml(opt.text)}</option>`;
                });
            } else {
                // For dynamic select, rebuild from unique values in filtered tickets
                html += `<option value="" style="background: #ffffff; color: #334155;">${config.defaultText}</option>`;
                const sortedVals = Array.from(presentValues).map(v => String(v).trim()).sort((a, b) => a.localeCompare(b));
                sortedVals.forEach(val => {
                    html += `<option value="${escapeHtml(val)}" style="background: #ffffff; color: #334155;">${escapeHtml(val)}</option>`;
                });
            }

            selectEl.innerHTML = html;
            selectEl.value = prevValue; // restore value
        }
    }

    async function loadJiraTickets(silent = false) {
        const tbody = document.getElementById('jira-tickets-table-body');
        if (!tbody) return;

        // Only show loading spinner on first load; silent mode keeps existing rows visible
        if (!silent) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="14" style="padding: 40px; text-align: center; color: var(--text-secondary);">
                        <div style="font-size: 24px; animation: pulse-opacity 1.5s infinite; display: inline-block;">⏳</div>
                        <p style="margin-top: 8px;">A carregar tickets do JIRA...</p>
                    </td>
                </tr>
            `;
        }

        try {
            const res = await fetch('/api/jira/tickets?limit=50000&exclude_closed=true');
            if (!res.ok) throw new Error('Falha ao carregar tickets');
            const data = await res.json();
            jiraTickets = data.tickets || [];

            filterAndRenderJiraTickets();
            updateKpiCardHighlight();

            // Update last sync timestamp
            const now = new Date();
            const pad = n => String(n).padStart(2, '0');
            const syncText = `${pad(now.getDate())}/${pad(now.getMonth()+1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
            const el = document.getElementById('jira-last-sync-text');
            if (el) el.textContent = `Última sincronização: ${syncText}`;
        } catch (err) {
            if (!silent) {
                tbody.innerHTML = `
                    <tr>
                        <td colspan="14" style="padding: 40px; text-align: center; color: var(--danger);">
                            ❌ Erro ao carregar tickets: ${err.message}
                        </td>
                    </tr>
                `;
            }
        }
    }

    function filterAndRenderJiraTickets() {
        const tbody = document.getElementById('jira-tickets-table-body');
        if (!tbody) return;

        const normalizeValue = value => String(value ?? '').trim();
        const normalizeValueLower = value => normalizeValue(value).toLowerCase();

        // Build team to assignees map
        const teamAssigneesMap = {};
        for (const [team, members] of Object.entries(TEAM_MEMBERS)) {
            teamAssigneesMap[team] = new Set(members);
        }

        jiraTickets.forEach(t => {
            if (t.team && t.assignee && t.assignee !== 'Sem responsável' && t.assignee.trim() !== '') {
                const team = t.team.trim();
                const assignee = t.assignee.trim();
                if (!teamAssigneesMap[team]) {
                    teamAssigneesMap[team] = new Set();
                }
                teamAssigneesMap[team].add(assignee);
            }
        });
        
        // Build list of unique assignees (fallback)
        const uniqueAssignees = new Set();
        jiraTickets.forEach(t => {
            if (t.assignee && t.assignee !== 'Sem responsável' && t.assignee.trim() !== '') {
                uniqueAssignees.add(t.assignee.trim());
            }
        });
        uniqueAssignees.add("Clayton Lopes");
        for (const members of Object.values(TEAM_MEMBERS)) {
            members.forEach(m => uniqueAssignees.add(m));
        }
        const sortedAssignees = Array.from(uniqueAssignees).sort((a, b) => a.localeCompare(b));

        // 1. Faceted update: refresh dropdown options list based on active filters (only for open tickets to avoid freezing the browser)
        const openTicketsOnly = jiraTickets.filter(t => {
            const s = (t.status || '').toLowerCase().trim();
            return !['done', 'closed', 'concluído', 'resolvido', 'fechado', 'fechada', 'cancelled'].includes(s);
        });
        populateAllDropdownsFaceted(openTicketsOnly);

        // 2. Read values of active filters (which are preserved during populating)
        const valTeam = document.getElementById('jira-filter-team')?.value || '';
        const valStream = document.getElementById('jira-filter-stream')?.value || '';
        const valType = document.getElementById('jira-filter-ticket-type')?.value || '';
        const valPriority = document.getElementById('jira-filter-priority')?.value || '';
        const valAssignee = document.getElementById('jira-filter-assignee')?.value || '';
        const valCreator = document.getElementById('jira-filter-creator')?.value || '';
        const valProcess = document.getElementById('jira-filter-process')?.value || '';
        const valSupplier = document.getElementById('jira-filter-supplier')?.value || '';
        const valKey = (document.getElementById('jira-filter-key')?.value || '').toLowerCase().trim();
        const query = (document.getElementById('jira-search-input')?.value || '').toLowerCase().trim();

        const CLOSED_STATUSES = new Set(['done', 'closed', 'concluído', 'resolvido', 'fechado', 'fechada', 'cancelled']);

        const filtered = jiraTickets.filter(t => {
            // Excluir sempre tickets fechados/concluídos da fila activa
            if (CLOSED_STATUSES.has(normalizeValueLower(t.status))) return false;

            if (valTeam && normalizeValueLower(t.team) !== normalizeValueLower(valTeam)) return false;
            if (valStream && normalizeValueLower(t.stream) !== normalizeValueLower(valStream)) return false;
            if (valType && normalizeValueLower(t.ticket_type) !== normalizeValueLower(valType)) return false;
            if (valPriority && normalizeValueLower(t.priority) !== normalizeValueLower(valPriority)) return false;
            if (valAssignee && normalizeValueLower(t.assignee) !== normalizeValueLower(valAssignee)) return false;
            if (valCreator && normalizeValueLower(t.creator) !== normalizeValueLower(valCreator)) return false;
            if (valProcess && normalizeValueLower(t.process) !== normalizeValueLower(valProcess)) return false;
            if (valSupplier && normalizeValueLower(t.supplier) !== normalizeValueLower(valSupplier)) return false;
            if (valKey && (!t.key || !normalizeValueLower(t.key).includes(valKey))) return false;
            if (query && (!t.summary || !normalizeValueLower(t.summary).includes(query))) return false;

            // Apply JIRA KPI Card filter
            if (activeKpiCardFilter) {
                const status = (t.status || '').toLowerCase().trim();
                const type = (t.ticket_type || '').toLowerCase().trim();
                const sla = (t.time_to_resolution || '').toLowerCase();

                if (activeKpiCardFilter === 'today') {
                    const now = new Date();
                    const createdDate = new Date(t.created_at);
                    const isToday = createdDate.getDate() === now.getDate() &&
                                    createdDate.getMonth() === now.getMonth() &&
                                    createdDate.getFullYear() === now.getFullYear();
                    if (!isToday) return false;
                }
                if (activeKpiCardFilter === 'review' && status !== 'in review') return false;
                if (activeKpiCardFilter === 'wip' && status !== 'work in progress') return false;
                if (activeKpiCardFilter === 'waiting-customer' && status !== 'waiting for customer') return false;
                if (activeKpiCardFilter === 'waiting-support' && status !== 'waiting for specialized support') return false;
                if (activeKpiCardFilter === 'incidents' && type !== 'incident') return false;
                if (activeKpiCardFilter === 'unassigned') {
                    const isUnassigned = !t.assignee || t.assignee.trim() === "" || t.assignee.toLowerCase() === "sem responsável";
                    if (!isUnassigned) return false;
                }
                if (activeKpiCardFilter === 'sla-breached') {
                    const isSlaBreached = t.time_to_resolution && t.time_to_resolution.toLowerCase().includes("excedido");
                    if (!isSlaBreached) return false;
                }
            }

            return true;
        });

        if (jiraSortColumn) {
            filtered.sort((a, b) => {
                let valA = a[jiraSortColumn] || '';
                let valB = b[jiraSortColumn] || '';
                
                // Case-insensitive string comparison
                valA = String(valA).toLowerCase().trim();
                valB = String(valB).toLowerCase().trim();
                
                // Natural sort for key (like IZ-123 vs IZ-12)
                if (jiraSortColumn === 'key') {
                    return valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' }) * (jiraSortAscending ? 1 : -1);
                }
                
                if (valA < valB) return jiraSortAscending ? -1 : 1;
                if (valA > valB) return jiraSortAscending ? 1 : -1;
                return 0;
            });
        }

        // Não agrupar mais como linhas separadas, vamos usar 'filtered' diretamente
        if (filtered.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="14" style="padding: 40px; text-align: center; color: var(--text-secondary);">
                        Nenhum ticket encontrado.
                    </td>
                </tr>
            `;
            updateJiraKpis([]);
            return;
        }

        tbody.innerHTML = filtered.map(t => {
            const ticketKey = normalizeValue(t.key);
            const statusLower = normalizeValueLower(t.status);
            let badgeStyle = 'background: rgba(255,255,255,0.05); color: #9ca3af; border: 1px solid rgba(255,255,255,0.1);';
            
            if (statusLower.includes('done') || statusLower.includes('closed') || statusLower.includes('concluído') || statusLower.includes('resolvido')) {
                badgeStyle = 'background: rgba(16, 185, 129, 0.12); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.25);';
            } else if (statusLower.includes('progress') || statusLower.includes('desenvolvimento') || statusLower.includes('review') || statusLower.includes('teste')) {
                badgeStyle = 'background: rgba(59, 130, 246, 0.12); color: #3b82f6; border: 1px solid rgba(59, 130, 246, 0.25);';
            } else if (statusLower.includes('todo') || statusLower.includes('open') || statusLower.includes('aberto') || statusLower.includes('novo')) {
                badgeStyle = 'background: rgba(245, 158, 11, 0.12); color: #f59e0b; border: 1px solid rgba(245, 158, 11, 0.25);';
            }

            const jiraBase = (window.__COCKPIT__ && window.__COCKPIT__.jiraBase) || 'https://salsajeans.atlassian.net';
            const ticketUrl = `${jiraBase}/browse/${ticketKey}`;

            let slaBadge = '<span style="color: #9ca3af;">-</span>';
            if (t.time_to_resolution) {
                const slaLower = t.time_to_resolution.toLowerCase();
                let slaBadgeStyle = 'background: rgba(255,255,255,0.05); color: #9ca3af; border: 1px solid rgba(255,255,255,0.1);';
                if (slaLower.includes('excedido') || slaLower.startsWith('-')) {
                    slaBadgeStyle = 'background: rgba(239, 68, 68, 0.12); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.25);';
                } else if (slaLower.includes('resolvido')) {
                    slaBadgeStyle = 'background: rgba(16, 185, 129, 0.12); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.25);';
                } else if (slaLower.includes('h') || slaLower.includes('m')) {
                    slaBadgeStyle = 'background: rgba(59, 130, 246, 0.12); color: #3b82f6; border: 1px solid rgba(59, 130, 246, 0.25);';
                }
                slaBadge = `
                    <span style="display: inline-block; padding: 4px 8px; border-radius: 6px; font-size: 11px; font-weight: 600; ${slaBadgeStyle}">
                        ${escapeHtml(t.time_to_resolution)}
                    </span>
                `;
            }

            const expandId = `reply-expand-${ticketKey.replace(/[^a-z0-9]/gi,'-')}`;
            
            // Colocar as chaves linkadas na mesma célula
            let keyDisplay = `<a href="${ticketUrl}" target="_blank" style="color: var(--primary); text-decoration: none; display: inline-flex; align-items: center; gap: 4px;">${escapeHtml(ticketKey)} 🔗</a>`;
            if (t.linked_keys && t.linked_keys.length > 0) {
                const linksHtml = t.linked_keys.map(k => {
                    return `<div style="margin-top: 4px;"><a href="${jiraBase}/browse/${k}" target="_blank" style="color: var(--text-secondary); text-decoration: none; font-size: 0.9em; display: inline-flex; align-items: center; gap: 4px;">↳ ${escapeHtml(k)}</a></div>`;
                }).join('');
                keyDisplay += linksHtml;
            }

            return `
                    <tr class="kpi-table-row" id="row-${ticketKey.replace(/[^a-z0-9]/gi,'-')}">
                    <td style="padding: 14px 16px; font-family: monospace; font-weight: bold;">
                        ${keyDisplay}
                    </td>
                    <td style="padding: 14px 16px; color: var(--text-primary); font-weight: 500;">
                        <a class="summary-link js-reply-toggle" id="summary-link-${ticketKey.replace(/[^a-z0-9]/gi,'-')}" data-key="${ticketKey.replace(/"/g,'&quot;')}" data-expand-id="${expandId}" style="cursor:pointer;">
                            ${escapeHtml(t.summary)}
                            <span class="link-icon">▾</span>
                        </a>
                    </td>
                    <td style="padding: 10px 16px; vertical-align: middle;">
                        <select class="status-cell-select" style="${badgeStyle}" 
                                onfocus="loadTransitionsForSelect('${ticketKey}', this)" 
                                onchange="handleStatusChange('${ticketKey}', this)">
                            <option value="" selected>${escapeHtml(normalizeValue(t.status) || '-')}</option>
                        </select>
                    </td>
                    <td style="padding: 14px 16px; color: var(--text-secondary);">
                        ${t.team ? escapeHtml(t.team) : '<span style="color: #9ca3af;">-</span>'}
                    </td>
                    <td style="padding: 14px 16px; color: var(--text-secondary);">
                        ${t.stream ? escapeHtml(t.stream) : '<span style="color: #9ca3af;">-</span>'}
                    </td>
                    <td style="padding: 10px 16px; vertical-align: middle;">
                        ${(() => {
                            const currentType = t.ticket_type ? t.ticket_type.trim() : "";
                            const ticketTypes = ["Service Request", "Incident", "Project"];
                            const typesSet = new Set(ticketTypes);
                            if (currentType) {
                                typesSet.add(currentType);
                            }
                            const rowTypes = Array.from(typesSet).sort((a, b) => a.localeCompare(b));
                            const typeSelectOptions = rowTypes.map(type => {
                                const selected = (currentType && type.toLowerCase() === currentType.toLowerCase()) ? ' selected' : '';
                                return `<option value="${escapeHtml(type)}"${selected}>${escapeHtml(type)}</option>`;
                            }).join('');
                            
                            const isIncident = currentType.toLowerCase() === 'incident';
                            const selectClass = isIncident ? 'ticket-type-cell-select is-incident' : 'ticket-type-cell-select';
                            return `
                                <select class="${selectClass}" onchange="updateTicketType('${ticketKey}', this.value, this)">
                                    ${typeSelectOptions}
                                </select>
                            `;
                        })()}
                    </td>
                    <td style="padding: 14px 16px; color: var(--text-secondary); white-space: nowrap;">${formatJiraDate(t.created_at)}</td>
                    <td style="padding: 14px 16px; color: var(--text-secondary); white-space: nowrap;">${formatJiraDate(t.updated_at)}</td>
                    <td style="padding: 14px 16px; color: var(--text-secondary);">${escapeHtml(t.creator || '-')}</td>
                    <td style="padding: 10px 16px; vertical-align: middle;">
                        ${(() => {
                            const currentAssignee = t.assignee ? t.assignee.trim() : "";
                            const isNoAssignee = !currentAssignee || currentAssignee === "Sem responsável";
                            
                            // Determine assignees options based on the ticket's team
                            let rowAssignees = [];
                            const ticketTeam = t.team ? t.team.trim() : "";
                            if (ticketTeam && teamAssigneesMap[ticketTeam]) {
                                const teamSet = new Set(teamAssigneesMap[ticketTeam]);
                                if (!isNoAssignee) {
                                    teamSet.add(currentAssignee);
                                }
                                rowAssignees = Array.from(teamSet).sort((a, b) => a.localeCompare(b));
                            } else {
                                rowAssignees = sortedAssignees;
                            }
                            
                            const assigneeSelectOptions = rowAssignees.map(name => {
                                const selected = (!isNoAssignee && name.toLowerCase() === currentAssignee.toLowerCase()) ? ' selected' : '';
                                return `<option value="${escapeHtml(name)}"${selected}>${escapeHtml(name)}</option>`;
                            }).join('');
                            return `
                                <select class="assignee-cell-select" onchange="updateTicketAssignee('${ticketKey}', this.value)">
                                    <option value=""${isNoAssignee ? ' selected' : ''}></option>
                                    ${assigneeSelectOptions}
                                </select>
                            `;
                        })()}
                    </td>
                    <td style="padding: 14px 16px; color: var(--text-secondary);">${escapeHtml(t.process || '-')}</td>
                    <td style="padding: 10px 16px; vertical-align: middle;">
                        ${(() => {
                            const currentSupplier = t.supplier ? t.supplier.trim() : "";
                            const supplierSelectOptions = JIRA_SUPPLIERS.map(sup => {
                                const selected = (currentSupplier && sup.toLowerCase() === currentSupplier.toLowerCase()) ? ' selected' : '';
                                return `<option value="${escapeHtml(sup)}"${selected}>${escapeHtml(sup)}</option>`;
                            }).join('');
                            return `
                                <select class="supplier-cell-select" onchange="updateTicketSupplier('${ticketKey}', this.value)">
                                    <option value=""${!currentSupplier ? ' selected' : ''}></option>
                                    ${supplierSelectOptions}
                                </select>
                            `;
                        })()}
                    </td>
                    <td style="padding: 14px 16px; color: var(--text-secondary); font-weight: 600;">
                        ${t.priority ? escapeHtml(t.priority) : '<span style="color: #9ca3af;">-</span>'}
                    </td>
                    <td style="padding: 14px 16px;">${slaBadge}</td>
                </tr>
                <tr class="reply-expand-row" id="${expandId}-row">
                    <td colspan="14">
                        <div class="reply-expand-panel" id="${expandId}">
                            <div class="reply-expand-inner">
                                <div class="reply-expand-header">
                                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#3b82f6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>
                                    Adicionar comentário
                                    <div style="display:flex;gap:4px;margin-left:2px;">
                                        <span style="padding:2px 10px;font-size:11px;font-weight:600;border-radius:20px;background:rgba(59,130,246,0.1);color:#3b82f6;border:1px solid rgba(59,130,246,0.2);">Add internal note</span>
                                        <span style="padding:2px 10px;font-size:11px;font-weight:700;border-radius:20px;background:rgba(59,130,246,0.18);color:#3b82f6;border:1px solid rgba(59,130,246,0.35);text-decoration:underline;cursor:default;">Reply to customer</span>
                                    </div>
                                    <span class="badge-required" id="reply-required-badge-${expandId}" style="display:none;">* Obrigatório para Resolved</span>
                                </div>
                                <textarea class="reply-textarea" id="reply-text-${expandId}" placeholder="Escreva a sua resposta ao cliente... (use /ai para sugestões de IA)" rows="3"></textarea>
                                <div class="reply-error-msg" id="reply-error-${expandId}">
                                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                                    O campo "Reply to customer" é obrigatório quando o status é alterado para <strong style="margin-left:3px;">Resolved</strong>.
                                </div>
                                <div class="reply-actions">
                                    <button class="reply-btn-cancel js-reply-toggle" data-key="${t.key.replace(/"/g,'&quot;')}" data-expand-id="${expandId}">Fechar</button>
                                    <button class="reply-btn-save js-reply-save" data-key="${t.key.replace(/"/g,'&quot;')}" data-expand-id="${expandId}">
                                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>
                                        Guardar comentário
                                    </button>
                                </div>
                            </div>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
        updateJiraKpis(filtered);
    }

    function updateJiraKpis(tickets) {
        const totalEl = document.getElementById('jira-kpi-total');
        const progressEl = document.getElementById('jira-kpi-progress');
        const doneEl = document.getElementById('jira-kpi-done');
        const unassignedEl = document.getElementById('jira-kpi-unassigned');
        const incidentsEl = document.getElementById('jira-kpi-incidents');
        if (!totalEl) return;
 
        totalEl.textContent = tickets.length;

        // Card: Criados Hoje
        const todayEl = document.getElementById('jira-kpi-today');
        if (todayEl) {
            const now = new Date();
            const todayCount = tickets.filter(t => {
                if (!t.created_at) return false;
                const createdDate = new Date(t.created_at);
                return createdDate.getDate() === now.getDate() &&
                       createdDate.getMonth() === now.getMonth() &&
                       createdDate.getFullYear() === now.getFullYear();
            }).length;
            todayEl.textContent = todayCount;
        }
        
        if (progressEl) {
            const inProgressCount = tickets.filter(t => {
                const s = t.status.toLowerCase();
                return s.includes('progress') || s.includes('review') || s.includes('teste');
            }).length;
            progressEl.textContent = inProgressCount;
        }
 
        if (doneEl) {
            const doneCount = tickets.filter(t => {
                const s = t.status.toLowerCase();
                return s.includes('done') || s.includes('closed') || s.includes('resolvido');
            }).length;
            doneEl.textContent = doneCount;
        }
 
        if (unassignedEl) {
            const unassignedCount = tickets.filter(t => {
                return !t.assignee || t.assignee.trim() === "" || t.assignee.toLowerCase() === "sem responsável";
            }).length;
            unassignedEl.textContent = unassignedCount;
 
            const cardEl = document.getElementById('jira-card-unassigned');
            if (cardEl) {
                if (unassignedCount > 0) {
                    cardEl.classList.add('status-red');
                } else {
                    cardEl.classList.remove('status-red');
                }
            }
        }
 
        if (incidentsEl) {
            const incidentsCount = tickets.filter(t => {
                return t.ticket_type && t.ticket_type.toLowerCase() === "incident";
            }).length;
            incidentsEl.textContent = incidentsCount;
 
            const cardEl = document.getElementById('jira-card-incidents');
            if (cardEl) {
                if (incidentsCount !== 0) {
                    cardEl.style.backgroundColor = '';
                    cardEl.style.borderColor = '';
                    cardEl.classList.add('status-red');
                } else {
                    cardEl.classList.remove('status-red');
                }
            }
        }

        // New Card: Work In Progress (WIP)
        const wipEl = document.getElementById('jira-kpi-wip');
        if (wipEl) {
            const wipCount = tickets.filter(t => {
                const s = (t.status || '').toLowerCase().trim();
                return s === 'work in progress';
            }).length;
            wipEl.textContent = wipCount;
        }

        // New Card: In Review
        const reviewEl = document.getElementById('jira-kpi-review');
        if (reviewEl) {
            const reviewCount = tickets.filter(t => {
                const s = (t.status || '').toLowerCase().trim();
                return s === 'in review';
            }).length;
            reviewEl.textContent = reviewCount;
        }

        // New Card: Waiting for Customer
        const customerEl = document.getElementById('jira-kpi-waiting-customer');
        if (customerEl) {
            const customerCount = tickets.filter(t => {
                const s = (t.status || '').toLowerCase().trim();
                return s === 'waiting for customer';
            }).length;
            customerEl.textContent = customerCount;
        }

        // New Card: Waiting for Specialized Support
        const supportEl = document.getElementById('jira-kpi-waiting-support');
        if (supportEl) {
            const supportCount = tickets.filter(t => {
                const s = (t.status || '').toLowerCase().trim();
                return s === 'waiting for specialized support';
            }).length;
            supportEl.textContent = supportCount;
        }
 
        const slaBreachedEl = document.getElementById('jira-kpi-sla-breached');
        if (slaBreachedEl) {
            const breachedCount = tickets.filter(t => {
                return t.time_to_resolution && t.time_to_resolution.toLowerCase().includes("excedido");
            }).length;
            slaBreachedEl.textContent = breachedCount;
 
            const cardEl = document.getElementById('jira-card-sla-breached');
            if (cardEl) {
                if (breachedCount > 0) {
                    cardEl.classList.add('status-red');
                } else {
                    cardEl.classList.remove('status-red');
                }
            }
        }
    }

    // Set up JIRA action listeners
    document.addEventListener('DOMContentLoaded', () => {
        const searchInput = document.getElementById('jira-search-input');
        if (searchInput) {
            searchInput.addEventListener('input', filterAndRenderJiraTickets);
        }

        const jiraFilters = [
            'jira-filter-team',
            'jira-filter-stream',
            'jira-filter-ticket-type',
            'jira-filter-priority',
            'jira-filter-assignee',
            'jira-filter-creator',
            'jira-filter-process',
            'jira-filter-supplier'
        ];

        jiraFilters.forEach(id => {
            const el = document.getElementById(id);
            if (el) {
                el.addEventListener('change', filterAndRenderJiraTickets);
            }
        });

        const filterKey = document.getElementById('jira-filter-key');
        if (filterKey) {
            filterKey.addEventListener('input', filterAndRenderJiraTickets);
        }

        const btnClearFilters = document.getElementById('btn-clear-jira-filters');
        if (btnClearFilters) {
            btnClearFilters.addEventListener('click', () => {
                jiraFilters.forEach(id => {
                    const el = document.getElementById(id);
                    if (el) el.value = '';
                });
                if (filterKey) filterKey.value = '';
                if (searchInput) searchInput.value = '';
                
                // Reset sorting state
                jiraSortColumn = null;
                jiraSortAscending = true;
                const columns = ['key', 'summary', 'status', 'project', 'team', 'stream', 'ticket_type', 'priority', 'time_to_resolution', 'creator', 'assignee', 'process', 'supplier'];
                columns.forEach(col => {
                    const iconEl = document.getElementById(`sort-icon-${col}`);
                    if (iconEl) {
                        iconEl.innerHTML = '↕';
                        iconEl.style.opacity = '0.5';
                        iconEl.style.color = '';
                    }
                });

                // Reset active JIRA KPI Card filter
                activeKpiCardFilter = null;
                updateKpiCardHighlight();
                
                filterAndRenderJiraTickets();
            });
        }

        const syncBtn = document.getElementById('jira-sync-btn');
        const syncIcon = document.getElementById('jira-sync-icon');
        if (syncBtn) {
            syncBtn.addEventListener('click', async () => {
                syncBtn.disabled = true;
                if (syncIcon) syncIcon.style.animation = 'spin 1.5s linear infinite';
                
                try {
                    const res = await fetch('/api/jira/sync', { method: 'POST' });
                    if (!res.ok) throw new Error('Falha ao sincronizar com o JIRA');
                    await loadJiraTickets();
                } catch (err) {
                    alert('Erro durante sincronização manual: ' + err.message);
                } finally {
                    syncBtn.disabled = false;
                    if (syncIcon) syncIcon.style.animation = '';
                }
            });
        }

        // ── Auto-sync JIRA tickets on a configurable interval ────────────
        setInterval(async () => {
            const ticketsWrapper = document.getElementById('view-jira-tickets-wrapper');
            const dashWrapper = document.getElementById('view-jira-dashboard');
            const jiraView = document.getElementById('view-jira');
            const jiraVisible =
                (ticketsWrapper && ticketsWrapper.style.display !== 'none') ||
                (dashWrapper && dashWrapper.style.display !== 'none') ||
                (jiraView && jiraView.style.display !== 'none');
            // Refresh whenever any JIRA surface is visible so the UI does not stay stale.
            if (jiraVisible) {
                await loadJiraTickets(true);
                if (dashWrapper && dashWrapper.style.display !== 'none') {
                    applyDashFilters();
                }
            }
        }, JIRA_POLL_SECONDS * 1000);
        // ─────────────────────────────────────────────────────────────────

        // Set up Dashboard filter listeners
        const dashFilterIds = [
            'dash-filter-team', 'dash-filter-stream', 'dash-filter-ticket-type',
            'dash-filter-priority', 'dash-filter-assignee', 'dash-filter-creator', 'dash-filter-process'
        ];
        dashFilterIds.forEach(id => {
            const el = document.getElementById(id);
            if (el) el.addEventListener('change', applyDashFilters);
        });
        const dashSearch = document.getElementById('dash-filter-search');
        if (dashSearch) dashSearch.addEventListener('input', applyDashFilters);
        const dashClear = document.getElementById('dash-clear-filters');
        if (dashClear) {
            dashClear.addEventListener('click', () => {
                dashFilterIds.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
                const s = document.getElementById('dash-filter-search'); if (s) s.value = '';
                applyDashFilters();
            });
        }

        // Set up History action listeners
        const historySearchInput = document.getElementById('history-search-input');
        if (historySearchInput) {
            historySearchInput.addEventListener('input', filterAndRenderHistoryJobs);
        }

        const historyRefreshBtn = document.getElementById('history-refresh-btn');
        const historyRefreshIcon = document.getElementById('history-refresh-icon');
        if (historyRefreshBtn) {
            historyRefreshBtn.addEventListener('click', async () => {
                historyRefreshBtn.disabled = true;
                if (historyRefreshIcon) historyRefreshIcon.style.animation = 'spin 1.5s linear infinite';
                
                try {
                    await loadHistoryJobs();
                } catch (err) {
                    alert('Erro ao atualizar histórico: ' + err.message);
                } finally {
                    historyRefreshBtn.disabled = false;
                    if (historyRefreshIcon) historyRefreshIcon.style.animation = '';
                }
            });
        }
    });

    // ─── JIRA Dashboard ────────────────────────────────────────────────────────

    function populateDashFilter(selectId, tickets, field, placeholder) {
        const sel = document.getElementById(selectId);
        if (!sel) return;
        const current = sel.value;
        
        let filteredTickets = tickets;
        if (selectId === 'dash-filter-assignee') {
            const activeTeamFilter = document.getElementById('dash-filter-team')?.value || '';
            if (activeTeamFilter) {
                filteredTickets = tickets.filter(t => (t.team || '').trim() === activeTeamFilter);
            }
        }
        
        const valuesSet = new Set(filteredTickets.map(t => (t[field] || '').trim()).filter(v => v && String(v).trim() !== ""));
        
        if (selectId === 'dash-filter-assignee') {
            const activeTeamFilter = document.getElementById('dash-filter-team')?.value || '';
            if (activeTeamFilter && TEAM_MEMBERS[activeTeamFilter]) {
                TEAM_MEMBERS[activeTeamFilter].forEach(member => valuesSet.add(member));
            }
        }
        
        const values = Array.from(valuesSet).sort((a, b) => a.localeCompare(b));
        sel.innerHTML = `<option value="">${placeholder}</option>` +
            values.map(v => `<option value="${escapeHtml(v)}"${v === current ? ' selected' : ''}>${escapeHtml(v)}</option>`).join('');
    }

    function applyDashFilters() {
        const val = id => (document.getElementById(id)?.value || '').trim();
        const query = (document.getElementById('dash-filter-search')?.value || '').toLowerCase().trim();

        const filtered = jiraTickets.filter(t => {
            if (val('dash-filter-team') && (t.team || '').trim() !== val('dash-filter-team')) return false;
            if (val('dash-filter-stream') && (t.stream || '').trim() !== val('dash-filter-stream')) return false;
            if (val('dash-filter-ticket-type') && (t.ticket_type || '').trim() !== val('dash-filter-ticket-type')) return false;
            if (val('dash-filter-priority') && (t.priority || '').trim() !== val('dash-filter-priority')) return false;
            if (val('dash-filter-assignee') && (t.assignee || '').trim() !== val('dash-filter-assignee')) return false;
            if (val('dash-filter-creator') && (t.creator || '').trim() !== val('dash-filter-creator')) return false;
            if (val('dash-filter-process') && (t.process || '').trim() !== val('dash-filter-process')) return false;
            if (query && !((t.summary || '') + (t.key || '')).toLowerCase().includes(query)) return false;
            return true;
        });

        // Update count badge
        const countEl = document.getElementById('dash-filter-count');
        if (countEl) {
            const active = ['dash-filter-team','dash-filter-stream','dash-filter-ticket-type',
                'dash-filter-priority','dash-filter-assignee','dash-filter-creator','dash-filter-process']
                .filter(id => val(id)).length + (query ? 1 : 0);
            countEl.textContent = active > 0
                ? `${filtered.length} de ${jiraTickets.length} tickets`
                : `${jiraTickets.length} tickets`;
            countEl.style.color = active > 0 ? 'var(--primary)' : 'var(--text-secondary)';
        }

        renderJiraDashboard(filtered);
    }

    function renderJiraDashboard(tickets) {
        if (!tickets) tickets = jiraTickets;

        // ── Populate filter dropdowns from full data first ──
        populateDashFilter('dash-filter-team', jiraTickets, 'team', 'Todas as equipas');
        populateDashFilter('dash-filter-stream', jiraTickets, 'stream', 'Todos os streams');
        populateDashFilter('dash-filter-ticket-type', jiraTickets, 'ticket_type', 'Todos os tipos');
        populateDashFilter('dash-filter-priority', jiraTickets, 'priority', 'Todas as prioridades');
        populateDashFilter('dash-filter-assignee', jiraTickets, 'assignee', 'Todos os responsáveis');
        populateDashFilter('dash-filter-creator', jiraTickets, 'creator', 'Todos os criadores');
        populateDashFilter('dash-filter-process', jiraTickets, 'process', 'Todos os processos');

        // Update count on first render
        const dashCountEl = document.getElementById('dash-filter-count');
        if (dashCountEl && !dashCountEl.textContent) dashCountEl.textContent = `${tickets.length} tickets`;

        // ── helpers ──
        function count(arr, key, val) {
            return arr.filter(t => (t[key] || '').trim() === val).length;
        }
        function countByField(arr, field) {
            const map = {};
            arr.forEach(t => {
                const v = (t[field] || 'Desconhecido').trim();
                map[v] = (map[v] || 0) + 1;
            });
            return map;
        }
        function sortedEntries(map, limit) {
            return Object.entries(map).sort((a,b) => b[1]-a[1]).slice(0, limit || 999);
        }
        const isDark = document.documentElement.classList.contains('dark') ||
                       document.body.dataset.theme === 'dark' ||
                       getComputedStyle(document.documentElement).getPropertyValue('--bg-main').includes('1');
        const textColor = getComputedStyle(document.documentElement).getPropertyValue('--text-secondary').trim() || '#9ca3af';

        // ── PALETTE ──
        const PALETTE = [
            '#6366f1','#3b82f6','#10b981','#f59e0b','#ef4444',
            '#8b5cf6','#ec4899','#14b8a6','#f97316','#06b6d4'
        ];

        // ── DONUT CHART ──
        function drawDonut(canvasId, data, legendId, centerTotalId) {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            const ctx = canvas.getContext('2d');
            const W = canvas.width, H = canvas.height;
            ctx.clearRect(0, 0, W, H);
            const total = data.reduce((s, d) => s + d.value, 0);
            if (centerTotalId) {
                const el = document.getElementById(centerTotalId);
                if (el) el.textContent = total;
            }
            if (total === 0) {
                ctx.beginPath();
                ctx.arc(W/2, H/2, H/2 - 10, 0, Math.PI*2);
                ctx.strokeStyle = '#374151';
                ctx.lineWidth = 22;
                ctx.stroke();
                return;
            }
            let angle = -Math.PI / 2;
            const r = H/2 - 10;
            const inner = r * 0.58;
            data.forEach((d, i) => {
                const slice = (d.value / total) * Math.PI * 2;
                ctx.beginPath();
                ctx.moveTo(W/2, H/2);
                ctx.arc(W/2, H/2, r, angle, angle + slice);
                ctx.closePath();
                ctx.fillStyle = d.color;
                ctx.fill();
                angle += slice;
            });
            // Inner circle
            ctx.beginPath();
            ctx.arc(W/2, H/2, inner, 0, Math.PI*2);
            ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--card-bg') || '#1a1f2e';
            ctx.fill();
            // Legend
            if (legendId) {
                const leg = document.getElementById(legendId);
                if (leg) {
                    leg.innerHTML = data.map(d => `
                        <div style="display:flex;align-items:center;gap:7px;justify-content:space-between;">
                          <div style="display:flex;align-items:center;gap:6px;">
                            <span style="width:9px;height:9px;border-radius:50%;background:${d.color};flex-shrink:0;"></span>
                            <span style="color:var(--text-primary);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:130px;">${escapeHtml(d.label)}</span>
                          </div>
                          <span style="font-weight:700;color:var(--text-primary);">${d.value}</span>
                        </div>`).join('');
                }
            }
        }

        // ── BAR CHART ──
        function drawBar(canvasId, labels, values, colors) {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            // Set canvas pixel size from layout size
            const parent = canvas.parentElement;
            canvas.width = parent ? parent.clientWidth || 260 : 260;
            canvas.height = 180;
            const ctx = canvas.getContext('2d');
            const W = canvas.width, H = canvas.height;
            ctx.clearRect(0, 0, W, H);
            const max = Math.max(...values, 1);
            const pad = { top: 10, right: 10, bottom: 38, left: 28 };
            const chartW = W - pad.left - pad.right;
            const chartH = H - pad.top - pad.bottom;
            const barW = Math.max(4, (chartW / labels.length) - 6);

            // Grid lines
            ctx.strokeStyle = 'rgba(255,255,255,0.06)';
            ctx.lineWidth = 1;
            [0.25, 0.5, 0.75, 1].forEach(f => {
                const y = pad.top + chartH * (1 - f);
                ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(W - pad.right, y); ctx.stroke();
                ctx.fillStyle = textColor;
                ctx.font = '9px Inter,system-ui,sans-serif';
                ctx.textAlign = 'right';
                ctx.fillText(Math.round(max * f), pad.left - 4, y + 3);
            });

            labels.forEach((lbl, i) => {
                const x = pad.left + i * (chartW / labels.length) + (chartW / labels.length - barW) / 2;
                const barH = (values[i] / max) * chartH;
                const y = pad.top + chartH - barH;
                // Bar
                const grad = ctx.createLinearGradient(x, y, x, y + barH);
                const col = colors[i % colors.length];
                grad.addColorStop(0, col);
                grad.addColorStop(1, col + '55');
                ctx.fillStyle = grad;
                const radius = Math.min(5, barW / 2);
                ctx.beginPath();
                ctx.moveTo(x + radius, y);
                ctx.lineTo(x + barW - radius, y);
                ctx.quadraticCurveTo(x + barW, y, x + barW, y + radius);
                ctx.lineTo(x + barW, y + barH);
                ctx.lineTo(x, y + barH);
                ctx.lineTo(x, y + radius);
                ctx.quadraticCurveTo(x, y, x + radius, y);
                ctx.closePath();
                ctx.fill();
                // Value on top
                ctx.fillStyle = textColor;
                ctx.font = 'bold 10px Inter,system-ui,sans-serif';
                ctx.textAlign = 'center';
                if (values[i] > 0) ctx.fillText(values[i], x + barW / 2, y - 3);
                // Label
                ctx.fillStyle = textColor;
                ctx.font = '9px Inter,system-ui,sans-serif';
                ctx.textAlign = 'center';
                const truncLbl = lbl.length > 8 ? lbl.slice(0, 7) + '…' : lbl;
                ctx.fillText(truncLbl, x + barW / 2, H - pad.bottom + 14);
            });
        }

        // ── HORIZONTAL BAR CHART ──
        function drawHorizontalBar(canvasId, labels, values, colors) {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            const parent = canvas.parentElement;
            canvas.width = parent ? parent.clientWidth || 260 : 260;
            canvas.height = 180;
            const ctx = canvas.getContext('2d');
            const W = canvas.width, H = canvas.height;
            ctx.clearRect(0, 0, W, H);

            const max = Math.max(...values, 1);
            const pad = { top: 15, right: 30, bottom: 10, left: 40 };
            const chartW = W - pad.left - pad.right;
            const chartH = H - pad.top - pad.bottom;
            const barH = Math.max(4, (chartH / labels.length) - 8);

            // Vertical grid lines
            ctx.strokeStyle = 'rgba(255,255,255,0.06)';
            ctx.lineWidth = 1;
            [0.25, 0.5, 0.75, 1].forEach(f => {
                const x = pad.left + chartW * f;
                ctx.beginPath(); ctx.moveTo(x, pad.top); ctx.lineTo(x, H - pad.bottom); ctx.stroke();
            });

            labels.forEach((lbl, i) => {
                const y = pad.top + i * (chartH / labels.length) + (chartH / labels.length - barH) / 2;
                const barW = (values[i] / max) * chartW;
                const x = pad.left;

                // Gradient
                const grad = ctx.createLinearGradient(x, y, x + barW, y);
                const col = colors[i % colors.length];
                grad.addColorStop(0, col + '55');
                grad.addColorStop(1, col);
                ctx.fillStyle = grad;

                // Rounded bar
                const radius = Math.min(5, barH / 2);
                ctx.beginPath();
                ctx.moveTo(x, y);
                ctx.lineTo(x + barW - radius, y);
                ctx.quadraticCurveTo(x + barW, y, x + barW, y + radius);
                ctx.lineTo(x + barW, y + barH - radius);
                ctx.quadraticCurveTo(x + barW, y + barH, x + barW - radius, y + barH);
                ctx.lineTo(x, y + barH);
                ctx.closePath();
                ctx.fill();

                // Value next to the bar
                ctx.fillStyle = textColor;
                ctx.font = 'bold 10px Inter,system-ui,sans-serif';
                ctx.textAlign = 'left';
                if (values[i] > 0) {
                    ctx.fillText(values[i], x + barW + 5, y + barH / 2 + 3);
                }

                // Label on the left
                ctx.fillStyle = textColor;
                ctx.font = '9px Inter,system-ui,sans-serif';
                ctx.textAlign = 'right';
                ctx.fillText(lbl, x - 6, y + barH / 2 + 3);
            });
        }


        // ─── 1. Status list (Somente tickets em aberto, igual imagem JIRA Live)
        const openTicketsForStatus = tickets.filter(t => {
            const s = (t.status || '').toLowerCase().trim();
            return !['done', 'closed', 'concluído', 'resolvido', 'fechado', 'fechada', 'cancelled'].includes(s);
        });

        const countStatus = (statusNameList) => {
            return openTicketsForStatus.filter(t => {
                const s = (t.status || '').toLowerCase().trim();
                return statusNameList.includes(s);
            }).length;
        };

        const inReviewCount = countStatus(['in review', 'sob revisão']);
        const wipCount = countStatus(['work in progress', 'em progresso', 'wip']);
        const reopenedCount = countStatus(['reopened', 'reaberto']);
        const waitingSupportCount = countStatus(['waiting for specialized support', 'aguarda suporte especializado', 'waiting for specialized support']);
        const waitingCustomerCount = countStatus(['waiting for customer', 'aguarda cliente']);
        const totalOpen = inReviewCount + wipCount + reopenedCount + waitingSupportCount + waitingCustomerCount;

        // Update UI elements
        const totalEl = document.getElementById('status-val-total');
        if (totalEl) totalEl.textContent = totalOpen;

        const inReviewEl = document.getElementById('status-val-in-review');
        if (inReviewEl) inReviewEl.textContent = inReviewCount;

        const wipEl = document.getElementById('status-val-wip');
        if (wipEl) wipEl.textContent = wipCount;

        const reopenedEl = document.getElementById('status-val-reopened');
        if (reopenedEl) reopenedEl.textContent = reopenedCount;

        const supportEl = document.getElementById('status-val-waiting-support');
        if (supportEl) supportEl.textContent = waitingSupportCount;

        const customerEl = document.getElementById('status-val-waiting-customer');
        if (customerEl) customerEl.textContent = waitingCustomerCount;

        // Update last refreshed time
        const refreshEl = document.getElementById('status-card-refresh-time');
        if (refreshEl) {
            const now = new Date();
            const pad = n => String(n).padStart(2, '0');
            refreshEl.textContent = `Last refreshed at ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
        }

        // ─── 1.5. User Status Table (Distribuição por Utilizador / Status) ───
        const userStatusGroups = {};
        openTicketsForStatus.forEach(t => {
            let assignee = (t.assignee || '').trim();
            if (!assignee || assignee.toLowerCase() === 'sem responsável') {
                assignee = 'Sem responsável';
            }
            if (!userStatusGroups[assignee]) {
                userStatusGroups[assignee] = {
                    rev: 0,
                    wip: 0,
                    reo: 0,
                    sup: 0,
                    cli: 0,
                    total: 0
                };
            }
            const s = (t.status || '').toLowerCase().trim();
            if (['in review', 'sob revisão'].includes(s)) {
                userStatusGroups[assignee].rev++;
            } else if (['work in progress', 'em progresso', 'wip'].includes(s)) {
                userStatusGroups[assignee].wip++;
            } else if (['reopened', 'reaberto'].includes(s)) {
                userStatusGroups[assignee].reo++;
            } else if (['waiting for specialized support', 'aguarda suporte especializado'].includes(s)) {
                userStatusGroups[assignee].sup++;
            } else if (['waiting for customer', 'aguarda cliente'].includes(s)) {
                userStatusGroups[assignee].cli++;
            }
            userStatusGroups[assignee].total++;
        });

        const sortedStatusAssignees = Object.keys(userStatusGroups).sort((a, b) => {
            if (a === 'Sem responsável') return 1;
            if (b === 'Sem responsável') return -1;
            return a.localeCompare(b);
        });

        let totalRev = 0, totalWip = 0, totalReo = 0, totalSup = 0, totalCli = 0, grandStatusTotal = 0;

        sortedStatusAssignees.forEach(assignee => {
            const data = userStatusGroups[assignee];
            totalRev += data.rev;
            totalWip += data.wip;
            totalReo += data.reo;
            totalSup += data.sup;
            totalCli += data.cli;
            grandStatusTotal += data.total;
        });

        const statusTbodyEl = document.getElementById('user-status-table-body');
        if (statusTbodyEl) {
            if (sortedStatusAssignees.length === 0) {
                statusTbodyEl.innerHTML = `<tr><td colspan="7" style="padding: 16px; text-align: center; color: var(--text-secondary);">Sem dados de pendências por utilizador.</td></tr>`;
            } else {
                statusTbodyEl.innerHTML = sortedStatusAssignees.map(assignee => {
                    const data = userStatusGroups[assignee];
                    return `
                        <tr style="border-bottom: 1px solid var(--border-color); height: 38px;">
                          <td style="text-align: left; padding: 10px 14px; font-weight: 500; color: var(--text-primary); white-space: nowrap;">${escapeHtml(assignee)}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.rev > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.rev}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.wip > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.wip}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.reo > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.reo}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.sup > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.sup}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.cli > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.cli}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 800; color: var(--text-primary);">${data.total}</td>
                        </tr>
                    `;
                }).join('');
            }
        }

        const footerRev = document.getElementById('user-status-total-rev');
        if (footerRev) footerRev.textContent = totalRev;
        const footerWip = document.getElementById('user-status-total-wip');
        if (footerWip) footerWip.textContent = totalWip;
        const footerReo = document.getElementById('user-status-total-reo');
        if (footerReo) footerReo.textContent = totalReo;
        const footerSup = document.getElementById('user-status-total-sup');
        if (footerSup) footerSup.textContent = totalSup;
        const footerCli = document.getElementById('user-status-total-cli');
        if (footerCli) footerCli.textContent = totalCli;
        const footerGrand = document.getElementById('user-status-total-grand');
        if (footerGrand) footerGrand.textContent = grandStatusTotal;

        const countStatsEl = document.getElementById('user-status-count-stats');
        if (countStatsEl) {
            countStatsEl.textContent = `Mostrando ${sortedStatusAssignees.length} de ${sortedStatusAssignees.length} estatísticas.`;
        }

        const userRefreshEl = document.getElementById('user-status-refresh-time');
        if (userRefreshEl) {
            const now = new Date();
            const pad = n => String(n).padStart(2, '0');
            userRefreshEl.textContent = `Last refreshed at ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
        }


        // ─── 1.6. User Ticket Type Table (Distribuição por Utilizador / Tipo de Ticket) ───
        const userTypeGroups = {};
        openTicketsForStatus.forEach(t => {
            let assignee = (t.assignee || '').trim();
            if (!assignee || assignee.toLowerCase() === 'sem responsável') {
                assignee = 'Sem responsável';
            }
            if (!userTypeGroups[assignee]) {
                userTypeGroups[assignee] = {
                    sr: 0,
                    inc: 0,
                    prj: 0,
                    total: 0
                };
            }
            const type = (t.ticket_type || '').toLowerCase().trim();
            if (type.includes('incident')) {
                userTypeGroups[assignee].inc++;
            } else if (type.includes('request') || type.includes('service')) {
                userTypeGroups[assignee].sr++;
            } else {
                userTypeGroups[assignee].prj++;
            }
            userTypeGroups[assignee].total++;
        });

        const sortedTypeAssignees = Object.keys(userTypeGroups).sort((a, b) => {
            if (a === 'Sem responsável') return 1;
            if (b === 'Sem responsável') return -1;
            return a.localeCompare(b);
        });

        let totalTypeSr = 0, totalTypeInc = 0, totalTypePrj = 0, grandTypeTotal = 0;

        sortedTypeAssignees.forEach(assignee => {
            const data = userTypeGroups[assignee];
            totalTypeSr += data.sr;
            totalTypeInc += data.inc;
            totalTypePrj += data.prj;
            grandTypeTotal += data.total;
        });

        const typeTbodyEl = document.getElementById('user-type-table-body');
        if (typeTbodyEl) {
            if (sortedTypeAssignees.length === 0) {
                typeTbodyEl.innerHTML = `<tr><td colspan="5" style="padding: 16px; text-align: center; color: var(--text-secondary);">Sem dados de pendências por utilizador.</td></tr>`;
            } else {
                typeTbodyEl.innerHTML = sortedTypeAssignees.map(assignee => {
                    const data = userTypeGroups[assignee];
                    return `
                        <tr style="border-bottom: 1px solid var(--border-color); height: 38px;">
                          <td style="text-align: left; padding: 10px 14px; font-weight: 500; color: var(--text-primary); white-space: nowrap;">${escapeHtml(assignee)}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.sr > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.sr}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.inc > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.inc}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 600; color: ${data.prj > 0 ? 'var(--text-primary)' : 'var(--text-secondary)'};">${data.prj}</td>
                          <td style="text-align: center; padding: 10px 14px; font-weight: 800; color: var(--text-primary);">${data.total}</td>
                        </tr>
                    `;
                }).join('');
            }
        }

        const footerSr = document.getElementById('user-type-total-sr');
        if (footerSr) footerSr.textContent = totalTypeSr;
        const footerInc = document.getElementById('user-type-total-inc');
        if (footerInc) footerInc.textContent = totalTypeInc;
        const footerPrj = document.getElementById('user-type-total-prj');
        if (footerPrj) footerPrj.textContent = totalTypePrj;
        const footerTypeGrand = document.getElementById('user-type-total-grand');
        if (footerTypeGrand) footerTypeGrand.textContent = grandTypeTotal;

        const countTypeStatsEl = document.getElementById('user-type-count-stats');
        if (countTypeStatsEl) {
            countTypeStatsEl.textContent = `Mostrando ${sortedTypeAssignees.length} de ${sortedTypeAssignees.length} estatísticas.`;
        }

        const typeRefreshEl = document.getElementById('user-type-refresh-time');
        if (typeRefreshEl) {
            const now = new Date();
            const pad = n => String(n).padStart(2, '0');
            typeRefreshEl.textContent = `Last refreshed at ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
        }


        // ─── 2. Priority bar ───
        const PRIORITY_ORDER = ['Critical','Highest','High','Medium','Low','Lowest'];
        const prioMap = countByField(openTicketsForStatus, 'priority');
        const prioColors = { Critical:'#ef4444', Highest:'#f97316', High:'#f59e0b', Medium:'#3b82f6', Low:'#10b981', Lowest:'#6366f1' };
        const prioLabels = PRIORITY_ORDER.filter(p => prioMap[p] > 0);
        const prioValues = prioLabels.map(p => prioMap[p]);
        const prioColArr = prioLabels.map(p => prioColors[p] || '#6366f1');
        // Also include unknown priorities
        Object.keys(prioMap).forEach(p => { if (!PRIORITY_ORDER.includes(p)) { prioLabels.push(p); prioValues.push(prioMap[p]); prioColArr.push('#8b5cf6'); } });
        drawBar('dash-chart-priority', prioLabels.length ? prioLabels : ['Sem dados'], prioValues.length ? prioValues : [0], prioColArr.length ? prioColArr : ['#6366f1']);

        // ─── 4. Process Table (Replacing Team bar) ───
        const processMap = {};
        let totalProcessTickets = 0;
        const currentYear = new Date().getFullYear();
        tickets.forEach(t => {
            const s = (t.status || '').toLowerCase().trim();
            const isOpen = !['done', 'closed', 'concluído', 'resolvido', 'fechado', 'fechada', 'cancelled'].includes(s);
            let isCurrentYear = isOpen;
            if (!isOpen && t.resolved_at) {
                const resDate = new Date(t.resolved_at);
                if (!isNaN(resDate.getTime()) && resDate.getFullYear() === currentYear) {
                    isCurrentYear = true;
                }
            }
            if (!isCurrentYear) return;

            const p = (t.process && t.process.trim()) ? t.process.trim() : 'Sem processo';
            processMap[p] = (processMap[p] || 0) + 1;
            totalProcessTickets++;
        });

        const sortedProcesses = Object.entries(processMap)
            .sort((a, b) => b[1] - a[1]);

        const processTbodyEl = document.getElementById('user-process-table-body');
        if (processTbodyEl) {
            if (sortedProcesses.length === 0) {
                processTbodyEl.innerHTML = `<tr><td colspan="3" style="padding: 16px; text-align: center; color: var(--text-secondary);">Sem dados de processos.</td></tr>`;
            } else {
                processTbodyEl.innerHTML = sortedProcesses.map(([proc, count]) => {
                    const pct = totalProcessTickets > 0 ? ((count / totalProcessTickets) * 100).toFixed(1) : '0.0';
                    return `
                        <tr style="border-bottom: 1px solid var(--border-color); height: 32px;">
                          <td style="text-align: left; padding: 6px 8px; font-weight: 500; color: var(--text-primary); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 140px;" title="${escapeHtml(proc)}">${escapeHtml(proc)}</td>
                          <td style="text-align: center; padding: 6px 8px; font-weight: 700; color: var(--text-primary);">${count}</td>
                          <td style="text-align: center; padding: 6px 8px; font-weight: 600; color: var(--text-secondary);">${pct}%</td>
                        </tr>
                    `;
                }).join('');
            }
        }

        // ─── 5. Tickets por Ano bar chart (Replacing SLA donut) ───
        const yearMap = {
            2023: 12815,
            2024: 7929,
            2025: 5691
        };
        tickets.forEach(t => {
            if (!t.resolved_at) return;
            const resDate = new Date(t.resolved_at);
            if (isNaN(resDate.getTime())) return;
            const year = resDate.getFullYear();
            if (year >= 2026) {
                yearMap[year] = (yearMap[year] || 0) + 1;
            }
        });

        const sortedYears = Object.keys(yearMap).sort((a, b) => parseInt(a) - parseInt(b));
        const yearLabels = sortedYears;
        const yearValues = sortedYears.map(y => yearMap[y]);

        drawHorizontalBar('dash-chart-sla', yearLabels.length ? yearLabels : ['Sem dados'], yearValues.length ? yearValues : [0], ['#10b981']);

        // ─── 6. Top assignees bar ───
        const assigneeMap = {};
        openTicketsForStatus.forEach(t => {
            const a = (t.assignee && t.assignee.trim() && t.assignee.toLowerCase() !== 'sem responsável') ? t.assignee.trim() : null;
            if (a) assigneeMap[a] = (assigneeMap[a] || 0) + 1;
        });
        const assigneeEntries = sortedEntries(assigneeMap, 8);
        drawBar('dash-chart-assignee', assigneeEntries.map(e=>e[0]), assigneeEntries.map(e=>e[1]), PALETTE);
 
        // ─── 6.5. Time to Resolution (TTR) ───
        const EXCLUDED_TTR_TICKETS = new Set(['IZ-52956']);
        const ttrSums = Array(12).fill(0);
        const ttrCounts = Array(12).fill(0);
        let averageSum = 0;
        let averageCount = 0;
        let overallCount = 0;
        let slaOkCount = 0;
        let slaBreachedCount = 0;

        tickets.forEach(t => {
            if (t.key && EXCLUDED_TTR_TICKETS.has(t.key)) return;
            if (!t.created_at || !t.resolved_at) return;
            const createdDate = new Date(t.created_at);
            const resolvedDate = new Date(t.resolved_at);
            if (isNaN(createdDate.getTime()) || isNaN(resolvedDate.getTime())) return;

            if (resolvedDate.getFullYear() === currentYear) {
                const isProject = t.ticket_type && t.ticket_type.trim().toLowerCase() === 'project';
                
                overallCount++;

                if (isProject) {
                    slaOkCount++;
                } else {
                    if (t.time_to_resolution) {
                        const slaLower = t.time_to_resolution.toLowerCase();
                        if (slaLower.includes('atraso') || slaLower.includes('excedido')) {
                            slaBreachedCount++;
                        } else if (slaLower.includes('resolvido')) {
                            slaOkCount++;
                        }
                    }
                }

                if (!isProject) {
                    const ttrMs = resolvedDate - createdDate;
                    const ttrDays = Math.max(0, ttrMs / (1000 * 60 * 60 * 24)); // ttr in days

                    const month = resolvedDate.getMonth(); // 0-11
                    ttrSums[month] += ttrDays;
                    ttrCounts[month]++;

                    averageSum += ttrDays;
                    averageCount++;
                }
            }
        });

        const MONTH_LABELS = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        const ttrAverages = MONTH_LABELS.map((_, i) => {
            return ttrCounts[i] > 0 ? parseFloat((ttrSums[i] / ttrCounts[i]).toFixed(1)) : 0;
        });

        // Draw Monthly TTR Bar Chart
        drawBar('dash-chart-ttr', MONTH_LABELS, ttrAverages, ['#3b82f6']);

        // Update TTR Statistics Labels
        const avgEl = document.getElementById('ttr-val-average');
        if (avgEl) {
            avgEl.textContent = averageCount > 0 ? `${(averageSum / averageCount).toFixed(1)} dias` : '—';
        }
        const countEl = document.getElementById('ttr-val-resolved-count');
        if (countEl) {
            countEl.textContent = overallCount;
        }

        // SLA percentages
        const totalSlaResolved = slaOkCount + slaBreachedCount;
        const slaOkPercent = totalSlaResolved > 0 ? ((slaOkCount / totalSlaResolved) * 100).toFixed(1) : '0.0';
        const slaBreachedPercent = totalSlaResolved > 0 ? ((slaBreachedCount / totalSlaResolved) * 100).toFixed(1) : '0.0';

        const slaOkEl = document.getElementById('ttr-val-sla-ok');
        if (slaOkEl) {
            slaOkEl.textContent = `${slaOkPercent}% (${slaOkCount})`;
        }
        const slaBreachedEl = document.getElementById('ttr-val-sla-breached');
        if (slaBreachedEl) {
            slaBreachedEl.textContent = `${slaBreachedPercent}% (${slaBreachedCount})`;
        }

        // ─── 7. Backlog crítico ───
        const critical = tickets.filter(t => {
            const slaEx = t.time_to_resolution && t.time_to_resolution.toLowerCase().includes('excedido');
            const noAssignee = !t.assignee || t.assignee.trim() === '' || t.assignee.toLowerCase() === 'sem responsável';
            return slaEx || noAssignee;
        }).slice(0, 20);

        const backlogCountEl = document.getElementById('dash-backlog-count');
        if (backlogCountEl) backlogCountEl.textContent = critical.length + ' ticket' + (critical.length !== 1 ? 's' : '');

        const tbody = document.getElementById('dash-backlog-body');
        if (tbody) {
            if (critical.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" style="padding:20px;text-align:center;color:var(--text-secondary);">Sem tickets críticos ✅</td></tr>`;
            } else {
                const PRIORITY_COLORS = { Critical:'#ef4444', Highest:'#f97316', High:'#f59e0b', Medium:'#3b82f6', Low:'#10b981', Lowest:'#6366f1' };
                tbody.innerHTML = critical.map(t => {
                    const slaEx = t.time_to_resolution && t.time_to_resolution.toLowerCase().includes('excedido');
                    const noAssignee = !t.assignee || t.assignee.trim() === '' || t.assignee.toLowerCase() === 'sem responsável';
                    const prioColor = PRIORITY_COLORS[t.priority] || '#9ca3af';
                    const rowBg = slaEx ? 'background: rgba(239,68,68,0.04);' : 'background: rgba(245,158,11,0.04);';
                    return `<tr style="border-bottom: 1px solid var(--border-color); ${rowBg}">
                        <td style="padding:10px 12px;font-family:monospace;font-weight:bold;color:var(--primary);white-space:nowrap;">${escapeHtml(t.key || '-')}</td>
                        <td style="padding:10px 12px;color:var(--text-primary);max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${escapeHtml(t.summary || '')}">${escapeHtml(t.summary || '-')}</td>
                        <td style="padding:10px 12px;"><span style="font-size:11px;padding:3px 8px;border-radius:5px;background:rgba(99,102,241,0.12);color:#6366f1;font-weight:600;">${escapeHtml(t.status || '-')}</span></td>
                        <td style="padding:10px 12px;"><span style="font-weight:700;color:${prioColor};">${escapeHtml(t.priority || '-')}</span></td>
                        <td style="padding:10px 12px;color:${noAssignee ? '#f59e0b' : 'var(--text-secondary)'};font-weight:${noAssignee ? '600' : '400'}">${noAssignee ? '⚠️ Sem responsável' : escapeHtml(t.assignee)}</td>
                        <td style="padding:10px 12px;"><span style="font-size:11px;padding:3px 8px;border-radius:5px;font-weight:600;${slaEx ? 'background:rgba(239,68,68,0.12);color:#ef4444;' : 'color:var(--text-secondary);'}">${escapeHtml(t.time_to_resolution || '-')}</span></td>
                    </tr>`;
                }).join('');
            }
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  DEFINIÇÕES – CRUD de Regras de Contexto do Agente SAP
    // ═══════════════════════════════════════════════════════════════

    const DEF_CAMPOS = [
        'IT SALSA - Categoria SAP',
        'Tipo de Ticket',
        'Stream',
    ];

    const DEF_CAMPO_LABELS = {
        'IT SALSA - Categoria SAP': { icon: '🏷️', color: '#6366f1', bg: 'rgba(99,102,241,0.1)' },
        'Tipo de Ticket':           { icon: '📋', color: '#f59e0b', bg: 'rgba(245,158,11,0.1)' },
        'Stream':                   { icon: '🔀', color: '#3b82f6', bg: 'rgba(59,130,246,0.1)'  },
    };

    let defEditingRuleId = null;

    async function defLoadRules() {
        const tbody = document.getElementById('def-rules-tbody');
        if (!tbody) return;
        tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:32px; color:#94a3b8; font-size:0.8rem;">A carregar regras...</td></tr>`;
        try {
            const res = await fetch('/api/agent/rules', { cache: 'no-store' });
            if (!res.ok) throw new Error('Falha ao carregar regras');
            const data = await res.json();
            defRenderRules(data.rules || []);
        } catch (e) {
            tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:32px; color:#ef4444; font-size:0.8rem;">&#x26A0; Erro ao carregar regras: ${e.message}</td></tr>`;
        }
    }

    function defRenderRules(rules) {
        const tbody = document.getElementById('def-rules-tbody');
        if (!tbody) return;
        if (rules.length === 0) {
            tbody.innerHTML = `
                <tr>
                  <td colspan="9" style="text-align:center; padding:48px 24px;">
                    <div style="font-size:2rem; margin-bottom:12px;">📭</div>
                    <div style="font-size:0.85rem; font-weight:700; color:#475569; margin-bottom:6px;">Sem regras configuradas</div>
                    <div style="font-size:0.75rem; color:#94a3b8;">Clique em <strong>Nova Regra</strong> para adicionar a primeira regra de contexto.</div>
                  </td>
                </tr>`;
            return;
        }

        tbody.innerHTML = rules.map(rule => {
            const cl = DEF_CAMPO_LABELS[rule.campo] || { icon: '📌', color: '#64748b', bg: 'rgba(100,116,139,0.1)' };
            const tagsHtml = rule.tags
                ? rule.tags.split(',').map(t => t.trim()).filter(Boolean)
                    .map(t => `<span style="display:inline-block; font-size:0.62rem; font-weight:700; padding:1px 7px; border-radius:20px; background:rgba(245,158,11,0.1); color:#d97706; border:1px solid rgba(245,158,11,0.2); margin:1px 2px;">${escapeHtml(t)}</span>`)
                    .join('')
                : '<span style="color:#94a3b8; font-size:0.72rem;">—</span>';
            const transHtml = rule.transacao_sap
                ? `<span style="display:inline-flex; align-items:center; gap:4px; font-size:0.78rem; font-weight:800; padding:3px 10px; border-radius:8px; background:rgba(16,185,129,0.1); color:#059669; border:1px solid rgba(16,185,129,0.2); font-family:monospace; letter-spacing:0.05em;">&#x1F527; ${escapeHtml(rule.transacao_sap)}</span>`
                : '<span style="color:#94a3b8; font-size:0.72rem;">—</span>';
            const notasHtml = rule.notas
                ? `<span style="font-size:0.73rem; color:#475569; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;" title="${escapeHtml(rule.notas)}">${escapeHtml(rule.notas)}</span>`
                : '<span style="color:#94a3b8; font-size:0.72rem;">—</span>';

            return `<tr class="def-rule-row" data-id="${rule.id}" data-valor="${escapeHtml(rule.valor || '')}">
              <td style="font-size:0.8rem; max-width:180px; color:#475569;">${rule.nome_parametro ? escapeHtml(rule.nome_parametro) : '<span style="color:#94a3b8; font-size:0.72rem;">—</span>'}</td>
              <td>
                <span style="display:inline-flex; align-items:center; gap:5px; font-size:0.72rem; font-weight:700; padding:3px 10px; border-radius:20px; background:${cl.bg}; color:${cl.color}; border:1px solid ${cl.color}33; white-space:nowrap;">
                  ${cl.icon} ${escapeHtml(rule.campo || '')}
                </span>
              </td>
              <td style="font-size:0.8rem; font-weight:600; color:#334155; max-width:180px;">${escapeHtml(rule.valor || '')}</td>
              <td style="font-size:0.8rem; max-width:140px; color:#475569;">${rule.processo ? escapeHtml(rule.processo) : '<span style="color:#94a3b8; font-size:0.72rem;">—</span>'}</td>
              <td style="font-size:0.8rem; max-width:160px; color:#475569;">${rule.subprocesso ? escapeHtml(rule.subprocesso) : '<span style="color:#94a3b8; font-size:0.72rem;">—</span>'}</td>
              <td>${transHtml}</td>
              <td style="max-width:160px;">${tagsHtml}</td>
              <td style="max-width:260px;">${notasHtml}</td>
              <td style="text-align:center; white-space:nowrap;">
                <button onclick="defEditRule('${rule.id}')" style="display:inline-flex; align-items:center; gap:4px; padding:5px 10px; font-size:0.7rem; font-weight:700; color:#6366f1; background:rgba(99,102,241,0.08); border:1px solid rgba(99,102,241,0.2); border-radius:7px; cursor:pointer; transition:all 0.15s; margin-right:4px;" onmouseover="this.style.background='rgba(99,102,241,0.15)'" onmouseout="this.style.background='rgba(99,102,241,0.08)'">
                  ✏️ Editar
                </button>
                <button onclick="defDeleteRule('${rule.id}', this.closest('tr').dataset.valor)" style="display:inline-flex; align-items:center; gap:4px; padding:5px 10px; font-size:0.7rem; font-weight:700; color:#ef4444; background:rgba(239,68,68,0.06); border:1px solid rgba(239,68,68,0.15); border-radius:7px; cursor:pointer; transition:all 0.15s;" onmouseover="this.style.background='rgba(239,68,68,0.12)'" onmouseout="this.style.background='rgba(239,68,68,0.06)'">
                  🗑️ Apagar
                </button>
              </td>
            </tr>`;
        }).join('');
    }

    function defOpenModal(rule = null) {
        defEditingRuleId = rule ? rule.id : null;
        document.getElementById('def-modal-title').textContent = rule ? 'Editar Regra de Contexto' : 'Nova Regra de Contexto';
        document.getElementById('def-rule-id').value      = rule ? rule.id : '';
        document.getElementById('def-rule-campo').value          = rule ? rule.campo : 'IT SALSA - Categoria SAP';
        document.getElementById('def-rule-valor').value          = rule ? rule.valor : '';
        document.getElementById('def-rule-nome-parametro').value = rule ? rule.nome_parametro : '';
        document.getElementById('def-rule-processo').value       = rule ? rule.processo : '';
        document.getElementById('def-rule-subprocesso').value    = rule ? rule.subprocesso : '';
        document.getElementById('def-rule-transacao').value      = rule ? rule.transacao_sap : '';
        document.getElementById('def-rule-tags').value           = rule ? rule.tags : '';
        document.getElementById('def-rule-notas').value          = rule ? rule.notas : '';
        const modal = document.getElementById('def-rule-modal');
        if (modal) { modal.style.display = 'flex'; setTimeout(() => modal.querySelector('div').style.opacity = '1', 10); }
    }

    function defCloseModal() {
        const modal = document.getElementById('def-rule-modal');
        if (modal) modal.style.display = 'none';
        defEditingRuleId = null;
    }

    async function defSaveRule() {
        const campo          = (document.getElementById('def-rule-campo')?.value          || '').trim();
        const valor          = (document.getElementById('def-rule-valor')?.value          || '').trim();
        const nomeParametro  = (document.getElementById('def-rule-nome-parametro')?.value || '').trim();
        const processo       = (document.getElementById('def-rule-processo')?.value       || '').trim();
        const subprocesso    = (document.getElementById('def-rule-subprocesso')?.value    || '').trim();
        const transac        = (document.getElementById('def-rule-transacao')?.value      || '').trim();
        const tags           = (document.getElementById('def-rule-tags')?.value           || '').trim();
        const notas          = (document.getElementById('def-rule-notas')?.value          || '').trim();

        if (!nomeParametro || !campo || !valor) {
            showToast('Os campos "Nome do parâmetro", "Campo JIRA" e "Valor" são obrigatórios.', 'error');
            return;
        }

        const btn = document.getElementById('def-modal-save-btn');
        if (btn) { btn.disabled = true; btn.textContent = 'A guardar...'; }

        try {
            const isEdit = !!defEditingRuleId;
            const url    = isEdit ? `/api/agent/rules/${defEditingRuleId}` : '/api/agent/rules';
            const method = isEdit ? 'PUT' : 'POST';

            const res = await fetch(url, {
                method,
                cache: 'no-store',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    campo,
                    valor,
                    nome_parametro: nomeParametro,
                    processo,
                    subprocesso,
                    transacao_sap: transac,
                    notas,
                    tags,
                }),
            });
            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || 'Erro ao guardar regra');
            }
            showToast(isEdit ? 'Regra actualizada com sucesso!' : 'Regra criada com sucesso!', 'success');
            defCloseModal();
            await defLoadRules();
        } catch (e) {
            showToast('Erro: ' + e.message, 'error');
        } finally {
            if (btn) { btn.disabled = false; btn.innerHTML = '&#x1F4BE; Guardar Regra'; }
        }
    }

    async function defEditRule(ruleId) {
        try {
            const res = await fetch('/api/agent/rules', { cache: 'no-store' });
            const data = await res.json();
            const rule = (data.rules || []).find(r => r.id === ruleId);
            if (rule) defOpenModal(rule);
        } catch (e) {
            showToast('Erro ao carregar regra: ' + e.message, 'error');
        }
    }

    async function defDeleteRule(ruleId, valor) {
        if (!confirm(`Tem a certeza que deseja eliminar a regra "${valor}"?`)) return;
        try {
            const res = await fetch(`/api/agent/rules/${ruleId}`, { method: 'DELETE', cache: 'no-store' });
            if (!res.ok) throw new Error('Falha ao eliminar regra');
            showToast('Regra eliminada.', 'success');
            defLoadRules();
        } catch (e) {
            showToast('Erro: ' + e.message, 'error');
        }
    }

    // Close modal on backdrop click
    document.addEventListener('click', e => {
        const modal = document.getElementById('def-rule-modal');
        if (modal && e.target === modal) defCloseModal();
    });

    // ═══════════════════════════════════════════════════════════════
    //  SAP AGENT – Filter Sidebar Logic
    // ═══════════════════════════════════════════════════════════════

    async function saLoadTicketList() {
        const listEl    = document.getElementById('sa-ticket-list');
        const loadingEl = document.getElementById('sa-ticket-list-loading');

        try {
            // Reuse jiraTickets if already loaded, otherwise fetch
            let tickets = (typeof jiraTickets !== 'undefined' && jiraTickets.length > 0)
                ? jiraTickets
                : await (async () => {
                    const r = await fetch('/api/jira/tickets?limit=500&exclude_closed=false');
                    if (!r.ok) throw new Error('Falha ao carregar tickets');
                    const d = await r.json();
                    return d.tickets || [];
                })();

            saAllTickets = tickets;
            saPopulateFilters(tickets);
            saApplyFilters();
        } catch (e) {
            if (loadingEl) {
                loadingEl.innerHTML = '<div style="color:#ef4444; font-size:0.75rem;">&#x26A0; Erro ao carregar tickets</div>';
            }
        }
    }

    function saPopulateFilters(tickets) {
        const statusSet   = new Set();
        const typeSet     = new Set();
        const streamSet   = new Set();
        const assigneeSet = new Set();

        tickets.forEach(t => {
            if (t.status)   statusSet.add(t.status);
            if (t.ticket_type) typeSet.add(t.ticket_type);
            if (t.stream)   streamSet.add(t.stream);
            if (t.assignee) assigneeSet.add(t.assignee);
        });

        function fillSelect(id, values, placeholder) {
            const el = document.getElementById(id);
            if (!el) return;
            const current = el.value;
            el.innerHTML = `<option value="">${placeholder}</option>`;
            [...values].sort().forEach(v => {
                const opt = document.createElement('option');
                opt.value = v;
                opt.textContent = v;
                if (v === current) opt.selected = true;
                el.appendChild(opt);
            });
        }

        fillSelect('sa-filter-status',   statusSet,   'Todos os status');
        fillSelect('sa-filter-type',     typeSet,     'Todos os tipos');
        fillSelect('sa-filter-stream',   streamSet,   'Todos os streams');
        fillSelect('sa-filter-assignee', assigneeSet, 'Todos os responsáveis');
    }

    function saApplyFilters() {
        const search   = (document.getElementById('sa-filter-search')?.value   || '').toLowerCase().trim();
        const status   = (document.getElementById('sa-filter-status')?.value   || '');
        const type     = (document.getElementById('sa-filter-type')?.value     || '');
        const stream   = (document.getElementById('sa-filter-stream')?.value   || '');
        const assignee = (document.getElementById('sa-filter-assignee')?.value || '');

        saFilteredTickets = saAllTickets.filter(t => {
            if (status   && t.status      !== status)   return false;
            if (type     && t.ticket_type !== type)     return false;
            if (stream   && t.stream      !== stream)   return false;
            if (assignee && t.assignee    !== assignee) return false;
            if (search) {
                const haystack = `${t.key || ''} ${t.summary || ''}`.toLowerCase();
                if (!haystack.includes(search)) return false;
            }
            return true;
        });

        saRenderTicketList(saFilteredTickets);

        const countEl = document.getElementById('sa-filter-count');
        if (countEl) {
            countEl.textContent = saFilteredTickets.length === saAllTickets.length
                ? `${saAllTickets.length} ticket${saAllTickets.length !== 1 ? 's' : ''}`
                : `${saFilteredTickets.length} de ${saAllTickets.length} tickets`;
        }
    }

    function saClearFilters() {
        ['sa-filter-search', 'sa-filter-status', 'sa-filter-type', 'sa-filter-stream', 'sa-filter-assignee'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = '';
        });
        saApplyFilters();
    }

    function saRenderTicketList(tickets) {
        const listEl    = document.getElementById('sa-ticket-list');
        const loadingEl = document.getElementById('sa-ticket-list-loading');
        if (!listEl) return;

        if (loadingEl) loadingEl.style.display = 'none';

        // Remove previously rendered items (keep loading placeholder)
        listEl.querySelectorAll('.sa-ticket-item').forEach(el => el.remove());

        if (tickets.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'sa-ticket-item';
            empty.style.cssText = 'padding: 20px; text-align: center; color: #94a3b8; font-size: 0.75rem;';
            empty.textContent = 'Nenhum ticket corresponde aos filtros.';
            listEl.appendChild(empty);
            return;
        }

        // Status colour map (same as JIRA tickets page)
        const statusColors = {
            'In Review':    { bg: 'rgba(59,130,246,0.1)',  color: '#3b82f6', border: 'rgba(59,130,246,0.25)' },
            'Work In Progress': { bg: 'rgba(245,158,11,0.1)', color: '#f59e0b', border: 'rgba(245,158,11,0.25)' },
            'Reopened':     { bg: 'rgba(239,68,68,0.1)',   color: '#ef4444', border: 'rgba(239,68,68,0.25)' },
            'Waiting for Customer': { bg: 'rgba(107,114,128,0.1)', color: '#6b7280', border: 'rgba(107,114,128,0.25)' },
            'Waiting for Specialized Support': { bg: 'rgba(139,92,246,0.1)', color: '#8b5cf6', border: 'rgba(139,92,246,0.25)' },
            'Done':         { bg: 'rgba(16,185,129,0.1)',  color: '#10b981', border: 'rgba(16,185,129,0.25)' },
        };

        tickets.forEach(ticket => {
            const key     = ticket.key     || '';
            const summary = ticket.summary || '';
            const status  = ticket.status  || '';
            const type    = ticket.ticket_type || '';
            const isActive = key === saActiveTicketKey;

            const sc = statusColors[status] || { bg: 'rgba(100,116,139,0.08)', color: '#64748b', border: 'rgba(100,116,139,0.2)' };
            const isIncident = (type || '').toLowerCase().includes('incident');

            const item = document.createElement('div');
            item.className = 'sa-ticket-item';
            item.dataset.key = key;
            item.style.cssText = [
                'padding: 10px 14px',
                'border-bottom: 1px solid #f1f5f9',
                'cursor: pointer',
                'transition: background 0.15s, transform 0.1s',
                isActive ? 'background: rgba(99,102,241,0.08); border-left: 3px solid #6366f1;' : 'border-left: 3px solid transparent;',
            ].join(';');

            item.innerHTML = `
                <div style="display:flex; align-items:center; justify-content:space-between; gap:6px; margin-bottom:4px;">
                    <span style="font-size:0.72rem; font-weight:800; color:#6366f1; letter-spacing:0.02em; white-space:nowrap;">${key}</span>
                    <span style="font-size:0.62rem; font-weight:700; padding:2px 7px; border-radius:20px; background:${sc.bg}; color:${sc.color}; border:1px solid ${sc.border}; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:100px;">${status}</span>
                </div>
                <div style="font-size:0.73rem; color:#475569; line-height:1.35; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;">${summary}</div>
                ${isIncident ? '<div style="margin-top:4px;"><span style="font-size:0.6rem; font-weight:700; padding:1px 6px; border-radius:10px; background:rgba(239,68,68,0.1); color:#ef4444; border:1px solid rgba(239,68,68,0.2);">INCIDENT</span></div>' : ''}
            `;

            item.addEventListener('mouseenter', () => {
                if (!isActive) item.style.background = 'rgba(99,102,241,0.04)';
            });
            item.addEventListener('mouseleave', () => {
                if (!isActive) item.style.background = '';
            });
            item.addEventListener('click', () => saSelectTicket(key));

            listEl.appendChild(item);
        });
    }

    function saSelectTicket(key) {
        // Update active highlight
        saActiveTicketKey = key;
        document.querySelectorAll('.sa-ticket-item').forEach(el => {
            const isThis = el.dataset.key === key;
            el.style.background    = isThis ? 'rgba(99,102,241,0.08)' : '';
            el.style.borderLeft    = isThis ? '3px solid #6366f1' : '3px solid transparent';
        });

        // Fill the ticket key input and trigger analysis
        const input = document.getElementById('sa-ticket-key');
        if (input) {
            input.value = key;
            // Smooth scroll to top of right column
            const resultCard = document.getElementById('sa-result-card');
            if (resultCard) resultCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }
        saRunAnalysis();
    }

    function saRenderPreview(ticketKey, details) {
        const card = document.getElementById('sa-result-card');
        if (!card) return;

        const sig = details.signal_preview || {};
        const matches = Array.isArray(details.context_matches) ? details.context_matches : [];
        const firstMatch = matches.length > 0 ? matches[0] : null;

        const categoriaVal = String(details.categoria_sap || '').trim();
        const categBadgeHtml = categoriaVal
            ? ` <span style="display:inline-flex;align-items:center;gap:4px;font-size:0.7rem;font-weight:700;padding:3px 9px;border-radius:20px;background:rgba(245,158,11,0.12);color:#f59e0b;border:1px solid rgba(245,158,11,0.3);white-space:nowrap;vertical-align:middle;">
                    <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;"><path d="M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82z"/><line x1="7" y1="7" x2="7.01" y2="7"/></svg>
                    IT SALSA - Categoria SAP &rarr; ${escapeHtml(categoriaVal)}
                </span>`
            : '';
        document.getElementById('sa-result-title').innerHTML = `Pr&eacute;-an&aacute;lise do Ticket&nbsp;&mdash;&nbsp;<a href="${jiraBase}/browse/${ticketKey}" target="_blank" style="color: var(--primary); text-decoration: none; font-weight: 800;">${ticketKey}</a>${categBadgeHtml}`;

        const badge = document.getElementById('sa-confidence-badge');
        badge.textContent = 'PREVIA DO TICKET';
        badge.style.background = 'rgba(59,130,246,0.15)';
        badge.style.color = '#3b82f6';
        badge.style.border = '1px solid rgba(59,130,246,0.3)';

        const signalsList = document.getElementById('sa-signals-list');
        signalsList.innerHTML = '';

        const fields = [
            { label: 'Transacao', val: sig.transaction || (firstMatch && firstMatch.transacao_sap) || null },
            { label: 'Programa/Classe', val: sig.program || null },
            { label: 'Mensagem SAP', val: sig.message_id ? `${sig.message_id} ${sig.message_number || ''}`.trim() : null },
            { label: 'Empresa', val: sig.company_code || null },
            { label: 'Documento', val: sig.document_number || null },
            { label: 'Exercicio', val: sig.fiscal_year || null },
            { label: 'Job', val: sig.job_name || null },
            { label: 'Utilizador', val: sig.user || null },
            { label: 'Categoria SAP', val: details.categoria_sap || null },
        ];

        if (firstMatch && firstMatch.nome_parametro) {
            fields.push({ label: 'Regra de contexto', val: firstMatch.nome_parametro });
        }

        fields.forEach(f => {
            const displayVal = f.val
                ? `<span style="font-weight:600; color:var(--text-primary);">${escapeHtml(f.val)}</span>`
                : '<span style="color:var(--text-secondary); font-style:italic;">nao identificada</span>';
            signalsList.innerHTML += `<div style="font-size: 0.8rem; border-bottom: 1px solid rgba(255,255,255,0.02); padding-bottom: 6px;">
                <div style="color:var(--text-secondary); font-size:0.72rem; font-weight:600; text-transform:uppercase; margin-bottom: 2px;">${f.label}</div>
                <div>${displayVal}</div>
            </div>`;
        });

        document.getElementById('sa-attachment-texts-container').style.display = 'none';
        document.getElementById('sa-attachment-texts').textContent = '';

        const evidencesList = document.getElementById('sa-evidences-list');
        evidencesList.innerHTML = '';
        const previewEvidences = [
            {
                name: 'Leitura do ticket',
                details: 'Os sinais acima foram preenchidos a partir do texto do ticket antes da validacao SAP.',
            },
            {
                name: firstMatch ? 'Regra de contexto aplicada' : 'Validacao SAP pendente',
                details: firstMatch
                    ? `A regra "${firstMatch.nome_parametro || 'Sem nome'}" foi usada para enriquecer a previa antes do login SAP.`
                    : 'O cockpit vai agora abrir a sessao SAP e validar os sinais acima em modo leitura.',
            },
        ];

        previewEvidences.forEach(e => {
            evidencesList.innerHTML += `<div style="font-size:0.8rem; display:flex; align-items:flex-start; gap:8px; margin-bottom: 6px;">
                <span>&bull;</span>
                <div>
                    <strong style="color:var(--text-primary);">${escapeHtml(e.name)}</strong>:
                    <span style="color:var(--text-secondary);">${escapeHtml(e.details)}</span>
                </div>
            </div>`;
        });

        document.getElementById('sa-possible-cause').textContent = 'Pre-leitura concluida. A causa provavel sera refinada depois da validacao no SAP.';
        document.getElementById('sa-proposed-solution').textContent = 'Os sinais identificados no ticket ja foram preenchidos. A seguir o agente valida os dados em SAP, em modo leitura.';
        document.getElementById('sa-suggested-tests').textContent = '1. Aguardar a conclusao da leitura SAP.\n2. Comparar os sinais extraidos do ticket com as evidencias validadas no SAP.';

        card.style.display = 'block';
    }

    async function saRunAnalysis() {
        const input = document.getElementById('sa-ticket-key');
        if (!input) return;
        const key = input.value.trim().toUpperCase();
        if (!key) {
            showToast('Por favor, introduza a chave do ticket JIRA.', 'error');
            return;
        }

        const btn = document.getElementById('sa-btn-analyze');
        const card = document.getElementById('sa-result-card');
        const detailsBox = document.getElementById('sa-ticket-description-box');
        const detailsText = document.getElementById('sa-ticket-description-text');
        if (card) card.style.display = 'none';

        if (btn) {
            btn.disabled = true;
            btn.innerHTML = `
                <svg class="spinner" width="14" height="14" viewBox="0 0 50 50" style="animation: rotate 2s linear infinite; margin-right: 8px; display: inline-block; vertical-align: middle;">
                    <circle class="path" cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" style="stroke-dasharray: 1, 150; stroke-dashoffset: 0; animation: dash 1.5s ease-in-out infinite;"></circle>
                </svg>
                <span>A analisar...</span>
            `;
        }

        try {
            // 1. Buscar descricao do ticket antes da validacao SAP
            if (detailsBox) {
                detailsBox.style.display = 'none';
            }
            if (detailsText) {
                detailsText.textContent = 'A carregar descricao...';
            }

            let preview = null;
            try {
                const previewRes = await fetch(`/api/jira/tickets/${encodeURIComponent(key)}/details`, { cache: 'no-store' });
                const previewData = await previewRes.json().catch(() => ({}));
                if (!previewRes.ok) {
                    throw new Error(previewData.detail || `Falha ao consultar o Jira (HTTP ${previewRes.status}).`);
                }
                preview = previewData;
            } catch (previewError) {
                throw new Error(`Não foi possível receber as informações do Jira: ${previewError.message}`);
            }

            if (preview) {
                const previewDesc = String(preview.description || '').trim();
                const previewSumm = String(preview.summary || '').trim();

                saTicketSummary = previewSumm;
                saTicketDescription = previewDesc;

                if (detailsBox) {
                    detailsBox.style.display = previewDesc ? 'block' : 'none';
                }
                if (detailsText) {
                    detailsText.textContent = previewDesc || '';
                }

                saRenderPreview(key, preview);

                if (saChatTicketKey === key) {
                    saInitChat(key, previewSumm, previewDesc);
                }
            }

            // 2. Lancar analise SAP
            const res = await fetch(`/api/sap-agent/analyze/${encodeURIComponent(key)}`, { method: 'POST' });
            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.detail || 'Falha ao iniciar análise.');
            }
            const data = await res.json();
            const jobId = data.job_id;

            // Começar a fazer polling do job
            let attempts = 0;
            const maxAttempts = 45; // 90 segundos máximo
            const interval = setInterval(async () => {
                attempts++;
                if (attempts > maxAttempts) {
                    clearInterval(interval);
                    if (btn) {
                        btn.disabled = false;
                        btn.innerHTML = '<span>🔍 Analisar</span>';
                    }
                    showToast('Tempo limite de análise esgotado. Confirme se o worker Windows está ligado.', 'error');
                    return;
                }

                try {
                    const jobRes = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`);
                    if (!jobRes.ok) return;
                    const job = await jobRes.json();

                    if (job.state === 'succeeded') {
                        clearInterval(interval);
                        if (btn) {
                            btn.disabled = false;
                            btn.innerHTML = '<span>🔍 Analisar</span>';
                        }
                        showToast('Análise concluída com sucesso!', 'success');
                        saRenderResult(job.status);
                    } else if (job.state === 'failed') {
                        clearInterval(interval);
                        if (btn) {
                            btn.disabled = false;
                            btn.innerHTML = '<span>🔍 Analisar</span>';
                        }
                        showToast('Erro durante a análise do Agente SAP.', 'error');
                        saRenderError(job.status || 'Erro desconhecido');
                    }
                } catch (e) {
                    // Ignora erros temporários de rede no polling
                }
            }, 2000);

        } catch (e) {
            showToast('Erro: ' + e.message, 'error');
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = '<span>🔍 Analisar</span>';
            }
        }
    }

    function saRenderResult(statusJson) {
        const card = document.getElementById('sa-result-card');
        if (!card) return;

        try {
            const report = JSON.parse(statusJson);
            
            // 1. Título e Confiança
            document.getElementById('sa-result-title').innerHTML = `Relatório de Análise — <a href="${jiraBase}/browse/${report.ticket_key}" target="_blank" style="color: var(--primary); text-decoration: none; font-weight: 800;">${report.ticket_key}</a>`;
            const badge = document.getElementById('sa-confidence-badge');
            badge.textContent = `Confiança: ${report.confidence.toUpperCase()}`;
            if (report.confidence === 'alta') {
                badge.style.background = 'rgba(16,185,129,0.15)';
                badge.style.color = '#10b981';
                badge.style.border = '1px solid rgba(16,185,129,0.3)';
            } else if (report.confidence === 'média') {
                badge.style.background = 'rgba(245,158,11,0.15)';
                badge.style.color = '#f59e0b';
                badge.style.border = '1px solid rgba(245,158,11,0.3)';
            } else {
                badge.style.background = 'rgba(239,68,68,0.15)';
                badge.style.color = '#ef4444';
                badge.style.border = '1px solid rgba(239,68,68,0.3)';
            }

            // 2. Sinais identificados
            const signalsList = document.getElementById('sa-signals-list');
            signalsList.innerHTML = '';
            const sig = report.signal || {};
            const fields = [
                { label: 'Transação', val: sig.transaction ? (sig.transaction_description ? `${sig.transaction} (${sig.transaction_description})` : sig.transaction) : null },
                { label: 'Programa/Classe', val: sig.program ? (sig.program_description ? `${sig.program} (${sig.program_description})` : sig.program) : null },
                { label: 'Mensagem SAP', val: sig.message_id ? `${sig.message_id} ${sig.message_number || ''}` : null },
                { label: 'Empresa', val: sig.company_code },
                { label: 'Documento', val: sig.document_number },
                { label: 'Exercício', val: sig.fiscal_year },
                { label: 'Job', val: sig.job_name },
                { label: 'Utilizador', val: sig.user },
                { label: 'Anexos', val: report.ticket_attachments && report.ticket_attachments.length > 0 ? report.ticket_attachments.join(', ') : null }
            ];
            fields.forEach(f => {
                const displayVal = f.val ? `<span style="font-weight:600; color:var(--text-primary);">${escapeHtml(f.val)}</span>` : '<span style="color:var(--text-secondary); font-style:italic;">não identificada</span>';
                signalsList.innerHTML += `<div style="font-size: 0.8rem; border-bottom: 1px solid rgba(255,255,255,0.02); padding-bottom: 6px;">
                    <div style="color:var(--text-secondary); font-size:0.72rem; font-weight:600; text-transform:uppercase; margin-bottom: 2px;">${f.label}</div>
                    <div>${displayVal}</div>
                </div>`;
            });

            // Attachment texts
            const attTexts = report.ticket_attachment_texts || [];
            if (attTexts.length > 0) {
                document.getElementById('sa-attachment-texts-container').style.display = 'block';
                document.getElementById('sa-attachment-texts').textContent = attTexts.join('\n\n');
            } else {
                document.getElementById('sa-attachment-texts-container').style.display = 'none';
                document.getElementById('sa-attachment-texts').textContent = '';
            }

            // 3. Evidências
            const evidencesList = document.getElementById('sa-evidences-list');
            evidencesList.innerHTML = '';
            const evs = report.evidences || [];
            if (evs.length === 0) {
                evidencesList.innerHTML = '<div style="font-size:0.8rem; color:var(--text-secondary); font-style:italic;">Nenhuma evidência recolhida.</div>';
            } else {
                evs.forEach(e => {
                    let dot = '';
                    if (e.status === 'ok') dot = '🟢';
                    else if (e.status === 'warning') dot = '🟡';
                    else dot = '🔴';
                    evidencesList.innerHTML += `<div style="font-size:0.8rem; display:flex; align-items:flex-start; gap:8px; margin-bottom: 6px;">
                        <span>${dot}</span>
                        <div>
                            <strong style="color:var(--text-primary);">${escapeHtml(e.name)}</strong>: 
                            <span style="color:var(--text-secondary);">${escapeHtml(e.details)}</span>
                        </div>
                    </div>`;
                });
            }

            // 4. Possível Causa, Solução e Testes
            document.getElementById('sa-possible-cause').textContent = report.probable_cause || 'Sem diagnóstico de causa.';
            document.getElementById('sa-proposed-solution').textContent = report.proposed_solution || 'Sem solução proposta.';
            
            const testsList = report.tests_to_execute || [];
            document.getElementById('sa-suggested-tests').textContent = testsList.length > 0 ? testsList.map((t, i) => `${i+1}. ${t}`).join('\n') : 'Sem testes sugeridos.';

            card.style.display = 'block';
            saInitChat(report.ticket_key);

        } catch (err) {
            saRenderError('Falha ao processar relatório da análise: ' + err.message);
        }
    }

    function saRenderError(errMsg) {
        const card = document.getElementById('sa-result-card');
        if (!card) return;
        
        document.getElementById('sa-result-title').textContent = 'Falha na Análise';
        const badge = document.getElementById('sa-confidence-badge');
        badge.textContent = 'ERRO';
        badge.style.background = 'rgba(239,68,68,0.15)';
        badge.style.color = '#ef4444';
        badge.style.border = '1px solid rgba(239,68,68,0.3)';

        document.getElementById('sa-signals-list').innerHTML = '';
        document.getElementById('sa-evidences-list').innerHTML = '';
        
        document.getElementById('sa-possible-cause').innerHTML = `<span style="color:#ef4444; font-weight:600;">Ocorreu um erro ao processar o ticket:</span>\n${escapeHtml(errMsg)}`;
        document.getElementById('sa-proposed-solution').textContent = 'Por favor, valide se o JIRA está acessível, se a chave do ticket está correta e se a conexão SAP está configurada adequadamente no arquivo .env.';
        document.getElementById('sa-suggested-tests').textContent = '1. Verificar conexão de rede.\n2. Analisar logs detalhados do worker no PowerShell.';

        card.style.display = 'block';
        const inputKey = (document.getElementById('sa-ticket-key').value || '').trim().toUpperCase();
        saInitChat(inputKey || 'Ticket');
    }

    let saChatHistory = [];
    let saChatTicketKey = "";
    let saTicketSummary = "";
    let saTicketDescription = "";

    async function saInitChat(ticketKey, summary, description) {
        saChatTicketKey = ticketKey;
        saChatHistory = [];
        const messagesDiv = document.getElementById('sa-chat-messages');
        if (!messagesDiv) return;

        // Usar valores passados ou os guardados globalmente
        const summ = (summary || saTicketSummary || '').trim();
        const desc = (description || saTicketDescription || '').trim();

        messagesDiv.innerHTML = "";

        // Se não temos descrição do ticket ainda, mostrar mensagem estática simples e aguardar
        if (!summ && !desc) {
            const waitBubble = document.createElement('div');
            waitBubble.className = "chat-msg-bubble chat-msg-bot";
            waitBubble.innerHTML = `A carregar contexto do ticket <strong>${ticketKey}</strong>...`;
            messagesDiv.appendChild(waitBubble);
            return;
        }

        // Mostrar bubble de "a pensar..." enquanto a IA processa
        const loadingBubble = document.createElement('div');
        loadingBubble.className = "chat-msg-bubble chat-msg-bot chat-msg-loading";
        loadingBubble.innerHTML = '<span style="display:inline-flex;align-items:center;gap:6px;"><span style="font-size:1rem">🤖</span> A analisar o pedido do ticket...</span>';
        messagesDiv.appendChild(loadingBubble);
        messagesDiv.scrollTop = messagesDiv.scrollHeight;

        // Construir a primeira mensagem automática com o contexto do ticket
        let primeiroMensagem = 'Com base no pedido deste ticket JIRA, o que devemos fazer em SAP?';
        if (summ || desc) {
            const partes = [];
            if (summ) partes.push(`Motivo de abertura: ${summ}`);
            if (desc) partes.push(`Pedido do utilizador:\n${desc}`);
            primeiroMensagem = `${partes.join('\n\n')}\n\nCom base neste pedido, o que devemos fazer em SAP? Dá-me orientações concretas e práticas.`;
        }

        try {
            const payload = {
                ticket_key: ticketKey,
                message: primeiroMensagem,
                history: [],
                sap_query_enabled: false
            };

            const response = await fetch('/api/sap-agent/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (loadingBubble.parentNode) messagesDiv.removeChild(loadingBubble);

            if (!response.ok) {
                const errData = await response.json();
                const detail = errData.detail || 'Falha de comunicação.';
                const is503 = response.status === 503 || detail.includes('503') || detail.includes('high demand') || detail.includes('UNAVAILABLE');
                const errBubble = document.createElement('div');
                errBubble.className = "chat-msg-bubble chat-msg-system";
                errBubble.innerHTML = is503
                    ? `⏳ <strong>A API de IA está temporariamente sobrecarregada.</strong> O servidor tentou automaticamente — por favor aguarda alguns segundos e clica em Tentar novamente.<br><br>
                       <button onclick="saInitChat(saChatTicketKey, saTicketSummary, saTicketDescription)" style="margin-top:6px;padding:5px 14px;border-radius:7px;border:none;background:linear-gradient(135deg,#3b82f6,#2563eb);color:#fff;font-size:0.78rem;font-weight:700;cursor:pointer;display:inline-flex;align-items:center;gap:6px;">
                         🔄 Tentar novamente
                       </button>`
                    : `❌ Erro ao gerar análise inicial: ${escapeHtml(detail)}<br><br>
                       <button onclick="saInitChat(saChatTicketKey, saTicketSummary, saTicketDescription)" style="margin-top:6px;padding:5px 14px;border-radius:7px;border:none;background:rgba(255,255,255,0.08);color:var(--text-primary);font-size:0.78rem;font-weight:600;cursor:pointer;border:1px solid rgba(255,255,255,0.12);">
                         🔄 Tentar novamente
                       </button>`;
                messagesDiv.appendChild(errBubble);
            } else {
                const data = await response.json();
                const reply = data.reply || "";

                const botBubble = document.createElement('div');
                botBubble.className = "chat-msg-bubble chat-msg-bot";
                let formatted = escapeHtml(reply)
                    .replace(/\n/g, '<br>')
                    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                    .replace(/`([^`]+)`/g, '<code>$1</code>');
                botBubble.innerHTML = formatted;
                messagesDiv.appendChild(botBubble);

                // Guardar no histórico para contexto nas mensagens seguintes
                saChatHistory.push({ role: "user", text: primeiroMensagem });
                saChatHistory.push({ role: "model", text: reply });
            }
        } catch (e) {
            if (loadingBubble.parentNode) messagesDiv.removeChild(loadingBubble);
            const errBubble = document.createElement('div');
            errBubble.className = "chat-msg-bubble chat-msg-system";
            errBubble.innerHTML = `⚠️ Falha de rede: <em>${escapeHtml(e.message)}</em><br><br>
                <button onclick="saInitChat(saChatTicketKey, saTicketSummary, saTicketDescription)" style="margin-top:6px;padding:5px 14px;border-radius:7px;border:none;background:rgba(255,255,255,0.08);color:var(--text-primary);font-size:0.78rem;font-weight:600;cursor:pointer;border:1px solid rgba(255,255,255,0.12);">
                  🔄 Tentar novamente
                </button>`;
            messagesDiv.appendChild(errBubble);
        }

        messagesDiv.scrollTop = messagesDiv.scrollHeight;
    }

    async function saSendChatMessage() {
        const input = document.getElementById('sa-chat-input');
        const btn = document.getElementById('sa-chat-btn-send');
        const messagesDiv = document.getElementById('sa-chat-messages');
        if (!input || !messagesDiv || !saChatTicketKey) return;

        const message = input.value.trim();
        if (!message) return;

        // Render user message
        const userBubble = document.createElement('div');
        userBubble.className = "chat-msg-bubble chat-msg-user";
        userBubble.textContent = message;
        messagesDiv.appendChild(userBubble);
        
        input.value = "";
        messagesDiv.scrollTop = messagesDiv.scrollHeight;

        input.disabled = true;
        if (btn) btn.disabled = true;

        // Detectar intenção SAP para mostrar loading contextual
        const sapGuiPatterns = /\b(entr[ae]\s+na?|abr[ae]|pesquisa|pesquise|consulta|consulte|mostra|mostre|vai para|vá para)\b.{0,50}\b(tabela|transaction|se16n|se16|ko03|me23n|fb03|ekko|aufk|bkpf|ekpo|anla|prps)\b/i;
        const sapObjectPatterns = /\b(entr[ae]|analisa|analise|verifica|verifique|vai|vá|abre|veja|ver|mostra|mostre|pedido|ordem|consulta|consulte)\b.{0,50}\b(\d{7,12}|pedido|ordem|po|imobilizado|documento|wbs)\b/i;
        const isSapIntent = sapGuiPatterns.test(message) || sapObjectPatterns.test(message) || /\b\d{7,12}\b/.test(message);

        const loadingBubble = document.createElement('div');
        loadingBubble.className = "chat-msg-bubble chat-msg-bot chat-msg-loading";
        loadingBubble.innerHTML = isSapIntent
            ? '<span style="display:inline-flex;align-items:center;gap:6px;"><span style="font-size:1rem">🔌</span> A consultar SAP em tempo real...</span>'
            : 'A pensar...';
        messagesDiv.appendChild(loadingBubble);
        messagesDiv.scrollTop = messagesDiv.scrollHeight;

        try {
            const payload = {
                ticket_key: saChatTicketKey,
                message: message,
                history: saChatHistory,
                sap_query_enabled: true
            };

            const response = await fetch('/api/sap-agent/chat', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (loadingBubble.parentNode) {
                messagesDiv.removeChild(loadingBubble);
            }

            if (!response.ok) {
                const errData = await response.json();
                const errBubble = document.createElement('div');
                errBubble.className = "chat-msg-bubble chat-msg-system";
                errBubble.textContent = `Erro: ${errData.detail || 'Falha ao obter resposta do assistente.'}`;
                messagesDiv.appendChild(errBubble);
            } else {
                const data = await response.json();
                const reply = data.reply || "";

                // === SAP GUI Action: o worker irá executar no SAP ===
                if (data.waiting_sap && data.job_id) {
                    // Mostrar bubble inicial com spinner de execução SAP
                    const sapBubble = document.createElement('div');
                    sapBubble.className = "chat-msg-bubble chat-msg-bot";
                    sapBubble.innerHTML = `
                        <span style="display:inline-flex;align-items:center;gap:8px;margin-bottom:8px;">
                            <span style="display:inline-flex;align-items:center;gap:4px;background:rgba(234,179,8,0.15);border:1px solid rgba(234,179,8,0.3);border-radius:6px;padding:2px 8px;font-size:0.72rem;color:#fbbf24;font-weight:700;">
                                ⚙️ SAP GUI
                            </span>
                        </span><br>
                        <strong>${escapeHtml(reply.replace(/\*\*(.*?)\*\*/g, '$1'))}</strong>
                        <div id="sap-job-status-${data.job_id}" style="margin-top:10px;color:var(--text-muted);font-size:0.8rem;">
                            <span class="chat-loading-dots">⏳ A aguardar o worker SAP</span>
                        </div>`;
                    messagesDiv.appendChild(sapBubble);

                    saChatHistory.push({ role: "user", text: message });

                    // Iniciar polling para o resultado
                    saStartJobPolling(data.job_id, sapBubble, saChatHistory);
                    return;
                }

                // === Resposta de texto normal do Gemini ===
                const botBubble = document.createElement('div');
                botBubble.className = "chat-msg-bubble chat-msg-bot";

                // Badge "📊 Dados SAP" quando a resposta menciona dados reais
                const hasSapData = reply.includes('📊 Dados reais lidos do SAP') || reply.includes('📌 Orientação de consulta SAP');
                let badge = '';
                if (hasSapData) {
                    badge = '<span style="display:inline-flex;align-items:center;gap:4px;background:rgba(59,130,246,0.15);border:1px solid rgba(59,130,246,0.3);border-radius:6px;padding:2px 7px;font-size:0.72rem;color:#60a5fa;margin-bottom:8px;font-weight:600;">📊 Dados lidos do SAP</span><br>';
                }
                
                let formatted = escapeHtml(reply)
                    .replace(/\n/g, '<br>')
                    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                    .replace(/`([^`]+)`/g, '<code>$1</code>');
                botBubble.innerHTML = badge + formatted;
                messagesDiv.appendChild(botBubble);

                saChatHistory.push({ role: "user", text: message });
                saChatHistory.push({ role: "model", text: reply });
            }
        } catch (e) {
            if (loadingBubble.parentNode) {
                messagesDiv.removeChild(loadingBubble);
            }
            const errBubble = document.createElement('div');
            errBubble.className = "chat-msg-bubble chat-msg-system";
            errBubble.textContent = `Falha de rede: ${e.message}`;
            messagesDiv.appendChild(errBubble);
        } finally {
            input.disabled = false;
            if (btn) btn.disabled = false;
            input.focus();
            messagesDiv.scrollTop = messagesDiv.scrollHeight;
        }
    }

    function saClearChatHistory() {
        if (confirm("Tem a certeza que deseja limpar o histórico de conversção deste ticket?")) {
            saInitChat(saChatTicketKey);
        }
    }

    // ── SAP Job Polling ──────────────────────────────────────────────────────

    async function saStartJobPolling(jobId, bubble, history) {
        const statusDiv = document.getElementById(`sap-job-status-${jobId}`);
        const messagesDiv = document.getElementById('sa-chat-messages');
        const MAX_POLLS = 30; // 30 × 2s = 60s máximo
        let polls = 0;

        const interval = setInterval(async () => {
            polls++;
            try {
                const r = await fetch(`/api/sap-agent/chat-job/${jobId}`);
                if (!r.ok) {
                    clearInterval(interval);
                    if (statusDiv) statusDiv.innerHTML = '❌ Erro ao verificar job SAP.';
                    return;
                }
                const job = await r.json();

                if (job.state === 'succeeded') {
                    clearInterval(interval);
                    const sapResultText = job.result_text || '';
                    const sapRows = job.rows || [];

                    // Atualizar a bubble com o resultado
                    if (statusDiv) {
                        let resultHtml = '';
                        if (sapRows.length > 0) {
                            resultHtml = saSapTableHtml(sapRows, job.description || '');
                        } else if (sapResultText) {
                            resultHtml = `<div style="margin-top:10px;padding:10px;background:rgba(0,0,0,0.1);border-radius:8px;font-size:0.8rem;">${
                                escapeHtml(sapResultText).replace(/\n/g, '<br>').replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                            }</div>`;
                        } else {
                            resultHtml = '<div style="margin-top:8px;color:var(--text-muted);">Execução concluída (sem dados para mostrar).</div>';
                        }

                        statusDiv.innerHTML = `
                            <div style="margin-top:6px;display:inline-flex;align-items:center;gap:4px;background:rgba(34,197,94,0.15);border:1px solid rgba(34,197,94,0.3);border-radius:6px;padding:2px 8px;font-size:0.72rem;color:#4ade80;font-weight:600;">
                                ✅ Executado no SAP GUI
                            </div>
                            ${resultHtml}`;
                    }

                    // Guardar no histórico do chat
                    history.push({ role: "model", text: sapResultText || 'Ação SAP executada com sucesso.' });
                    if (messagesDiv) messagesDiv.scrollTop = messagesDiv.scrollHeight;

                } else if (job.state === 'failed') {
                    clearInterval(interval);
                    if (statusDiv) {
                        statusDiv.innerHTML = `<div style="color:#f87171;margin-top:8px;">❌ Erro SAP: ${escapeHtml(job.error || 'Falha desconhecida')}</div>`;
                    }
                    history.push({ role: "model", text: `Erro ao executar no SAP: ${job.error || 'Falha desconhecida'}` });

                } else if (polls >= MAX_POLLS) {
                    clearInterval(interval);
                    if (statusDiv) statusDiv.innerHTML = '⏰ Timeout: o worker SAP não respondeu em 60 segundos.';
                } else {
                    // Ainda em progressão
                    const dots = '.'.repeat((polls % 3) + 1);
                    if (statusDiv) statusDiv.innerHTML = `<span class="chat-loading-dots">⏳ A executar no SAP GUI${dots}</span>`;
                }
            } catch (err) {
                clearInterval(interval);
                if (statusDiv) statusDiv.innerHTML = `❌ Falha de rede ao verificar job: ${escapeHtml(err.message)}`;
            }
        }, 2000);
    }

    function saSapTableHtml(rows, description) {
        if (!rows || rows.length === 0) return '';
        const cols = Object.keys(rows[0]);
        const headerRow = cols.map(c => `<th style="padding:4px 8px;background:rgba(255,255,255,0.06);border-bottom:1px solid var(--border);font-size:0.72rem;font-weight:700;white-space:nowrap;">${escapeHtml(c)}</th>`).join('');
        const dataRows = rows.map(row =>
            `<tr style="border-bottom:1px solid rgba(255,255,255,0.04);">${
                cols.map(c => `<td style="padding:3px 8px;font-size:0.72rem;white-space:nowrap;">${escapeHtml(String(row[c] || ''))}</td>`).join('')
            }</tr>`
        ).join('');

        return `
            <div style="margin-top:10px;overflow-x:auto;border-radius:8px;border:1px solid var(--border);">
                <div style="padding:6px 10px;background:rgba(59,130,246,0.1);border-bottom:1px solid var(--border);font-size:0.72rem;color:#60a5fa;font-weight:600;">
                    📊 ${escapeHtml(description || 'Dados SAP')} — ${rows.length} linha(s)
                </div>
                <div style="overflow-x:auto;">
                    <table style="width:100%;border-collapse:collapse;">
                        <thead><tr>${headerRow}</tr></thead>
                        <tbody>${dataRows}</tbody>
                    </table>
                </div>
            </div>`;
    }

    // ── Event delegation for Reply-to-customer toggle (avoids inline onclick escaping) ──
    document.addEventListener('click', function(e) {
        const el = e.target.closest('.js-reply-toggle');
        if (!el) return;
        e.preventDefault();
        const key      = el.dataset.key;
        const expandId = el.dataset.expandId;
        if (key && expandId) toggleReplyExpand(key, expandId);
    });

    // ── Event delegation for Reply-to-customer save button ──
    document.addEventListener('click', function(e) {
        const el = e.target.closest('.js-reply-save');
        if (!el) return;
        e.preventDefault();
        const key      = el.dataset.key;
        const expandId = el.dataset.expandId;
        if (key && expandId) saveReplyComment(key, expandId);
    });

    // ── Premium Tooltip for Ticket Description on Hover ──
    const ticketDescriptionCache = {};
    let tooltipTimeout = null;
    let activeTooltipKey = null;

    // Create global tooltip element
    const tooltipEl = document.createElement('div');
    tooltipEl.className = 'premium-tooltip';
    tooltipEl.style.position = 'absolute';
    tooltipEl.style.display = 'none';
    tooltipEl.style.zIndex = '99999';
    tooltipEl.style.pointerEvents = 'none';
    document.body.appendChild(tooltipEl);

    document.addEventListener('DOMContentLoaded', () => {
        const ticketsTable = document.getElementById('jira-tickets-table');
        if (!ticketsTable) return;

        ticketsTable.addEventListener('mouseover', (e) => {
            const link = e.target.closest('.summary-link');
            if (!link) return;

            const key = link.getAttribute('data-key');
            if (!key) return;

            if (tooltipTimeout) {
                clearTimeout(tooltipTimeout);
                tooltipTimeout = null;
            }

            activeTooltipKey = key;

            // Calculate position
            const rect = link.getBoundingClientRect();
            const tooltipWidth = 400;
            let left = window.scrollX + rect.left;
            if (rect.left + tooltipWidth > window.innerWidth) {
                left = window.scrollX + window.innerWidth - tooltipWidth - 20;
            }
            tooltipEl.style.left = `${left}px`;
            tooltipEl.style.top = `${window.scrollY + rect.bottom + 8}px`;
            tooltipEl.style.display = 'block';

            if (ticketDescriptionCache[key] !== undefined) {
                tooltipEl.textContent = ticketDescriptionCache[key] || 'Sem descrição disponível.';
            } else {
                tooltipEl.innerHTML = `
                    <div style="display: flex; align-items: center; gap: 8px; color: #9ca3af;">
                        <svg class="spinner" width="12" height="12" viewBox="0 0 50 50" style="animation: spin 1s linear infinite; display: inline-block;">
                            <circle class="path" cx="25" cy="25" r="20" fill="none" stroke="currentColor" stroke-width="5" stroke-linecap="round" style="stroke-dasharray: 1, 150; stroke-dashoffset: 0; animation: dash 1.5s ease-in-out infinite;"></circle>
                        </svg>
                        <span>A carregar descrição...</span>
                    </div>
                `;

                fetch(`/api/jira/tickets/${encodeURIComponent(key)}/details`)
                    .then(async r => {
                        const data = await r.json().catch(() => ({}));
                        if (!r.ok) throw new Error(data.detail || `Erro HTTP ${r.status}`);
                        return data;
                    })
                    .then(d => {
                        if (activeTooltipKey !== key) return;
                        const desc = d && d.description ? d.description.trim() : null;
                        ticketDescriptionCache[key] = desc || '';
                        tooltipEl.textContent = ticketDescriptionCache[key] || 'Sem descrição disponível.';
                    })
                    .catch(err => {
                        if (activeTooltipKey !== key) return;
                        tooltipEl.textContent = `Erro ao carregar descrição: ${err.message}`;
                    });
            }
        });

        ticketsTable.addEventListener('mouseout', (e) => {
            const link = e.target.closest('.summary-link');
            if (!link) return;

            activeTooltipKey = null;
            tooltipTimeout = setTimeout(() => {
                tooltipEl.style.display = 'none';
            }, 100);
        });
    });

    loadJobs().then(() => {
        startPolling();
        switchView('jira');
    });


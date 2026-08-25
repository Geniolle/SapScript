(function () {
    let durationInterval = null;
    let currentlyFocusedJob = null;

    function setText(id, value) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = value;
        }
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

    function startDurationTicker() {
        if (durationInterval) return;
        durationInterval = setInterval(() => {
            if (currentlyFocusedJob && (currentlyFocusedJob.state === 'running' || currentlyFocusedJob.state === 'pending')) {
                const durationStr = formatDuration(currentlyFocusedJob.created_at, null);
                setText('focused-kpi-duration', durationStr);
            } else {
                stopDurationTicker();
            }
        }, 1000);
    }

    function stopDurationTicker() {
        if (durationInterval) {
            clearInterval(durationInterval);
            durationInterval = null;
        }
    }

    window.DashboardKpis = {
        updateFocusedJob(job) {
            currentlyFocusedJob = job;
            if (!job) {
                stopDurationTicker();
                setText('focused-kpi-roles-title', 'Roles concluídas');
                setText('focused-kpi-roles', '0/0');
                setText('focused-kpi-roles-rate', '100%');
                setText('focused-kpi-actions', '0');
                setText('focused-kpi-duration', '00:00:00');
                setText('focused-kpi-errors', '0');
                setText('focused-kpi-errors-sub', 'Sem erros');
                return;
            }

            const p = job.params || {};
            const proc = p.processo || '';
            const sub = p.subprocesso || '';
            const isCadeiaProcess = proc.toLowerCase().includes('cadeia') || sub.toLowerCase().includes('cadeia') || job.task.toLowerCase().includes('cadeia');

            // Parse logs
            let logStr = job.log || '';
            logStr = logStr.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n');
            const lines = logStr.split('\n').filter(l => l.trim() !== '');

            let totalRoles = p.roles_count || 0;
            let currentRoleIndex = 0;
            let rolesFromLogs = [];
            let insideSummary = false;
            let completedRolesList = [];
            let currentRoleInLoop = '';
            let roleMetadata = {};
            let errorCount = 0;

            lines.forEach(line => {
                const tr = line.trim();
                if (!tr) return;

                if ((tr.includes('🔴 ERRO') || tr.includes('❌ SAP Erro:') || tr.includes('❌ Erro') || tr.includes('❌ Falha') || tr.startsWith('❌'))
                    && !tr.toLowerCase().includes('heartbeat')) {
                    errorCount++;
                }

                const totalCadeiasMatch = tr.match(/🔍\s*Total de cadeias a verificar:\s*(\d+)/i);
                if (totalCadeiasMatch) {
                    totalRoles = parseInt(totalCadeiasMatch[1]);
                }

                const cadeiaVerifyMatch = tr.match(/^([✅❌])\s*([^-\n]+)\s*-\s*(.*)$/);
                if (cadeiaVerifyMatch) {
                    const isSuccess = (cadeiaVerifyMatch[1] === '✅');
                    const cadeiaName = cadeiaVerifyMatch[2].trim();
                    const cadeiaStatus = cadeiaVerifyMatch[3].trim();
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
                    currentRoleIndex = rolesFromLogs.length;
                }

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

                const bracketMatch = tr.match(/▶\s*\[\d+\/\d+\]\s*(?:ROLE|CADEIA):\s*([A-Za-z0-9_\\\-]+)/i) || tr.match(/▶\s*\[(\d+)\/(\d+)\]/);
                if (bracketMatch) {
                    let idx = null;
                    let tot = null;
                    if (bracketMatch[1] && bracketMatch[2]) {
                        idx = parseInt(bracketMatch[1]);
                        tot = parseInt(bracketMatch[2]);
                    }
                    if (idx && idx > currentRoleIndex) {
                        currentRoleIndex = idx;
                    }
                    if (tot && tot > totalRoles) {
                        totalRoles = tot;
                    }
                    const nameMatch = tr.match(/(?:ROLE|CADEIA):\s*([A-Za-z0-9_\\\-]+)/i);
                    if (nameMatch && nameMatch[1]) {
                        currentRoleInLoop = nameMatch[1];
                        if (!rolesFromLogs.some(r => r.name === currentRoleInLoop)) {
                            rolesFromLogs.push({ name: currentRoleInLoop });
                        }
                        if (!roleMetadata[currentRoleInLoop]) {
                            roleMetadata[currentRoleInLoop] = { tcodes: null, actions: 0, duration: '' };
                        }
                    }
                }

                if (currentRoleInLoop && (tr.startsWith('├─') || tr.startsWith('└─'))) {
                    if (!roleMetadata[currentRoleInLoop]) {
                        roleMetadata[currentRoleInLoop] = { tcodes: null, actions: 0, duration: '' };
                    }
                    roleMetadata[currentRoleInLoop].actions++;
                }

                if (currentRoleInLoop && (tr.includes('🟢 SUCESSO') || tr.includes('🔴 ERRO') || tr.includes('Role concluida') || tr.includes('Role concluída') || tr.includes('tratada por completo'))) {
                    const durationMatch = tr.match(/\(Tempo:\s*([^)]+)\)/i);
                    if (durationMatch) {
                        if (!roleMetadata[currentRoleInLoop]) {
                            roleMetadata[currentRoleInLoop] = { tcodes: null, actions: 0, duration: '' };
                        }
                        roleMetadata[currentRoleInLoop].duration = durationMatch[1].trim();
                    }
                }

                if (tr.includes('🟢 SUCESSO') || tr.includes('Role concluida:') || tr.includes('Role concluída:') || tr.includes('[OK] Role concluida:') || tr.includes('[OK] Role concluída:')) {
                    if (currentRoleInLoop) {
                        completedRolesList.push(currentRoleInLoop);
                    }
                    const parts = tr.split(':');
                    if (parts[1] && (parts[0].includes('Role concluida') || parts[0].includes('Role concluída'))) {
                        completedRolesList.push(parts[1].trim());
                    }
                }
            });

            if (totalRoles === 0 && rolesFromLogs.length > 0) {
                totalRoles = rolesFromLogs.length;
            }
            if (job.state === 'failed' && errorCount === 0) {
                errorCount = 1;
            }
            completedRolesList = [...new Set(completedRolesList)];

            if (job.state === 'running' && !currentRoleInLoop) {
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

            let totalActions = 0;
            const allRolesParam = (p.roles && p.roles.length > 0) ? p.roles : rolesFromLogs;
            if (totalRoles > 0) {
                allRolesParam.forEach((r) => {
                    const meta = roleMetadata[r.name] || {};
                    totalActions += meta.actions || 0;
                });
            }

            const durationStr = formatDuration(job.created_at, (job.state === 'running' || job.state === 'pending') ? null : job.updated_at);

            // Update UI
            setText('focused-kpi-roles-title', isCadeiaProcess ? 'Cadeias verificadas' : 'Roles concluídas');
            setText('focused-kpi-roles', `${concludedRoles}/${totalRoles}`);
            const rate = totalRoles > 0 ? Math.round((concludedRoles / totalRoles) * 100) : 100;
            setText('focused-kpi-roles-rate', `${rate}%`);
            setText('focused-kpi-actions', totalActions);
            setText('focused-kpi-duration', durationStr);
            setText('focused-kpi-errors', errorCount);
            setText('focused-kpi-errors-sub', errorCount > 0 ? 'Requer atenção' : 'Sem erros');

            if (job.state === 'running' || job.state === 'pending') {
                startDurationTicker();
            } else {
                stopDurationTicker();
            }
        }
    };
})();

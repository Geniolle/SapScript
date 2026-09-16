# Diretrizes do Projeto SapScript (Antigravity Rules)

## REGRA DE OURO Nº 1: PRÉ-VALIDAÇÃO MANDATÓRIA (CHECKLIST PRÉ-VOO)
Antes de iniciar qualquer execução, modificação no SAP (PFCG/SU01/CUA) ou alteração em lote no Excel:
1. **Verificação de Pré-requisitos**:
   - Confirmar se o departamento possui utilizadores válidos na folha `Proposta Ativa`.
   - Confirmar se o departamento está configurado na folha `DEFINIÇÕES` (regras organizacionais).
   - Verificar se as roles necessárias já existem na folha `PFCG_CREATE` / `PFCG_COMPOSTA` e no SAP PRD.
   - Se faltar qualquer pré-requisito, **NÃO iniciar a automação**: reportar as pendências imediatamente ao utilizador.
2. **Estimativa de Escopo & Faseamento**:
   - Para volumes superiores a 10 utilizadores ou processos complexos, segmentar o processamento em etapas claras e avisar o utilizador sobre o plano de execução.

## REGRA DE OURO Nº 2: EXECUÇÃO ATÓMICA E IDEMPOTENTE (CHECKPOINTS)
- **Não executar rotinas em bloco cego**: cada utilizador processado deve ser auditado em tempo real ($P == Q$) e imediatamente registado.
- Se o processo for interrompido a qualquer momento, o reinício deve ser 100% idempotente (ignora quem já está concluído e retoma no primeiro pendente).
- Sempre criar backup com timestamp em `output/` antes de salvar ficheiros Excel.

## REGRA DE OURO Nº 3: EFICIÊNCIA DE RECURSOS E BACKGROUND TASKS
- Execuções longas no SAP GUI e processamentos de dados devem correr via scripts Python em **background tasks locais** (`run_command`), evitando consumo desnecessário de tokens de contexto durante loops operacionais.
- O contexto da conversa é reservado para raciocínio analítico, planeamento, auditoria e comunicação executiva com o utilizador.

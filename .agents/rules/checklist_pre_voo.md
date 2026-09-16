# Regra de Pré-Validação Obrigatória e Gestão Segura de Execução

## 1. Regra de Ouro: Checklist Pré-Voo
Antes de qualquer modificação no SAP (PFCG, CUA, SU01) ou alteração em massa no Excel:
1. Validar se o departamento possui utilizadores válidos na folha `Proposta Ativa`.
2. Validar se o departamento está configurado na folha `DEFINIÇÕES` (regras organizacionais).
3. Validar se as roles necessárias já existem na folha `PFCG_CREATE` / `PFCG_COMPOSTA` e no SAP PRD.
4. Se faltar qualquer pré-requisito, **NÃO iniciar a automação**: reportar as pendências imediatamente ao utilizador.

## 2. Idempotência e Checkpoints
- Cada utilizador processado deve ser auditado em tempo real ($P == Q$) e imediatamente registado.
- Em caso de interrupção, o reinício deve ser 100% idempotente (retoma no primeiro utilizador pendente).
- Sempre criar backup com timestamp em `output/` antes de salvar ficheiros Excel.

## 3. Background Tasks Locais
- Processamentos longos devem correr via scripts Python em background tasks (`run_command`), preservando os tokens do contexto da conversa para raciocínio analítico e comunicação.

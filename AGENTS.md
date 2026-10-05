# Instruções persistentes do projeto

## Comando de encerramento "Perfeito"

Quando o utilizador disser **"Perfeito"** como confirmação de que o trabalho atual foi concluído:

1. Atualizar a documentação ou memória relevante do projeto com as decisões, resultados e validações do trabalho concluído.
2. Rever as alterações pendentes e não incluir ficheiros com segredos, credenciais ou artefactos locais.
3. Executar as verificações técnicas proporcionais às alterações.
4. Criar um commit descritivo e publicar a branch atual no remoto GitHub configurado.
5. Informar o utilizador sobre os ficheiros atualizados, verificações executadas, commit e branch publicados.

Não interpretar a palavra dentro de uma citação, exemplo ou pergunta como ordem de publicação. A regra aplica-se quando ela for usada pelo utilizador como aprovação/encerramento do trabalho.

## Segurança SAP

As validações SAP devem ser feitas em modo somente leitura. Nunca alterar funções, atribuições, utilizadores ou outros dados SAP sem um pedido explícito e específico do utilizador.

## Posição da janela SAP GUI

Sempre que uma sessão SAP GUI for aberta ou trazida para primeiro plano, restaurar a
janela e posicioná-la na metade esquerda do ecrã, salvo indicação diferente do utilizador.

## Consulta de Processos em Estrutura.md

Antes de pesquisar a implementação de um processo no repositório:

1. Consultar primeiro `Estrutura.md` (na raiz do projeto `SapScript`).
2. Procurar o processo funcional solicitado.
3. Se existir, utilizar os caminhos, ficheiros e funções documentados como ponto inicial.
4. Confirmar essas referências no código atual.
5. Fazer pesquisa global apenas se o processo não estiver documentado, a referência estiver desatualizada ou forem necessárias dependências adicionais.
6. Não assumir que `Estrutura.md` contém todo o projeto: o documento é incremental.

O fluxo de trabalho esperado é:
`Pedido do utilizador → Estrutura.md → processo → ficheiros/funções → validação no código → análise`

## Regra "atualizar Estrutura.md"

Sempre que o utilizador solicitar **"atualiza o Estrutura.md"**, interpretar como:

1. Identificar o processo em que estamos atualmente a trabalhar.
2. Analisar a implementação REAL e atual desse processo no repositório.
3. Localizar frontend, API, worker/orquestração, serviços, SAP/RFC, scripts e testes relevantes.
4. Comparar com o que já existe em `Estrutura.md`.
5. Atualizar somente a área correspondente.
6. Adicionar caminhos/funções novos.
7. Corrigir referências alteradas.
8. Remover referências obsoletas daquela área.
9. Preservar integralmente as restantes áreas já documentadas.
10. Não tentar mapear processos não relacionados apenas para completar o documento.

`Estrutura.md` deve permanecer um índice técnico conciso e navegável, e não transformar-se numa documentação extensa da implementação.

## Extração de Dados SAP & Ficheiros Temporários

1. **Filtragem Local de Estruturas (`INTTAB`):** Em scripts de extração de tabelas via RFC (ex.: `RFC_READ_TABLE`), filtrar sempre por `TABCLASS = 'TRANSP'`, ignorando estruturas de memória (`INTTAB`) localmente. Elas não contêm registos persistidos e a sua tentativa de leitura gera falsos erros nos inventários.
2. **Isolamento de Artefactos Temporários (`output/.tmp/`):** Leituras temporárias, logs brutos e varreduras de teste devem ser guardados dentro de diretórios temporários (ex.: `SapScript/output/.tmp/` ou `tests/.tmp/`), evitando acumular ficheiros não estruturados na raiz.
3. **Garantia de Segurança SAP:** Qualquer eliminação ou limpeza é estritamente LOCAL nos inventários/manifestos da máquina. Nunca executar alterações ou eliminações no ambiente SAP.


# AGENTS.md

## Regra Operacional

- Sempre que forem feitas alterações no backend, reiniciar o Docker antes de validar o comportamento em runtime.
- Isso aplica-se a mudanças em `web_api`, `worker`, endpoints, jobs, handlers e qualquer lógica servida pelo container.
- Antes de pedir credenciais, host, cliente, sistema SAP ou dados RFC ao utilizador, verificar primeiro o `.env` da raiz `C:\workspace\SapScript\.env` e a configuração local do projeto.
- Se o `.env` já contiver a informação necessária, reutilizá-la diretamente e não solicitar novamente ao utilizador.

## Padrão de Design e Layout de Cards no Agente Salsa IT

Todo e qualquer cartão de resultado gerado no chat do Agente Salsa IT deve assumir estritamente a configuração compacta e harmoniosa do design system:

1. **Largura Contida (`wide: false`):**
   - Mensagens de resultado não devem usar `wide: true` por padrão. Devem manter largura contida (`max-width: 580px; width: 100%;` ou a restrição orgânica da bolha `max-width: min(78%, 720px);`), evitando esticar de ponta a ponta na janela com espaços vazios laterais.

2. **Estrutura Visual Padrão (`.asi-pfcg-result-card`):**
   - Chamar sempre `asiEnsurePfcgResultStyles()`.
   - **Cabeçalho na Mesma Linha:** Usar `.asi-pfcg-result-heading-row` contendo o título em `.asi-pfcg-result-heading` e o status badge em `.asi-pfcg-result-pill` (ex.: `.asi-pfcg-result-pill--success` ou `--warning`), eliminando quebras de linha e margens excessivas de `<h4>`.
   - **Subtítulo/Nota:** Usar `.asi-pfcg-result-note` com espaçamento vertical mínimo (`margin: 4px 0 6px 0;`).
   - **Grelha de Metadados:** Dispor campos-chave em grelhas compactas com `.asi-pfcg-result-grid` e `asiBuildPfcgResultField` (ex.: 2 a 3 colunas), sem vãos vazios.
   - **Blocos de Detalhes:** Transações, objetos ou listas devem ser apresentados em itens compactos estilizados (bordas suaves, fundo translúcido, tamanhos de fonte contidos `0.78rem - 0.82rem`), sem tags `<ul>`/`<li>` soltas com espaçamento padrão de navegador.

3. **Cache-Busting Obrigatório em Mudanças Frontend:**
   - Sempre que `cockpit.agent.js`, `cockpit.core.js` ou `styles.css` forem alterados, incrementar a versão nas tags `<script>` e `<link>` em `web_api/templates/index.html` (ex.: `?v=...`) e reiniciar o container Docker (`docker restart sap-script-web`).

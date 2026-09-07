# AGENTS.md

## Regra Operacional

- Sempre que forem feitas alterações no backend, reiniciar o Docker antes de validar o comportamento em runtime.
- Isso aplica-se a mudanças em `web_api`, `worker`, endpoints, jobs, handlers e qualquer lógica servida pelo container.
- Antes de pedir credenciais, host, cliente, sistema SAP ou dados RFC ao utilizador, verificar primeiro o `.env` da raiz `C:\workspace\SapScript\.env` e a configuração local do projeto.
- Se o `.env` já contiver a informação necessária, reutilizá-la diretamente e não solicitar novamente ao utilizador.

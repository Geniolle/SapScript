# Política de Runtime do SAP Cockpit Web

Este projeto usa um modelo fixo:

- `web_api` pode correr em Docker/Linux apenas como camada HTTP, fila e UI.
- A execução SAP real ocorre no worker Windows.
- O fluxo `FI` com `PyRFC` não deve ser executado em WSL/Linux.
- O cockpit não deve depender de `sys.executable` como fallback para passos de workflow.

## Execução correta

Use sempre caminhos explícitos para os executáveis Windows:

- `WORKFLOW_PYTHON_EXEC`
- `SAP_FI_BRIDGE_PYTHON`

Exemplo:

```env
WORKFLOW_PYTHON_EXEC=C:\workspace\SapScript\.venv\Scripts\python.exe
SAP_FI_BRIDGE_PYTHON=C:\workspace\SapScript\.venv-rfc\Scripts\python.exe
```

## Regra fixa

Se estas variáveis não existirem ou apontarem para outro runtime, o sistema deve falhar
de forma explícita. Não trocar o meio de execução por WSL, Docker ou `sys.executable`
fora de Windows para estes fluxos.

Quando o ambiente de teste estiver indisponível e a tarefa for apenas validar dados,
a ligação produtiva pode ser usada apenas em modo de leitura. Nessa situação:

- não criar job;
- não gerar proposta;
- não gravar movimento;
- não alterar estado funcional no SAP.

## Motivo

O processo `FI` depende de SAP NetWeaver RFC SDK e `pyrfc`, e no cockpit web o SAP GUI
vive no worker Windows. Manter a execução no mesmo meio reduz falhas intermitentes,
diferenças de encoding e erros de bridge como `UtilBindVsockAnyPort`.

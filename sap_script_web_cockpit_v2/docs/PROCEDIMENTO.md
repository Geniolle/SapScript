# Procedimento adotado

## Antes

```text
VSCode -> Run no ficheiro SAP cockpit -> menus no terminal -> SAP GUI
```

## Agora

```text
Pagina web -> cria job -> worker Windows -> run_sap_cockpit(payload) -> SAP GUI
```

## Primeiro teste recomendado

1. Subir o Docker.
2. Abrir o SAP GUI e fazer login no ambiente pretendido.
3. Iniciar o worker Windows.
4. Na web, executar primeiro `Ler STATUS atual do SAP`.
5. Depois executar `Abrir transacao` com `SE10`.
6. Por fim executar `Executar SAP Cockpit` preenchendo ambiente, processo e subprocesso.

## Campos minimos para executar o Cockpit

```json
{
  "ambiente": "S4Q",
  "processo": "NOME_DA_PASTA_DO_PROCESSO",
  "subprocesso": "NOME_DO_SCRIPT.py",
  "request_option": "4"
}
```

## Exemplo com request existente

```json
{
  "ambiente": "S4Q",
  "processo": "NOME_DA_PASTA_DO_PROCESSO",
  "subprocesso": "NOME_DO_SCRIPT.py",
  "request_option": "1",
  "request_number": "S4QK900396"
}
```

## Exemplo criando nova request

```json
{
  "ambiente": "S4Q",
  "processo": "NOME_DA_PASTA_DO_PROCESSO",
  "subprocesso": "NOME_DO_SCRIPT.py",
  "request_option": "2",
  "request_type": "1",
  "request_desc": "REQUEST CRIADA VIA WEB"
}
```

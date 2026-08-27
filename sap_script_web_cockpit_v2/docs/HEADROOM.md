# Headroom

Comandos úteis para usar o proxy local do Headroom com as ferramentas do projeto.

## Proxy

```powershell
$env:PATH += ";C:\Users\clayton.silva\AppData\Roaming\Python\Python312\Scripts"
headroom proxy --port 8787
```

## Antigravity

```text
agy
agy --dangerously-skip-permissions
```

## Codex

```powershell
$env:OPENAI_BASE_URL="http://127.0.0.1:8787/v1" ; codex --approve-for-me
```

## Claude

```powershell
$env:ANTHROPIC_BASE_URL="http://127.0.0.1:8787" ; claude --dangerously-skip-permissions
```

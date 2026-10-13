# MIGRAÇÃO DO SAP SCRIPT COCKPIT PARA ORACLE CLOUD (ALWAYS FREE)

## 1. Dados de Acesso e Credenciais do Servidor

### Acesso Web ao Cockpit (Navegador)
- **URL Direta:** [http://140.238.123.30](http://140.238.123.30)
- **Tipo de Autenticação:** HTTP Basic Auth (Protegido por Nginx)
- **Utilizador:** `Geniolle`
- **Senha:** `P@1geniolle.`

> **Nota de Acesso:** O Cockpit pode ser aberto em qualquer computador, telemóvel ou tablet. Ao aceder, o navegador solicitará o utilizador e a senha acima.

---

### Acesso SSH Administrativo (Servidor Linux)
- **IP Público:** `140.238.123.30`
- **Utilizador SSH:** `ubuntu`
- **Porta:** `22`
- **Sistema Operativo:** Ubuntu 24.04 LTS (Shape VM.Standard.E2.1.Micro)
- **Chave Privada SSH (Local Windows):**
  `C:\workspace\SapScript\Servidor\ssh-key-2026-10-13.key`

#### Comando de Ligação SSH:
```powershell
ssh -i "C:\workspace\SapScript\Servidor\ssh-key-2026-10-13.key" ubuntu@140.238.123.30
```

---

## 2. Visão Geral da Arquitetura

```
+-------------------------------------------------------------------------+
|                      ORACLE CLOUD (Always Free VM)                      |
|  IP: 140.238.123.30 | Ubuntu 24.04 LTS                                  |
|                                                                         |
|  [Nginx (Porta 80)]                                                     |
|    - Proteção HTTP Basic Auth (Geniolle)                                |
|    - Proxy reverso para http://127.0.0.1:8010                          |
|                                                                         |
|  [Docker Compose: sap-script-web]                                       |
|    - FastAPI / Uvicorn (porta interna: 8000)                            |
|    - Base de Dados SQLite persistente: /data/sap_script_jobs.sqlite3    |
|    - Volume Docker persistente: sap_script_cockpit_sap_script_data      |
|    - Limite de RAM: 450 MB | Swapfile: 1 GB                             |
+-------------------------------------------------------------------------+
                             ▲
                             │ HTTP de saída autenticado (X-Worker-Token)
                             ▼
+-------------------------------------------------------------------------+
|                       COMPUTADOR WINDOWS LOCAL                          |
|                                                                         |
|  - Worker Python: polling direto em http://140.238.123.30               |
|  - Executa as tarefas com SAP GUI Scripting / Automação desktop        |
|  - Docker local parado: 0% de CPU e 0 MB de RAM no Windows              |
+-------------------------------------------------------------------------+
```

---

## 3. Gestão e Manutenção do Worker Local

O Worker no Windows comunica diretamente com a nuvem.

### Como Iniciar o Worker:
No diretório `C:\workspace\SapScript\sap_script_web_cockpit_v2`:
```powershell
powershell -ExecutionPolicy Bypass -File .\start_all.ps1
```
*(O script garante que o processo do worker fica ativo em segundo plano).*

### Como Verificar se o Worker está a Correr:
```powershell
Get-Process python
```
Ou verificar o log em tempo real:
```powershell
Get-Content .\worker\worker_auto.log -Tail 20 -Wait
```

---

## 4. Comandos de Diagnóstico e Gestão no Servidor Oracle

Ao ligar-se via SSH (`ssh -i "C:\workspace\SapScript\Servidor\ssh-key-2026-10-13.key" ubuntu@140.238.123.30`):

- **Verificar estado do container:**
  ```bash
  docker ps
  ```
- **Ver consumo de CPU e RAM:**
  ```bash
  docker stats --no-stream
  ```
- **Ver logs do Cockpit em tempo real:**
  ```bash
  docker logs --tail 50 -f sap-script-web
  ```
- **Reiniciar o serviço do Cockpit:**
  ```bash
  cd ~/sap_script_cockpit && docker compose restart
  ```
- **Reiniciar o Nginx:**
  ```bash
  sudo systemctl restart nginx
  ```
- **Fazer backup manual da base de dados:**
  ```bash
  docker exec sap-script-web python -c "import sqlite3; s=sqlite3.connect('/data/sap_script_jobs.sqlite3'); b=sqlite3.connect('/data/backup_manual.sqlite3'); s.backup(b); b.close(); s.close()"
  ```

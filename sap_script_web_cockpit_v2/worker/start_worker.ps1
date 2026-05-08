$ErrorActionPreference = "Stop"

cd "C:\workspace\sap-script\sap_script_web_cockpit_v2\worker"

& ".\.venv\Scripts\Activate.ps1"

$env:API_BASE_URL = "http://localhost:8010"
$env:WORKER_TOKEN = "change-me"
$env:SAP_SCRIPT_PROJECT_DIR = "C:\workspace\sap-script"
$env:SAP_COCKPIT_MODULE = "sap_script_web_cockpit_v2.sap_cockpit_web_ready"
$env:POLL_SECONDS = "1"

python worker.py

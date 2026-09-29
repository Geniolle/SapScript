#!/usr/bin/env python3
import os, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from dotenv import load_dotenv
load_dotenv()

from pyrfc import Connection

configs = {
    'DEV': (os.getenv('SAP_DEV_ASHOST'), os.getenv('SAP_DEV_SYSNR'), os.getenv('SAP_DEV_CLIENT'), os.getenv('SAP_DEV_USER'), os.getenv('SAP_DEV_PASSWD'), '172.19.66.4'),
    'QAD': (os.getenv('SAP_QAD_ASHOST'), os.getenv('SAP_QAD_SYSNR'), os.getenv('SAP_QAD_CLIENT'), os.getenv('SAP_QAD_USER'), os.getenv('SAP_QAD_PASSWD'), '172.19.66.22'),
    'PRD': (os.getenv('SAP_PRD_ASHOST'), os.getenv('SAP_PRD_SYSNR'), os.getenv('SAP_PRD_CLIENT'), os.getenv('SAP_PRD_USER'), os.getenv('SAP_PRD_PASSWD'), '172.19.34.22'),
}

req_id = 'S4DK953658'
print(f"Procurando Request {req_id}...\n")

for env, (host, sysnr, client, user, passwd, ip) in configs.items():
    print(f"📡 {env} ({ip})...", end=" ")
    try:
        conn = Connection(user=user, passwd=passwd, ashost=host, sysnr=sysnr, client=client)
        result = conn.call('TR_READ_REQUEST', REQUEST=req_id)
        conn.close()

        if result and result.get('TRKORR'):
            print("✅ ENCONTRADA!")
            print(f"   Status: {result.get('TRSTATUS', '?')}")
            print(f"   Utilizador: {result.get('STRUSER', '?')}")
            print(f"   Tipo: {result.get('TRTYPE', '?')}")
            print(f"   Sistema origem: {result.get('TRSYSTEM', '?')}")
            sys.exit(0)
        else:
            print("❌")
    except Exception as e:
        print(f"❌ ({str(e)[:30]})")

print(f"\n❌ Request não encontrada")

#!/usr/bin/env python3
"""
Teste simples de conexão RFC ao SAP PRD.
Lê credenciais do .env local.
"""

import os
import sys
from pathlib import Path

# Adiciona raiz do projeto ao path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

# Carrega .env
load_dotenv(project_root / ".env")

def test_prd_connection():
    """Testa conexão ao SAP PRD via RFC."""

    # Recolhe credenciais
    config = {
        "user": os.getenv("SAP_PRD_USER"),
        "passwd": os.getenv("SAP_PRD_PASSWD"),
        "ashost": os.getenv("SAP_PRD_ASHOST"),
        "sysnr": os.getenv("SAP_PRD_SYSNR"),
        "client": os.getenv("SAP_PRD_CLIENT"),
        "lang": os.getenv("SAP_PRD_LANG", "PT"),
    }

    # Validação básica
    required = ["user", "passwd", "ashost", "sysnr", "client"]
    missing = [k for k in required if not config.get(k)]
    if missing:
        print(f"❌ Credenciais em falta: {', '.join(missing)}")
        return False

    print("📋 Configuração PRD carregada:")
    print(f"  User: {config['user']}")
    print(f"  Host: {config['ashost']}")
    print(f"  SysNr: {config['sysnr']}")
    print(f"  Client: {config['client']}")
    print(f"  Lang: {config['lang']}")
    print()

    try:
        print("🔄 A conectar ao SAP PRD via RFC...")
        from pyrfc import Connection

        conn = Connection(**config)

        print("✅ Conexão estabelecida com sucesso!")
        print()

        # Teste simples: chamar uma RFC
        print("📞 A chamar RFC_SYSTEM_INFO...")
        result = conn.call("RFC_SYSTEM_INFO")

        print("✅ RFC_SYSTEM_INFO respondeu:")
        print(f"  Sistema: {result.get('RFCHOST', 'N/A')}")
        print(f"  SysNr: {result.get('RFCSYSID', 'N/A')}")
        print(f"  Versão: {result.get('RFCRELEASE', 'N/A')}")
        print()

        conn.close()
        print("✅ Teste de conexão PRD concluído com sucesso!")
        return True

    except ModuleNotFoundError as e:
        print(f"❌ Erro de módulo: {e}")
        print("   Certifique-se de que PyRFC está instalado: pip install pyrfc")
        return False
    except Exception as e:
        error_class = e.__class__.__name__
        error_msg = str(e)
        print(f"❌ Erro de conexão ({error_class}):")
        print(f"   {error_msg}")
        return False

if __name__ == "__main__":
    success = test_prd_connection()
    sys.exit(0 if success else 1)

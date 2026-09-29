"""
Diagnóstico da Conexão RFC e Consulta EKKO
===========================================

OBJETIVO: Descobrir por que o código Python/RFC não consegue encontrar a PO 4300000002
que EXISTE em SE16H (S4Q, mandante 100, EKKO-EBELN = 4300000002).

ETAPA 1: Confirmar a conexão RFC efetiva
ETAPA 2: Consultar EKKO com EBELN = '4300000002' apenas
ETAPA 3: Mostrar resultado bruto
ETAPA 4: Diagnosticar o problema
ETAPA 5: Corrigir e validar
"""

from __future__ import annotations

import os
import sys
import json
from pathlib import Path
from typing import Any

try:
    from pyrfc import Connection
except Exception as exc:
    print(f"ERRO: PyRFC não disponível. {exc}")
    sys.exit(1)

# Setup do projeto
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
load_dotenv(project_root / ".env", override=False)


class RFCDiagnostic:
    """Diagnóstico detalhado da conexão RFC e consultas."""

    def __init__(self, environment: str = "QAD"):
        self.environment = environment
        self.conn = None
        self.connection_info = {}

    def stage_1_confirm_connection(self):
        """ETAPA 1: Confirma a conexão RFC efetiva."""

        print("\n" + "="*80)
        print("ETAPA 1: CONFIRMAR A CONEXÃO RFC")
        print("="*80)

        env_prefix = self.environment.upper()

        # Lê as credenciais do .env
        params = {
            "user": os.getenv(f"SAP_{env_prefix}_USER", "").strip(),
            "passwd": os.getenv(f"SAP_{env_prefix}_PASSWD", "").strip(),
            "ashost": os.getenv(f"SAP_{env_prefix}_ASHOST", "").strip(),
            "sysnr": os.getenv(f"SAP_{env_prefix}_SYSNR", "").strip(),
            "client": os.getenv(f"SAP_{env_prefix}_CLIENT", "").strip(),
            "lang": os.getenv(f"SAP_{env_prefix}_LANG", "EN").strip(),
        }

        print(f"\nParâmetros configurados para '{env_prefix}':")
        print(f"  ASHOST: {params['ashost']}")
        print(f"  SYSNR:  {params['sysnr']}")
        print(f"  CLIENT: {params['client']}")
        print(f"  USER:   {params['user']}")
        print(f"  LANG:   {params['lang']}")
        print(f"  PASSWORD: [OCULTADA]")

        if not all([params['user'], params['passwd'], params['ashost'], params['sysnr'], params['client']]):
            print("\n✗ ERRO: Credenciais incompletas no .env")
            raise RuntimeError(f"Credenciais incompletas para {env_prefix}")

        try:
            print("\n  Conectando ao SAP...")
            self.conn = Connection(**params)
            print("  ✓ Conexão estabelecida")

            # RFC para obter informações do sistema
            result = self.conn.call("RFC_PING")
            print("  ✓ RFC_PING bem-sucedido")

            # Tenta obter info do sistema
            try:
                sys_info = self.conn.call("RFC_GET_SYSTEM_INFO")
                print(f"\n  Informações do Sistema SAP (via RFC_GET_SYSTEM_INFO):")
                for key, value in sys_info.items():
                    if key not in ['PARTNER', 'TIME']:
                        print(f"    {key}: {value}")
                self.connection_info = sys_info
            except Exception as e:
                print(f"  ⚠️  RFC_GET_SYSTEM_INFO não disponível: {e}")

            # Tenta obter info via ABAP_GET_SYSINFO
            try:
                abap_info = self.conn.call("ABAP_GET_SYSINFO")
                print(f"\n  Informações do Sistema SAP (via ABAP_GET_SYSINFO):")
                if 'SYSINFO' in abap_info:
                    sysinfo = abap_info['SYSINFO']
                    if isinstance(sysinfo, dict):
                        for key, value in sysinfo.items():
                            print(f"    {key}: {value}")
                    else:
                        print(f"    {sysinfo}")
            except Exception as e:
                print(f"  ⚠️  ABAP_GET_SYSINFO não disponível: {e}")

            print(f"\n  ✓ Conectado ao ambiente: {self.environment}")
            print(f"  ✓ ASHOST: {params['ashost']}")
            print(f"  ✓ SYSNR: {params['sysnr']}")
            print(f"  ✓ CLIENT: {params['client']}")

            return True

        except Exception as e:
            print(f"\n✗ ERRO ao conectar: {e}")
            raise

    def stage_2_query_ekko_simple(self):
        """ETAPA 2: Consulta EKKO com apenas EBELN = '4300000002'."""

        print("\n" + "="*80)
        print("ETAPA 2: CONSULTAR EKKO (EBELN = '4300000002')")
        print("="*80)

        if not self.conn:
            raise RuntimeError("Não conectado")

        # CAMPOS SOLICITADOS
        fields = [
            "EBELN",    # Número do documento de compras
            "BUKRS",    # Empresa
            "BSART",    # Tipo de documento
            "LIFNR",    # Fornecedor
            "WAERS",    # Moeda
            "WKURS",    # Taxa de câmbio
            "KUFIX",    # Indicador de taxa fixa
            "LOEKZ",    # Indicador de exclusão
            "BEDAT",    # Data do documento
            "AEDAT",    # Data de última alteração
        ]

        # CONDIÇÃO EXATA
        options = ["EBELN = '4300000002'"]

        print(f"\nTabela: EKKO")
        print(f"Condição WHERE: {options[0]}")
        print(f"Campos solicitados: {fields}")
        print(f"ROWCOUNT: 0 (sem limite)")

        try:
            print("\n  Executando RFC_READ_TABLE...")

            result = self.conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="EKKO",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": f} for f in fields],
                OPTIONS=[{"TEXT": o} for o in options],
                ROWCOUNT=0,
            )

            print("  ✓ RFC_READ_TABLE executada com sucesso")

            # Processa resultado
            raw_data = result.get("DATA", [])

            print(f"\n  Resultado Bruto:")
            print(f"    Total de linhas retornadas: {len(raw_data)}")

            if raw_data:
                print(f"\n  ✓ SUCESSO! PO encontrada!")
                for i, item in enumerate(raw_data):
                    print(f"\n    Linha {i+1}:")
                    wa_raw = str(item.get("WA", ""))
                    print(f"      WA (raw): {wa_raw}")

                    # Separa campos
                    parts = [p.strip() for p in wa_raw.split("|")]
                    print(f"\n      Campos parseados:")
                    for j, field in enumerate(fields):
                        value = parts[j] if j < len(parts) else "[vazio]"
                        print(f"        {field}: {value}")
            else:
                print(f"\n    ✗ Nenhuma linha encontrada!")

            return raw_data

        except Exception as e:
            print(f"\n✗ ERRO ao executar RFC_READ_TABLE:")
            print(f"  Tipo: {type(e).__name__}")
            print(f"  Mensagem: {e}")
            if hasattr(e, 'key'):
                print(f"  Key: {e.key}")
            if hasattr(e, 'code'):
                print(f"  Code: {e.code}")
            raise

    def stage_3_test_control(self):
        """ETAPA 4: Teste de controlo - consulta por BUKRS = '2010'."""

        print("\n" + "="*80)
        print("ETAPA 4: TESTE DE CONTROLO (BUKRS = '2010')")
        print("="*80)

        if not self.conn:
            raise RuntimeError("Não conectado")

        fields = ["EBELN", "BUKRS", "LIFNR"]
        options = ["BUKRS = '2010'"]

        print(f"\nTabela: EKKO")
        print(f"Condição WHERE: {options[0]}")
        print(f"Campos: {fields}")

        try:
            result = self.conn.call(
                "RFC_READ_TABLE",
                QUERY_TABLE="EKKO",
                DELIMITER="|",
                FIELDS=[{"FIELDNAME": f} for f in fields],
                OPTIONS=[{"TEXT": o} for o in options],
                ROWCOUNT=50,
            )

            raw_data = result.get("DATA", [])

            print(f"\n  Total de linhas encontradas: {len(raw_data)}")

            # Procura por 4300000002
            found_po = False
            for item in raw_data:
                wa_raw = str(item.get("WA", ""))
                parts = [p.strip() for p in wa_raw.split("|")]
                if len(parts) > 0 and parts[0] == "4300000002":
                    found_po = True
                    print(f"\n  ✓ PO 4300000002 ENCONTRADA na busca por BUKRS = '2010'!")
                    print(f"    EBELN: {parts[0]}")
                    print(f"    BUKRS: {parts[1] if len(parts) > 1 else 'N/A'}")
                    print(f"    LIFNR: {parts[2] if len(parts) > 2 else 'N/A'}")
                    break

            if not found_po:
                print(f"\n  ✗ PO 4300000002 NÃO foi encontrada na busca por BUKRS = '2010'")
                print(f"  Primeiras 10 POs encontradas:")
                for i, item in enumerate(raw_data[:10]):
                    wa_raw = str(item.get("WA", ""))
                    parts = [p.strip() for p in wa_raw.split("|")]
                    print(f"    {parts[0] if len(parts) > 0 else 'N/A'}")

        except Exception as e:
            print(f"\n✗ ERRO: {e}")

    def close(self):
        """Fecha conexão."""
        if self.conn:
            try:
                self.conn.close()
                print("\n✓ Conexão RFC fechada")
            except:
                pass


def main():
    diagnostic = None
    try:
        diagnostic = RFCDiagnostic(environment="QAD")

        # ETAPA 1: Confirma conexão
        diagnostic.stage_1_confirm_connection()

        # ETAPA 2: Consulta EKKO
        result = diagnostic.stage_2_query_ekko_simple()

        # ETAPA 3: Teste de controlo
        diagnostic.stage_3_test_control()

        # CONCLUSÃO
        print("\n" + "="*80)
        print("RESUMO DO DIAGNÓSTICO")
        print("="*80)

        if result:
            print("\n✓ PO 4300000002 foi ENCONTRADA em EKKO via RFC")
        else:
            print("\n✗ PO 4300000002 NÃO foi encontrada em EKKO via RFC")
            print("\nPróximas investigações necessárias:")
            print("  1. Verificar se a conexão é realmente S4Q/mandante 100")
            print("  2. Verificar a configuração do .env para SAP_QAD_*")
            print("  3. Verificar se há autorizações RFC adequadas")
            print("  4. Verificar se há filtros adicionais na função RFC_READ_TABLE")

    except Exception as e:
        print(f"\n✗ ERRO FATAL: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if diagnostic:
            diagnostic.close()


if __name__ == "__main__":
    main()

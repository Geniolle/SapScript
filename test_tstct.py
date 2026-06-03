import sys
sys.path.insert(0, r'c:\workspace\sap-script')
from sap_agent.sap_rfc_client import SapRfcClient
from sap_agent.config import SapConnectionConfig
from dotenv import load_dotenv

load_dotenv(r'c:\workspace\sap-script\.env')
config = SapConnectionConfig.from_env()
client = SapRfcClient(config)
try:
    res = client.read_table('TSTCT', options=[{'TEXT': "TCODE = 'ZFI_FORW_PO_STOCK_PE'"}], fields=['TTEXT'])
    print('TSTCT:', res)
except Exception as e:
    print('Error:', e)

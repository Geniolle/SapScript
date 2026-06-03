import sys
sys.path.insert(0, r'c:\workspace\sap-script')
from sap_agent.sap_rfc_client import SapRfcClient
from sap_agent.config import SapConnectionConfig
from dotenv import load_dotenv

load_dotenv(r'c:\workspace\sap-script\.env')
config = SapConnectionConfig.from_env()
client = SapRfcClient(config)
try:
    res = client.read_table('DOKTL', options=[{'TEXT': "ID = 'RE'"}, {'TEXT': " AND OBJECT = 'ZFI_FORW_POS_PO_STOCK'"}], fields=[])
    print('DOKTL:', len(res), 'rows')
    if res: print(res[0])
except Exception as e:
    print('Error:', e)

import os
from datetime import datetime, timedelta
import logging
import json
from v2.scripts import read_json

# read the JSON data
data = read_json()

payment_documents = data.get('request', [])
fe = payment_documents[0].get('0', '')
registry_number = fe.get('registry_name', '').strip('РЕЕСТР ПЛАТЕЖЕЙ №')
print(registry_number)

date_str = datetime.now().strftime('%d/%m/%Y')
filename = f"Реестр от {date_str}.xlsx"

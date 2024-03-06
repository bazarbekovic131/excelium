import os
from datetime import datetime, timedelta
import logging

date_str = datetime.now().strftime('%d/%m/%Y')
filename = f"Реестр от {date_str}.xlsx"

print(filename)

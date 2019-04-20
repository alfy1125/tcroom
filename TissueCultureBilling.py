import pandas as pd
import numpy as ny
import random
import os
from pandas import ExcelWriter
from pandas import ExcelFile

print (os.getcwd())
print(os.listdir(os.getcwd()))

excel = pd.ExcelFile('MarchTC.xlsx')
sheet = pd.read_excel(excel)
global bill_hours
bill_hours = []
companies = ['Company A','Company B','Company C']

def weekly_run():
    for x in companies:
        results = sheet.loc[sheet['Organization'] == x]
        total_hours = float(sum(results['Hours']))
        bill_hours.append(total_hours)

weekly_run()
df = pd.DataFrame({'Company': companies, 'Hours': bill_hours})
writer = pd.ExcelWriter('MonthlyTCBilling.xlsx')
df.to_excel(writer, 'TC Billing', index=False)
writer.save()


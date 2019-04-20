import pandas as pd
import numpy as ny
import random
import os
from pandas import ExcelWriter
from pandas import ExcelFile

print (os.getcwd())
print(os.listdir(os.getcwd()))
global bill_hours
bill_hours=[]

pd.options.display.max_rows = 999

xls=pd.ExcelFile('conference.xlsx')
file=pd.read_excel(xls,'Sheet')

users=file['Email'].str.split("@",expand= True)
users.columns=['users','domain']
unique=users.domain.unique()
print(unique)


def weekly_run(x):
    for options in unique:
        file1 = pd.read_excel(xls, str(x))
        results = file1[file1['Email'].str.contains(options)]
        total_hours=(sum(results['Hours']))
        bill_hours.append(total_hours)


writer = pd.ExcelWriter('incidental.xlsx')

def multi_sheet():
    for i in range(1,5):
        weekly_run(str(i))
        d1=pd.DataFrame({'Company':unique,'Hours':bill_hours})
        d1.to_excel(writer,sheet_name=str(i))
        bill_hours.clear()

multi_sheet()
writer.save()

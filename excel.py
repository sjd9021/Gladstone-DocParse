import openpyxl
from tess import our_ref, branch, insurers, date, total_amount, policy
import pandas as pd
c = 0
df = pd.read_excel('gal.xlsx')
len = len(df['SL NO'])
a = df.iloc[:, 1]
# finish how to find so.no
serial = df.iloc[0:10]
print(serial)
for i in a:
   if our_ref == i:
      c = 1

if c == 0:
    wb = openpyxl.load_workbook('gal.xlsx')
# add so.no, change font and size and centre allign
    ws = wb.active
    newRowLocation = ws.max_row +1
    ws.cell(column=2,row=newRowLocation, value=our_ref)
    ws.cell(column=4,row=newRowLocation, value=insurers)
    ws.cell(column=5,row=newRowLocation, value=branch)
    ws.cell(column=6,row=newRowLocation, value=policy)
    ws.cell(column=7,row=newRowLocation, value=date)
    ws.cell(column=9, row=newRowLocation, value=total_amount)
    ws.cell(column=10, row=newRowLocation, value=total_amount)

wb.save('gal.xlsx')
wb.close()

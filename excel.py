
import openpyxl
from tess import our_ref
from os import path
# to be completed,this is prototype



wb = openpyxl.Workbook()
sheet = wb['Sheet']
data = [
    (1, 'Our_Ref'),
    (2, our_ref)
]
for x in data:
    sheet.append(x)
wb.save('glad.xlsx')


import random
import string
import os.path

import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=12.0.0.0, Culture=Neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel


def random_data(max_len):
    symbols = string.ascii_letters + string.digits + ' '
    return ''.join(random.choice(symbols) for i in range(random.randrange(max_len)))

number_of_groups = 5
excel_data = 'data\groups.xlsx'

test_data = [random_data(10) for i in range(number_of_groups)]
file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', excel_data)

if os.path.exists(os.path.join('..', excel_data)):
    os.remove(os.path.join('..', excel_data))

excel = Excel.ApplicationClass()
excel.Visible = False
workbook = excel.WorkBooks.Add()
sheet = workbook.ActiveSheet

for i in range(len(test_data)):
    sheet.Range['A%s' % (i+1)].Value2 = test_data[i]

workbook.SaveAs(file)
excel.Quit()

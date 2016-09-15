from model.group import Group
import random
import string
import os.path
import getopt
import sys
import time

import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=12.0.0.0, Culture=Neutral, PublicKeyToken=71e9bce111e9429c')

from Microsoft.Office.Interop import Excel

try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of groups", "file"])
except getopt.GetoptError as error:
    getopt.usage()
    sys.exit(2)

n = 2
f = "data\groups.xlsx"

for o, a in opts:
    if o == "-n":
        n = int(a)
    elif o == "-f":
        f = a


def random_data(prefix, max_len):
    symbols = string.ascii_letters + string.digits + string.punctuation + " "*10
    return prefix + "".join(random.choice(symbols) for i in range(random.randrange(max_len)))


test_data = [Group(name="")] + \
            [Group(name=random_data("name", 10)) for i in range(n)]

file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", f)


excel = Excel.ApplicationClass()
excel.Visible = True

if os.path.exists(os.path.join('..', f)):
    os.remove(os.path.join('..', f))

workbook = excel.WorkBooks.Add()
sheet = workbook.ActiveSheet

for i in range(len(test_data)):
    sheet.Range['A%s' % (i+1)].Value2 = test_data[i].name

workbook.SaveAs(file)

time.sleep(5)
excel.Quit()

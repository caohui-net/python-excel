from openpyxl import *

testpath = r'C:\Users\Administrator\Documents\GitHub\python-sql\python-excel\study_openpyxl\tt.xlsx'

book = load_workbook(testpath)
names = book.sheetnames
sheet = book[names[1]]
print(len(names))
for i in names:
    sheet = book[i]
    print(sheet.title)


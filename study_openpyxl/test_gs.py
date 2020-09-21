from openpyxl import *

testpath = r'C:\Users\Administrator\Documents\GitHub\python-sql\python-excel\study_openpyxl\test_gs.xlsx'
book = Workbook()
sheet = book.active

sheet['a1'] = 10
sheet['a2'] = 20
sheet['a3'] ='=sum(a1:a2)'

book.save(testpath)
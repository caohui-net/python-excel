from openpyxl import *
import time

test_path=r'C:\Users\Administrator\Documents\GitHub\python-sql\python-excel\study_openpyxl\sample.xlsx'
book = Workbook()
sheet = book.active

sheet['a1'] = 56
sheet['a2'] = 43

now = time.strftime("%x")
sheet['a3'] = now

book.save(test_path)
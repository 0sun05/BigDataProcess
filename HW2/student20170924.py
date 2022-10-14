#!/usr/bin/python3

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

wb = Workbook()
dest_filename = 'student.xlsx'

ws = wb.active
ws.title = 'Sheet1'

ws.append(['id', 'name', 'midterm', 'final', 'homework', 'attendance',
	'total', 'grade'])
ws.append(['20140001', 'Sophia', 23, 53, 41, 1])
ws.append(['20140002', 'Emily', 94, 36, 33, 1])
ws.append(['20140003', 'Lily', 37, 20, 46, 1])
ws.append(['20140004', 'Olivia', 73, 100, 72, 1])
ws.append(['20140005', 'Amelia', 93, 46, 0, 1])
ws.append(['20150001', 'Isla', 6, 30, 58, 1])
ws.append(['20150003', 'Isabella', 71, 51, 54, 1])
ws.append(['20150005', 'Ava', 43, 62, 56, 1])
ws.append(['20150007', 'Sophie', 48, 92, 14, 1])
ws.append(['20150009', 'Chloe', 91, 64, 39, 1])

for i in range(2, 12):
	ws.cell(i, 7, value = ws.cell(i, 3).value * 0.3 +
	ws.cell(i, 4).value * 0.35 + ws.cell(i, 5).value * 0.34 + 1)

wb.save(filename = dest_filename)

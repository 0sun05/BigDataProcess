#!/usr/bin/python3
import openpyxl

wb = openpyxl.load_workbook("student.xlsx")
ws = wb['Sheet1']

i = 1
for row in ws:
	if i != 1:
		ws.cell(row = i, column = 7).value = \
		ws.cell(row = i, column = 3).value * 0.3 + \
		ws.cell(row = i, column = 4).value * 0.35 + \
		ws.cell(row = i, column = 5).value * 0.34 + \
		ws.cell(row = i, column = 6).value
	i += 1

wb.save("student.xlsx")

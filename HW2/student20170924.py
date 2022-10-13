#!/usr/bin/python3

from openpyxl import Workbook
wb = Workbook()

ws = wb.active
ws['A1'] = "id"

wb.save("student.xlsx")

from copy import copy
import openpyxl as xl

filePath = './test-data/test-data.xlsx'

# Book
wb = xl.load_workbook(filePath)

# Sheet
ws1 = wb['Test1']
ws2 = wb['Test2']

# Cell
ws1c3 = ws1['C3']
ws2c3 = ws2['C3']

# Copy style
ws2c3.font = copy(ws1c3.font)
ws2c3.border = copy(ws1c3.border)
ws2c3.fill = copy(ws1c3.fill)
ws2c3.number_format = copy(ws1c3.number_format)
ws2c3.protection = copy(ws1c3.protection)
ws2c3.alignment = copy(ws1c3.alignment)

# Saving to a file
wb.save(filePath)

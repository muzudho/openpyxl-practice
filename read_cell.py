import openpyxl as xl

# Book
wb = xl.load_workbook('test-data/test-data.xlsx')

# Sheet
ws = wb['Test1']

# Cell
c3 = ws['C3']

# Cell value
print(f'cell(value={c3.value})')

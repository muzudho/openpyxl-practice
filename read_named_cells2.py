import openpyxl as xl

# Book
wb = xl.load_workbook('test-data/test-data.xlsx')
print(f'Sheet names| {wb.sheetnames}')

# 名前付き範囲
tileMap = wb.defined_names['TileMap']
print(f'tileMap(attr_text={tileMap.attr_text})')

# 飛び地一覧
print('''Enclaves
--------''')
for sheetName, coord in tileMap.destinations:
    print(f'(sheetName={sheetName} coord={coord})')

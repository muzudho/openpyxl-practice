import openpyxl as xl

# Book
wb = xl.load_workbook('test-data/test-data.xlsx')
print(f'Sheet names| {wb.sheetnames}')

# 名前付き範囲
tileMap = wb.defined_names['TileMap']
print(f'tileMap.attr_text| {tileMap.attr_text}')

tableList = [wb[s][r] for s, r in tileMap.destinations]
print(f'tableList| {tableList}')

for rowsTuple in tableList:
    print(f'rowsTuple| {rowsTuple}')
    for columnsTuple in rowsTuple:
        print(f'columnsTuple| {columnsTuple}')
        for cell in columnsTuple:
            print(f'cell(row={cell.row} column={cell.column} coordinate={cell.coordinate} value={cell.value})')

import openpyxl as xl

# Book
wb = xl.load_workbook('test-data/test-data.xlsx')
print(f'Sheet names| {wb.sheetnames}')

# 名前付き範囲
tileMap = wb.defined_names['TileMap']
print(f'tileMap.attr_text| {tileMap.attr_text}')

# 飛び地を集めています。リストの中に、ネストしたタプルが入っています
bookView = [wb[s][r] for s, r in tileMap.destinations]
print(f'bookView| {bookView}')

# 各セルに１つずつ訪れます
for sheetView in bookView:
    print(f'sheetView| {sheetView}')
    for rowView in sheetView:
        print(f'rowView| {rowView}')
        for cell in rowView:
            print(f'cell(row={cell.row} column={cell.column} coordinate={cell.coordinate} value={cell.value})')

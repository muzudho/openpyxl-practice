import openpyxl as xl

# Book
wb = xl.load_workbook('test-data/test-data.xlsx')

# Sheet
ws = wb['Test3']

def vlookup(key, keyColumnAlphabet, resultColumnAlphabet):
    """A列から 3 を探します
    """
    for rowNumber in range(2,10):
        id = ws[f'{keyColumnAlphabet}{rowNumber}'].value
        print(f'id={id}')
        if id == key:
            return ws[f'{resultColumnAlphabet}{rowNumber}']
        elif id is None or id == '':
            return None

    return None

result = vlookup(3, 'A', 'B')
print(f'found={result.value}')

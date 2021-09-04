import openpyxl as xl
# from openpyxl.styles.colors import Color
# from openpyxl.styles.fonts import Font
# from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side

def colorToDescription(name, color):
    """色にはいくつか種類があるので、対応します
    Parameters
    ----------
    color : Color
        色オブジェクト
    """
    if color.type=='theme':
        return f'{name}(theme={color.theme} tint={color.tint})'
    elif color.type=='indexed':
        return f'{name}(indexed={color.indexed})'
    elif color.type=='rgb':
        return f'{name}(rgb={color.rgb})'
    else:
        return f'{name}(type={color.type})'

def boarderSideToDescription(name, side):
    """上下左右にあるタイプの境界線を説明します
    Parameters
    ----------
    side : openpyxl.styles.borders.Side
        上下左右にあるタイプの境界線オブジェクト
    """
    if side.style=='thick':
        s = colorToDescription('color', side.color)
        return f'{name}(thick {s})'
    elif side.style=='thin':
        s = colorToDescription('color', side.color)
        return f'{name}(thin {s})'
    elif side.style=='medium':
        s = colorToDescription('color', side.color)
        return f'{name}(medium {s})'
    elif not(side.style is None):
        return f'{name}(style={side.style})'
    else:
        return ''

filePath = './test-data/test-data.xlsx'

# Book
wb = xl.load_workbook(filePath)

# 名前付き範囲
tileMap = wb.defined_names['TileMap2']

tableList = [wb[s][r] for s, r in tileMap.destinations]

for rowsTuple in tableList:
    for columnsTuple in rowsTuple:
        for cell in columnsTuple:
            # 値に 1 を足します
            cell.value += 1

            # 文字色を塗ります
            cell.font.color.rgb='00CC3333'

            # 背景色を塗ります
            cell.fill.fgColor.rgb = '00CCFFFF'

            # セルの上辺に線を引きます
            side = Side(style='thin', color='000000')
            cell.border = Border(top=side, bottom=None, left=None, right=None)

# Saving to a file
wb.save(filePath)

from openpyxl import load_workbook
from openpyxl.styles.colors import Color
from openpyxl.styles.fonts import Font

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
wb = load_workbook(filePath)

# 名前付き範囲
tileMap = wb.defined_names['TileMap2']

tableList = [wb[s][r] for s, r in tileMap.destinations]

for rowsTuple in tableList:
    for columnsTuple in rowsTuple:
        for cell in columnsTuple:
            # 値に 1 を足します
            cell.value += 1

            # 文字を赤色にします
            cell.font.color.rgb='00FF0000'

# Saving to a file
wb.save(filePath)

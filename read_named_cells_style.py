import openpyxl as xl
from openpyxl.styles.colors import Color

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

# Book
wb = xl.load_workbook('test-data/test-data.xlsx')

# 名前付き範囲
tileMap = wb.defined_names['TileMap']

tableList = [wb[s][r] for s, r in tileMap.destinations]

for rowsTuple in tableList:
    for columnsTuple in rowsTuple:
        for cell in columnsTuple:
            # 値
            print(f'|cell.value|{cell.value:2}|',end='')
            # フォント色
            s = colorToDescription('cell.font.color', cell.font.color)
            print(f'{s}|',end='')
            if cell.fill.patternType=='solid':
                # print(f'cell.fill|{cell.fill}|',end='')
                # フィル前景色
                s = colorToDescription('cell.fill.fgColor', cell.fill.fgColor)
                print(f'{s}|',end='')
                # フィル背景色
                s = colorToDescription('cell.fill.bgColor', cell.fill.bgColor)
                print(f'{s}|',end='')
            else:
                print(f'cell.fill.patternType|{cell.fill.patternType}|',end='')

            # 境界線
            # print(f'|cell.border|{cell.border}|',end='')
            # いろいろあるがとりあえずいくつか取る
            s = boarderSideToDescription('cell.border.left', cell.border.left)
            if s!='':
                print(f'{s}|',end='')
            s = boarderSideToDescription('cell.border.right', cell.border.right)
            if s!='':
                print(f'{s}|',end='')
            s = boarderSideToDescription('cell.border.top', cell.border.top)
            if s!='':
                print(f'{s}|',end='')
            s = boarderSideToDescription('cell.border.bottom', cell.border.bottom)
            if s!='':
                print(f'{s}|',end='')

            # 改行
            print('')

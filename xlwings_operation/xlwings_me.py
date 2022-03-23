import xlwings as xw
from importlib import import_module

# 数字とアルファベットを変換
convert_alphabet_to_num = import_module("convert_alphabet_to_num")
convert_num_to_alphabet = import_module("convert_num_to_alphabet")

# 1行目の最大列を取得して、その全ての列から最大行を返す
def all(sheet):
    max_col = sheet.range(1, sheet.cells.last_cell.column).end("left").column

    last_row = 1
    for i in range(max_col + 1)[1:]:
        max_row = sheet.range(sheet.cells.last_cell.row, i).end("up").row
        if max_row > last_row:
            last_row = max_row
    return last_row


# 1行目の最大列を取得して、1行目の最大列とその全ての列から最大行を返す
def last(sheet):
    last_col = sheet.range(1, sheet.cells.last_cell.column).end("left").column

    last_row = 1
    for i in range(last_col + 1)[1:]:
        max_row = sheet.range(sheet.cells.last_cell.row, i).end("up").row
        if max_row > last_row:
            last_row = max_row

    return last_col, last_row


# 指定されたcolの最大行を返す
def col(sheet, col):
    last_row = sheet.range(sheet.cells.last_cell.row, convert_alphabet_to_num.A2num(col)).end("up").row
    return last_row


# 最大行と最大列、列基準のリストを返す(縦向きのリスト)
def all_col(sheet, colstart: int = 0, colend: int = 0):
    last_col, last_row = last(sheet)

    if colend == 0:
        colend = last_col

    list_all: list = []
    for col in range(int(last_col))[colstart:colend]:
        strcol = convert_num_to_alphabet.num2A(col + 1)
        list_all.append(sheet.range(strcol + "1:" + strcol + str(last_row)).value)

    return last_col, last_row, list_all


# 最大行と最大列、行基準のリストを返す(横向きのリスト)
def all_row(sheet, rowstart: int = 0, rowend: int = 0):
    last_col, last_row = last(sheet)

    if rowend == 0:
        rowend = last_row

    strlastcol = convert_num_to_alphabet.num2A(last_col)

    list_all: list = []
    for row in range(int(last_row))[rowstart:rowend]:
        list_all.append(sheet.range("A" + str(row + 1) + ":" + str(strlastcol) + str(row + 1)).value)

    return last_col, last_row, list_all


# Windows環境のみ メッセージボックスを表示する
import win32api  # 参考:http://housoubu.mizusasi.net/data/prog/p003.html
import win32con


def MessageBox(application, alert="エラーが発生しました", title="エラー", button="MB_OK", icon="MB_ICONERROR"):
    flg = win32api.MessageBox(
        application.app.hwnd,
        alert,
        title,
        win32con.__dict__[button] | win32con.__dict__[icon],
    )
    return flg

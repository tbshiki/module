from itsdangerous import want_bytes
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
    flg = win32api.MessageBox(application.app.hwnd, alert, title, win32con.__dict__[button] | win32con.__dict__[icon],)
    return flg


def FreezePanes(ws=None, wb=None, row=1, col=0):
    """ウインドウ枠の固定

    Args:
        ws (_type_, optional): _description_. Defaults to xw.books.active.app.api.ActiveWindow.
        wb (_type_, optional): _description_. Defaults to xw.books.active.
        row (int, optional): _description_. Defaults to 1.
        col (int, optional): _description_. Defaults to 0.
    """

    if ws == None:
        try:
            ws = xw.books.active.sheets.active
        except:
            return False

    ws.activate()
    wb = xw.books.active
    aw = wb.app.api.ActiveWindow
    aw.FreezePanes = False
    aw.SplitColumn = col
    aw.SplitRow = row
    aw.FreezePanes = True


def FreezePanes0(ws=None, wb=None, row=0, col=0):
    """ウインドウ枠固定の解除

    Args:
        ws (_type_, optional): _description_. Defaults to xw.books.active.app.api.ActiveWindow.
        wb (_type_, optional): _description_. Defaults to xw.books.active.
        row (int, optional): _description_. Defaults to 1.
        col (int, optional): _description_. Defaults to 0.
    """
    if ws == None:
        try:
            ws = xw.books.active.sheets.active
        except:
            return False

    ws.activate()
    wb = xw.books.active
    aw = wb.app.api.ActiveWindow
    aw.FreezePanes = False
    aw.SplitColumn = 0
    aw.SplitRow = 0


def check_sheet_add(sheet_name, wb=None, position=0):
    """同名シートが存在するかチェックしてシート追加

    Args:
        sheet_name (str): シート名
        wb (xw.Book, optional): xw.Book. Defaults to None.
        position (int, optional): 位置. Defaults to 0.

    Returns:
        bool: True:存在する, False:存在しない
        sh: シート
    """

    if wb == None:  # Excelが起動していない場合はFalseを返す
        try:
            wb = xw.books.active
        except:
            return False

    try:
        sh = wb.sheets.add(sheet_name, before=wb.sheets[position])
    except:
        sh = wb.sheets[sheet_name]
        all_sh_name = [sh.name for sh in wb.sheets]
        counter = 2

        while True:
            if f"{sheet_name} ({counter})" in all_sh_name:
                counter += 1
                if counter > 50:
                    return False  # 50も作成してたらおかしいのでその場合はFalseを返す
            else:
                break

        sh.name = f"{sheet_name} ({counter})"
        sh = wb.sheets.add(sheet_name, before=wb.sheets[position])

    return sh


def check_sheet_copy(sheet_source_name, wb_source=None, wb_destination=None, position=0):
    """同名シートが存在するかチェックしてシートコピー

    Args:
        sheet_name (str): シート名
        wb (xw.Book, optional): xw.Book. Defaults to None.
        position (int, optional): 位置. Defaults to 0.

    Returns:
        bool: True:存在する, False:存在しない
        sh: シート
    """

    if wb_destination == None:  # Excelが起動していない場合はFalseを返す
        try:
            wb_destination = xw.books.active
        except:
            return False

    all_sh_name = [sh.name for sh in wb_destination.sheets]

    if sheet_source_name in all_sh_name:
        # 同名シートが存在するので(*)を付ける
        counter = 2

        while True:
            if f"{sheet_source_name} ({counter})" in all_sh_name:
                counter += 1
                if counter > 50:
                    return False  # 50も作成してたらおかしいのでその場合はFalseを返す
            else:
                break
        sheet = wb_destination.sheets[sheet_source_name]
        sheet.name = f"{sheet_source_name} ({counter})"

    add_sh = wb_source.sheets[sheet_source_name].copy(before=wb_destination.sheets[position])

    return add_sh

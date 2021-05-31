import xlwings as xw
import sys

# モジュールのあるパスを追加
sys.path.append('../module')

#数字とアルファベットを変換
import convert_alphabet_to_num
import convert_num_to_alphabet

def last(sheet):
    last_col_A = sheet.range(1, sheet.cells.last_cell.column).end('left').column

    last_row = 1
    for i in range(last_col_A + 1)[1:]:
        max_row = sheet.range(sheet.cells.last_cell.row, i).end('up').row
        if max_row > last_row:
            last_row = max_row

    return last_col_A,last_row


def col_dict(sheet, colstart : int = 0, colend : int = 0):
    last_col, last_row= last(sheet)

    if colend == 0:
        colend = last_col

    list_all: dict = []
    for col in range(int(last_col))[colstart:colend]:
        strcol = convert_num_to_alphabet.num2A(col + 1)
        list_all.append(sheet.range(strcol +'1:'+ strcol + str(last_row)).value)

    return last_col, last_row, list_all
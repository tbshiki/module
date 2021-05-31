
import xlwings as xw
import sys

# モジュールのあるパスを追加
sys.path.append('../module')
#数字とアルファベットを変換
import convert_alphabet_to_num
import convert_num_to_alphabet

def all(sheet):
    max_col = sheet.range(1, sheet.cells.last_cell.column).end('left').column
    last_row = 1
    for i in range(max_col + 1)[1:]:
        max_row = sheet.range(sheet.cells.last_cell.row, i).end('up').row
        if max_row > last_row:
            last_row = max_row
    return last_row

def col(sheet, col):
    last_row = sheet.range(sheet.cells.last_cell.row, convert_alphabet_to_num.A2num(col)).end('up').row
    return last_row
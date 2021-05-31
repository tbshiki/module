##
# gspread_update_cells(ワークシート変数, リスト変数, 開始セル)
#
##
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import pandas as pd
import gspread
import re
import sys
from datetime import datetime, timedelta

# モジュールのあるパスを追加
sys.path.append('../module')
#数字とアルファベットを変換
import convert_alphabet_to_num
import convert_num_to_alphabet

#https://tanuhack.com/gspread-dataframe/
#指定セルから連想配列を貼り付け
#gspread_me.free(worksheet, list, startcell)
def free(worksheet, list, startcell):
    df = pd.DataFrame(list)
    col_lastnum = len(df.columns) # DataFrameの列数
    row_lastnum = len(df.index)   # DataFrameの行数

    start_cell = startcell # 列はA〜Z列限定
    start_cell_col = re.sub(r'[\d]', '', start_cell)
    start_cell_row = int(re.sub(r'[\D]', '', start_cell))

    # 展開を開始するセルからA1セルの差分
    row_diff = start_cell_row - 1
    col_diff = convert_alphabet_to_num.A2num(start_cell_col) - convert_alphabet_to_num.A2num('A')

    # DataFrameのヘッダーと中身をスプレッドシートの任意のセルから展開する
    cell_list = worksheet.range(start_cell + ':' + convert_num_to_alphabet.num2A(col_lastnum + col_diff) + str(row_lastnum + row_diff))
    for cell in cell_list:
        val = df.iloc[cell.row - row_diff - 1][cell.col - col_diff - 1]
        cell.value = val
    worksheet.update_cells(cell_list)

#指定セル範囲へ配列を貼り付け
#gspread_me.just(worksheet, list, startcell, lastcell)
def just(worksheet, list, startcell, lastcell):
    cell_list = worksheet.range(startcell + ":" + lastcell)

    for cell, item in zip(cell_list,list):
            cell.value = item
    worksheet.update_cells(cell_list)

#ワークシートとワークブックを指定して取得 sheetの引数いれなければ一番左のシートが返る
#workbook, worksheet = gspread_me.get(path, SPREADSHEET_KEY, sheet)
def get(path, SPREADSHEET_KEY, sheet = 1):

    try:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
        gc = gspread.authorize(credentials)
        workbook = gc.open_by_key(SPREADSHEET_KEY)
    except:
        print('スプレッドシートが見つかりません')
        sys.exit()

    if type(sheet) is int:
        worksheet = workbook.get_worksheet(sheet - 1) #ワークシートのインデックスは0から始まる
    else:
        worksheet = workbook.worksheet(sheet)

    return workbook,worksheet

#ワークシートとワークブックを指定して取得 sheetの引数いれなければ一番左のシートが返る
#workbook = gspread_me.get_book(path, SPREADSHEET_KEY)
def get_book(path, SPREADSHEET_KEY):

    try:
        scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scope)
        gc = gspread.authorize(credentials)
        workbook = gc.open_by_key(SPREADSHEET_KEY)
    except:
        print('スプレッドシートが見つかりません')
        sys.exit()

    return workbook

#ワークシートとワークブックを指定して取得 sheetの引数いれなければ一番左のシートが返る
#list_all, last_row, last_col = gspread_me.get_all(worksheet)
def get_all(worksheet):
    list_all = worksheet.get_all_values()
    last_col = max([len(value) for value in list_all])
    last_row = len(list_all)

    return list_all, last_col, last_row
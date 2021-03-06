import os
import tkinter as tk
from tkinter import filedialog

def main(file_type = 'all', selftitle = '', dir = ''):

    ##
    # 引数 file_type
    # all or OTHER : '*'
    # xls          : '*.xls;*.xlsx;*.xlsm'
    # txt          : '*.txt;'
    # csv          : '*.csv'
    ##

    if file_type == 'xls':
        file_types = [('Excelファイル','*.xls;*.xlsx;*.xlsm')]
    elif file_type == 'txt':
        file_types = [('txtファイル','*.txt')]
    elif file_type == 'csv':
        file_types = [('csvファイル','*.csv')]
    else:
        file_types = [('','*')]

    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    root.lift()
    root.focus_force()

    if dir == '':
        dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__))) # 一つ上のディレクトリを参照

    if selftitle == '':
        dir_path = tk.filedialog.askopenfilename(filetypes = file_types, initialdir = dir)
    else:
        dir_path = tk.filedialog.askopenfilename(filetypes = file_types, initialdir = dir, title = selftitle)
    return dir_path

if __name__ == '__main__':
    main()
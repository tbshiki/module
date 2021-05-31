#リストを逆順に
def reverse_list(list):
    for i in range(len(list) // 2):
        list[i],list[-1-i] = list[-1-i],list[i]

#参考 https://qiita.com/take333/items/b61e43c68751260689a6

#リストをn個にわける
def split_list(list, n):
    """
    リストをサブリストに分割する
    :param l: リスト
    :param n: サブリストの要素数
    :return:
    """
    for idx in range(0, len(list), n):
        yield list[idx:idx + n]

    #参考 https://www.python.ambitious-engineer.com/archives/1843
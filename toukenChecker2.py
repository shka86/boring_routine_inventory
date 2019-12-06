#!/usr/bin/python3
# -*- coding: utf-8 -*-

import datetime
import openpyxl
from copy import copy

# windowsの場合はコレ↓しとかないとエラー出る
import locale
locale.setlocale(locale.LC_ALL, '')

class Katana(object):
    def __init__(self, row):
        self.name = row[0].value
        self.yomi = row[1].value
        self.owner = row[2].value
        self.category = row[3].value
        self.note = row[4].value
        self.inventory_this_month = row[5].value
        self.inventory_last_month = row[6].value

        # super(Katana, self).__init__()
        # self.arg = arg


def main():
    # ファイルを開く
    input_file = 'toukenDaicho.xlsx'
    wb = openpyxl.load_workbook(input_file)
    sheet = wb["Sheet1"]

    # 今月の棚卸状況が"済"でない刀とその所有者をリストアップする
    yet_katanas = []
    yet_bushos = []

    for i in range(1,sheet.max_row):
        katana = Katana(list(sheet.rows)[i])

        if not(katana.inventory_this_month == "済"):
            yet_katanas.append(katana.name)
            yet_bushos.append(katana.owner)

    # まだ棚卸されていない刀とその所有者を表示する
    print("【まだ棚卸されていない刀】")
    for i in range(len(yet_katanas)):
        print(yet_katanas[i] + " -- " + yet_bushos[i])

    # リスト書き出し
    with open("out.csv", "w") as f:
        for yet_katana in yet_katanas:
            f.write(yet_katana + "\n")


if __name__ == '__main__':
    main()

#!/usr/bin/python3
# -*- coding: utf-8 -*-

import datetime
import openpyxl
from copy import copy

# windowsの場合はコレ↓しとかないとエラー出る
import locale
locale.setlocale(locale.LC_ALL, '')

def main():
    # ファイルを開く
    input_file = 'toukenDaicho.xlsx'
    wb = openpyxl.load_workbook(input_file)
    sheet = wb["Sheet1"]

    # 今月の棚卸状況が"済"でない刀とその所有者をリストアップする
    yet_katanas = []
    yet_bushos = []
    for i in range(2,sheet.max_row):
        cell = list(sheet.columns)[5][i]
        if not cell.value == "済":
            yet_bushos.append(cell.offset(0,-3).value)
            yet_katanas.append(cell.offset(0,-5).value)

    # まだ棚卸されていない刀とその所有者を表示する
    print("【まだ棚卸されていない刀】")
    for i in range(len(yet_katanas)):
        print(yet_katanas[i] + " -- " + yet_bushos[i])

if __name__ == '__main__':
    main()
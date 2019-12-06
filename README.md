boring_routine_inventory
====

## 概要
エクセルによる毎月の棚卸作業をpythonで自動化します。

本リポジトリはこれまでプログラミングに縁がなかった方に向けて、
「こんなことができますよ。」というデモンストレーション的な意味合いで作成したものです。
すぐに試せるように、サンプルファイルとして戦国武将たちが管理している刀剣の台帳を同梱しています。というかそれをチェックすることしかできませんが、イメージを掴んでください。


## 必要環境
- Excel  
サンプルの台帳にエクセルを使用しています。csvでも構いません。

- python3

## インストール方法
- 本リポジトリ
    インストールというかダウンロードしていただければよいです。
    `Clone　or　Download`　から　`Download　ZIP`　でダウンロードしてください。
    お好みの場所で展開すればOKです。

- python
    インストールされていない場合は入れてください。
    なにそれ？と思った場合はおそらく入ってないです。
    linux使ってる人には説明は不要ですね。

    インストールはこちら↓  
    [https://www.python.org/downloads/windows/](https://www.python.org/downloads/windows/)

    - windows(64bit)の場合  
    "Download Windows x86-64 executable installer"

    - windows(32bit)の場合  
    "Download Windows x86 executable installer"

    をダウンロードして、インストールしてくださいね。

- openpyxl のインストール
    pythonの外部ライブラリを使用していますので、コマンドウィンドウに下記のコマンドを打ち込んでインストールしてください。
    `pip install openpyxl`

## 使い方
windowsの場合は、ダウンロードしたフォルダで右クリックして、
コマンドプロンプトだかPowerShell的なものを開いてください。
あの黒い画面が出てきますが我慢してください。

こちらもlinux使ってる人には説明は不要ですね。

次のコマンドを入力し、Enterで実行してください。
```bash
python toukenChecker1.py
```
```bash
python toukenChecker2.py
```

実行結果サンプル
```
【まだ棚卸されていない刀】
御手杵 -- 結城晴明
童子切安綱 -- 豊臣秀吉
大典太光世 -- 豊臣秀吉
数珠丸 -- 日蓮
```

サンプルの台帳`toukenDaicho.xlsx`を見ていただければわかりますが、
2019年3月現在、棚卸完了じるしである「済」が入っていない刀剣とその所有者がリストアップされます。


## License

MIT License

## Author

[shka86](https://github.com/shka86)

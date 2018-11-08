クエリ太郎 (Excel add-in)
====

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](./LICENSE)

DML文のテンプレート等を作成するExcelアドイン
DML: Data Manipulation Language

# 1. Description
Excelで以下の作業が簡単になります。
- DMLでのレコードのExcelシートの作成
- DML文の作成
- 記述したレコードの確認

本ソフトをインストールすることで、  
Excelの「データ」のリボンに「クエリ太郎」というグループが作成されます。  
そこから使いたい機能をクリックするだけで実行できます。

# 2. Demo
アニメGIFを貼付けて実際の動作例を見せます。

## 2.1. DMLのテンプレートの作成
![DMLテンプレートの作成.gif](.\movie\DMLテンプレートの作成.gif)

## 2.2. クエリの作成
![DML文の作成.gif](.\movie\DML文の作成.gif)

# 3. VS. 
面倒なレコードのDMLの作成が
本アドインが代わりに行ってくれます。

カラム名の上の行にPKと記述することで主キーを設定することも可能です。
レコードの値にNULLを記述することで、NULLにも対応することができます。

# 4. Usage
使い方

## 4.1. テンプレートの作成
1. 「データ」のリボンにある「テンプレートの作成」をクリック
2. テーブル名、カラム数、レコード数を入力
3. 「DMLテンプレートの作成」をクリック

## 4.2. レコードの作成
1. 主キーを記入 (カラム名の上の行にPKと記入)
2. 各レコードの値を記述
3. 備考があれば備考を記述

## 4.3. DMLの作成
1. 「データ」のリボンにあるDMLの作成をクリック


# 5. Requirement
以下の環境で動作確認をしております。  
* Excel 2010

※ Windows版Excelのみ動作確認済み

# 6. Install

1. 以下の場所から「Source code (zip)」をダウンロード  
https://github.com/ta-kon/Numbering-ExcelAddin/releases

2. ダウンロードしたzipファイルを展開し、
展開したフォルダ内にある  
Install.vbs を実行

Excelのリボンに「連番太郎」が追加されます。

# 7. Uninstall
アンインストール方法

Uninstall.vbs を実行


# 8. Thanks

以下のサイトを参考にしました。  

[ある SE のつぶやき - VBScript で Excel にアドインを自動でインストール/アンインストールする方法](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)  
Install.vbs / Uninstall.vbs に使用しております。  

# 9. Licence
本プログラムはフリーソフトウェアです。  
個人・法人に限らず利用者は自由に使用および配布することができます。  

本プログラムは無償で利用できますが、  
作者は本プログラムの使用にあたり生じる障害や問題に対して一切の責任を負いません。  

ソースを利用する場合にはMITライセンスです。

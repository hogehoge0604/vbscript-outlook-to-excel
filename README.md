# VBscript Outlook to Excel

## インストール
outlook-to-excel.vbsをダウンロードする  
https://github.com/hogehoge0604/vbscript-outlook-to-excel/archive/master.zip

## 使い方
1. Excelファイルを新規作成する
2. outlook-to-excel.vbsを開き抽出条件、出力先のExcelに応じて設定ファイルを修正する
3. outlook-to-excel.vbsをダブルクリックする

## 設定
以下を設定で変更可能

```
' 受信トレイの下層フォルダ。受信トレイを抽出対象としたい場合は空にする
config.Add "outlook_sub_folder", "<フォルダ名>"
' 追記するExcelファイルまでのパス
config.Add "excel_file_path", "<Excelファイル名>"
' 追記するExcelファイルのワークシート名
config.Add "excel_worksheet_name", "<ワークシート名>"
' 抽出対象の件名。設定した文字列が含まれるメールを対象とする
config.Add "pickup_title", "<抽出条件の文字列>"
' 抽出対象件数
config.Add "max_pickup_num", "<チェック対象のメール件数>"
' 本文抽出条件
config.Add "body_grep_rule", "<本文抽出条件の正規表現>"
' Excelに追記する最初の列値(A~Zまで)
config.Add "excel_start_row", "<列値>"
' Excelに追記する最初の行。値が0の場合は最初にヒットした空セルの行を起点とする
config.Add "excel_start_line", <行値>
```

## ライセンス
MIT

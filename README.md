# 給与明細自動化ツール

## このツールは何か
会社で使用する給与明細を自動で作成するツールです。

## できること
- GUIウィンドウから名前と月を入力するだけで給与明細を自動生成
- Googleフォームの回答データを取得
- 仕訳や手当の自動判別
- Excelファイルとして出力

## 使用技術
- Python
- tkinter / customtkinter（GUI）
- openpyxl（Excel操作）
- gspread（GoogleスプレッドシートAPI）
- geopy（住所・位置情報の判定）
- google-auth（Google認証）
# プログラム1｜ライブラリ設定
import os
import xlwings as xw

# プログラム2｜mainプログラム
def main():

    # プログラム3｜Sheet1の取得
    wb = xw.Book.caller()
    sheetname = 'Sheet1'
    ws = wb.sheets(sheetname)

    # プログラム4｜フォルダパスの取得
    folderpath = ws.range('B1').value

    # プログラム5｜作成したいフォルダ名を取得
    cmax = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row
    for i in range(5, cmax+1):
        filename = ws.range('B' + str(i)).value

        # プログラム6｜フォルダを作成
        newfolderpath = os.path.join(folderpath, filename)
        if os.path.exists(newfolderpath) == False:
            os.makedirs(newfolderpath)
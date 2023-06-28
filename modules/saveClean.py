from tkinter inmport filedialog
import os
import glob
import openpyxl
import xlwings

#クリックイベント
def btn_click():
    print('saveClean処理開始')

    #フォルダー選択ダイアログを開く
    dir = filedialog.askdirectory(
        title = '実行したフォルダーパス選択'
    )

    print(dir)

    #選択したディレクトリに移動
    os.chdir(dir)
    #フォルダーないのすべてのxlsxを読み込む
    xlsx_files = glob.glob(dir + '/**/*.xlsx' , recursive=True)

    #設定情報定数
    cell_no = 'A1'
    zoom_scale = 100

    #エクセルファイルのループ実行
    for xlsx_file in xlsx_files:
        print(xlsx_file)
        #モジュールを利用してブックを取得
        wb = openpyxl.load_workbook(xlsx_file)

        #エクセルの全シートループ
        for ws in wb.worksheets:
            #見栄え調整インスタンス生成
            sv = ws.sheet_view
            sv.tabSelected = False
            #セルを選択
            sv.selection[0].activeCell = cell_no
            sv.selection[0].sqref = cell_no
            sv.selection[0].activeCellId = None
            #スクロール位置を指定
            sv.topLeftCell = 'A1'
            #表示倍率を100％を設定
            sv.zoomScale = zoom_scale
            sv.zoomScaleNormal = zoom_scale

        #最初のシートをアクティブ
        wb.active = wb.worksheets[0]
        wb.save(xlsx_file)


    #上記モジュールではシートタブが先頭にいかないので別モジュールで整形
    #エクセルファイルのループを実行
    for xlsx_file in xlsx_files:
        #Excelが画面に表示されないようにする
        xlwings.App(visible=False)
        print(xlsx_file)
        #モジュールを利用してブックを取得
        wbxlwings = xlwings.Book(xlsx_file)
        #シートの先頭を選択
        wbxlwings.sheets[0].select()
        wbxlwings.save(xlsx_file)
        wbxlwings.close()

    print('saveClean完了')
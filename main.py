form modules import saveClean
from modules import sheetRename
import tkinter


#Tkクラス作成
frm = tkinter.tk()
#画面サイズ
frm.geometry('600x400')
#画面タイトル
frm.title('エクセル操作ツール')

#ボタン作成
btn = tkinter.Button(frm , text='エクセル綺麗に保存' , width=14 , command=saveClean.btn_click)
#配置
btn.pack(fill='x' , padx=20 , side='top')

#ボタン作成
btn = tkinter.Button(frm , text='シート芽衣A1反映' , width=14 , command=saveClean.btn_click)
#配置
btn.pack(fill='x' , padx=20 , side='top')

#画面をそのまま表示
frm.mainloop()
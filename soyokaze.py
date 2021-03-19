#メインモジュール：画面用

import os
import sys
import tkinter
import docx
import docx_test
import sqlite3
from pdfminer.high_level import extract_text


from tkinter import messagebox
from tkinter import filedialog



#テストモード切替
testMode = 0

#clickイベントの戻りイベント保持→画面適用
def testresult(txt):
    ret = docx_test.before_check(txt)
    label3['text'] = ret

# clickイベント
def button1_clicked(event):
    root.after(1,select_docfile)
def button2_clicked(event):
    root.after(1,select_docfile)
def button3_clicked(event):
    root.after(1,testresult(txt))

#動作確認用メソッド
def instest_clicked(event):
    root.after(1,docx_test.ins_test())
def seltest_clicked(event):
    root.after(1,docx_test.sel_test())


#ファイル選択処理→テキストボックスへ
def select_docfile():
    fTyp = [("","*.docx")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    filepath = filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
    txt.delete(0, tkinter.END)
    txt.insert(tkinter.END,filepath)

root = tkinter.Tk()
root.title(u"そよかぜ比較ツール")
root.geometry("600x400")

xpos = 50
ypos = 50

#ラベル
Static1 = tkinter.Label(text=u'ボタンは上から操作して下さいね',fg = "#ff0000")
Static1.pack()

label1 = tkinter.Label(text='wordファイルを読み込むボタンです')
label1.pack()

# ボタン
button1 = tkinter.Button(text='読み込み', width=40)
button1.bind('<Button-1>', button1_clicked)
button1.pack()

label2 = tkinter.Label(text='pdfファイルを読み込むボタンです')
label2.pack()

button2 = tkinter.Button(text='pdfと比較', width=40)
button2.bind('<Button-1>', button2_clicked)
button2.pack()

Static2 = tkinter.Label(text=u'ファイルパス',fg = "#ff0000")
Static2.pack()

txt = tkinter.Entry(width=50)
txt.pack()

button3 = tkinter.Button(text='読み込みエラーチェック', width=40)
button3.bind('<Button-1>', button3_clicked)
button3.pack()

#読み込みました ファイルの内容
label3 = tkinter.Label(text=u'')
label3.pack()

#動作確認用テストボタン
if testMode == 1:
    seltest = tkinter.Button(text='sel', width=40)
    seltest.bind('<Button-1>', seltest_clicked)
    seltest.pack()

    instest = tkinter.Button(text='ins', width=40)
    instest.bind('<Button-1>', instest_clicked)
    instest.pack()



root.mainloop() 


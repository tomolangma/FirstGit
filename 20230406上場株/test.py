import pandas as pd
import glob
import PyPDF2

from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl
from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.units import mm

print(pd.__version__)
# 1.2.2

# 読み込み
df = pd.read_excel('monthexc/202302.xlsx')
x = df.iloc[:,1]
y = df.iloc[:,2]
z = df.iloc[:,11]
# 必要な部分だけ抽出結合
df1 = pd.concat([x,y,z], axis=1)
# https://note.nkmk.me/python-pandas-concat/
df1.columns=["コード", "名前", "終値平均"]
# df2 = df1.dropna(how='any')
# https://note.nkmk.me/python-pandas-nan-dropna/
# print(df2)

# 検索
df3 =df1[df1['コード'] == 3181]
print(df1[df1['コード'] == 3181])
# https://note.nkmk.me/python-pandas-str-contains-match/

print(df1.reset_index().query('コード == 3181').index[0])
print(list(df1.reset_index().query('コード == 3181').index))
# https://note.nkmk.me/python-pandas-get-loc-row-column-num/

x = df1.reset_index().query('コード == 9073').index[0]
y = x // 43 + 1
print(y)
z = x % 43 - 4
print(z)

reader1 = PyPDF2.PdfReader('monthpdf\st_202302-2.pdf')

writer = PyPDF2.PdfWriter()

writer.add_page(reader1.pages[int(y-1)])

with open('完成品\st_202302-2xx.pdf', 'wb') as f:
    writer.write(f)


in_path = '完成品\st_202302-2xx.pdf'
out_path = '完成品\st_202302-2xxok.pdf'

# 保存先PDFデータを作成
cc = canvas.Canvas(out_path, pagesize=landscape(A4))

# PDFを読み込む
pdf = PdfReader(in_path, decompress=False)

# PDFのページデータを取得
page = pdf.pages[0]

# PDFデータへのページデータの展開
pp = pagexobj(page) #ページデータをXobjへの変換
rl_obj = makerl(cc, pp) # ReportLabオブジェクトへの変換  
cc.doForm(rl_obj) # 展開

# 長方形の描画
cc.setFillColor("yellow", 0.3)
# cc.setStrokeColorRGB(1.0, 0, 0)


cc.rect(32, 423-10*z, 777, 10, fill=1, stroke = 0)

# ページデータの確定
cc.showPage()

# PDFの保存
cc.save()
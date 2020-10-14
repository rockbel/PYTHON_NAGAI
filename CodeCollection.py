# エラー対処
#===================================================================================================================================
# SyntaxError: (unicode error) 'unicodeescape' codec can't decode bytes in position 2-3: truncated \UXXXXXXXX escape
# ※Windowsでパスを指定する時に気を付けること
# エラーの原因としては、パスなどの文字列に“\”が使われることによりその文字列がエスケープシーケンスとしてみなされているからです。
# https://office54.net/python/python-unicode-error


#===================================================================================================================================
#ファイルの存在を確認する
import os
if(os.path.exists(r"C:\Users\utkn009\Desktop\HBSCA104IM03.pdf")):
    print("ある")
else:
    print("ファイルない")

#===================================================================================================================================
#PDFをcsvに変換
#tabula documents
#https://tabula-py.readthedocs.io/en/latest/tabula.html
#同名のファイルは上書きする

import pandas as pd
import tabula
tabula.convert_into(r"\\126.0.0.42\代_共通\KAMBARA\ARRIVAL NOTICE\KCRA0104E\HBSCA104IM03.pdf", r"C:\Users\utkn009\Desktop\HBSCA104IM03.csv", stream=True, output_format="csv", pages="all")

#===================================================================================================================================
#ファイル名一覧
import os

path = r'\\126.0.0.42\代_共通\KAMBARA\ARRIVAL NOTICE\KCRA0104E'
files = os.listdir(path)
print(files)
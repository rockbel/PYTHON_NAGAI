#tabula documents
#https://tabula-py.readthedocs.io/en/latest/tabula.html
#同名のファイルは上書きする

import tabula
import glob
import sys
import pandas as pd

#動作確認済み
# args = sys.argv
# files = glob.glob(args[1] + r'/*.pdf')
# for file in files:
#     csvname = file[0:len(file) - 4] + ".csv"
#     tabula.convert_into(file, csvname, stream=True, output_format="csv", pages="all")
# print("完了")

#Ver2
args = sys.argv
files = glob.glob(args[1] + r'/*.pdf')
for file in files:
    csvname = file[0:len(file) - 4] + ".csv"
    df = tabula.read_pdf(file, stream=True, pages="all", encoding="Shift-JIS") #格子無　stream=True / 格子有　lattice=True
    d = pd.concat(df, ignore_index=True)
    d.to_csv(csvname, encoding='Shift-JIS',index=False )
print(args[1] + " pff →　csv　変換完了")
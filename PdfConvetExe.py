#tabula documents
#https://tabula-py.readthedocs.io/en/latest/tabula.html
#同名のファイルは上書きする

import tabula
import glob
import sys

args = sys.argv
files = glob.glob(args[1] + r'/*.pdf')
for file in files:
    csvname = file[0:len(file) - 4] + ".csv"
    tabula.convert_into(file, csvname, stream=True, output_format="csv", pages="all")
print("完了")
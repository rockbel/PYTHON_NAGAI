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

#===================================================================================================================================
#ブラウザを開いて表示

import webbrowser
webbrowser.open("http://fs2/scripts/cbag/ag.exe?page=MailSend&fid=&cp=el&flagged=")

#===================================================================================================================================
#サイボウズでメール送信サンプル

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import time

#WebDrivwr起動
driver = webdriver.Chrome(r"\\126.0.0.42\個人(フル)\田中智也\1.田中\chromedriver_win32\chromedriver.exe")
#最大化
driver.maximize_window()
#サイボウズの社外メール作成のURL入力
driver.get("http://fs2/scripts/cbag/ag.exe?page=MailSend&fid=&cp=el&flagged=")
#_IDが表示されるまでwait
WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, '_ID')))
#「切り替える」のxpathをクリック
driver.find_element_by_xpath('/html/body/div[2]/div[1]/div/table/tbody/tr/td/center/div/table/tbody/tr[2]/td[2]/div/div/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/a').click()
#セレクト要素の処理
select_element = driver.find_element_by_name("Group")
select_object = Select(select_element)
select_object.select_by_value("170")
#「切り替える」をnameでクリック
driver.find_element_by_name("Submit").click()
#セレクト要素の処理
select_element = driver.find_element_by_name("_ID")
select_object = Select(select_element)
select_object.select_by_value("1157")
#パスワード入力
driver.find_element_by_name("Password").send_keys("0117")
#「ログイン」をクリック
driver.find_element_by_name("Submit").click()
#「送信」ボタンを押せるようになるまでwait
WebDriverWait(driver, 10).until(EC.element_to_be_clickable( (By.NAME,"Submit")) )
#アドレスを宛先に入力　サイボウズの仕様上","で区切った文字を一気に挿入は不可
ToAds = ['"gum@kambarakisen.com" <gum@kambarakisen.com>', '"Gaojianping" <gaojp@kambarakisen.com>', '"HD group" <MLKK.TEIKI@tsuneishi.com>', '"Chenchenggang" <chencg@kambarakisen.com>', '"CONTAINER" <container@kambarakisen.com>', '"Zhuli" <zhul@kambarakisen.com>']
for ToAd in ToAds:
    driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[3]/td/div/ul/li/textarea").send_keys(ToAd)
    driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[3]/td/div/ul/li/textarea").send_keys(",")
#アドレスをCCに入力　サイボウズの仕様上","で区切った文字を一気に挿入は不可
CCAds = ['kawahara@nagaij.co.jp']
for CCAd in CCAds:
    driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[4]/td/div/ul/li/textarea").send_keys(CCAd)
    driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[4]/td/div/ul/li/textarea").send_keys(",")
#表題入力
driver.find_element_by_name("Subject").send_keys("INVENTRY REPORT IMARI 16th")
#"\n"で区切りながら本文入力
inputtext = "Dear Chen san\nDear Sandra san\n\nGood day,\nPls find attached file Inventry report IMARI\nPls advise me, Load quantity of KORA-0111W.\nIt may increase or decrease in number from these prospects.\nI will report these numbers after confirmation.\n\nTks & B.rgds\nIMARI AGENT/JAPAN\nTanaka"
for part in inputtext.split('\n'):
       driver.find_element_by_id("Data").send_keys(part)
       driver.find_element_by_id("Data").send_keys("\n")
#セレクト要素の処理　署名を(なし)に
select_element = driver.find_element_by_name("SigID")
select_object = Select(select_element)
select_object.select_by_visible_text('(なし)')
#ファイルを添付　※ファイル名をsend_keysで送る
AttachedFiles = [r"\\126.0.0.42\業務_共通\神原荷役作業\Booking_List_IMARI 2020.xls",r"\\126.0.0.42\業務_共通\神原荷役作業\IMARI INVENTRY REPORT.xlsx"]
for File in AttachedFiles:
    driver.find_element_by_id("files1").send_keys(File)

#参考資料
#
#.find_element_by_xpathと.find_element(s)_by_xpath
#https://www.366service.com/jp/qa/ab49bd049dbd66089443a84c8aee20d5

#[EC.visibility_of_element_located((By.で指定できる要素一覧]
# ID
# XPATH
# LINK_TEXT
# PARTIAL_LINK_TEXT
# NAME
# TAG_NAME
# CLASS_NAME
# CSS_SELECTOR

#選択要素の操作
#https://www.selenium.dev/documentation/ja/support_packages/working_with_select_elements/

#===================================================================================================================================
#
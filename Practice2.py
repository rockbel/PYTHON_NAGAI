import os,sys
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import subprocess

# フォルダ指定の関数
def dirdialog_clicked():
    iDir = os.path.abspath(os.path.dirname(__file__))
    iDirPath = filedialog.askdirectory(initialdir = r"c:/")
    #entry1.set(iDirPath)
    entry1.set('"' + iDirPath + '"')

# 実行ボタン押下時の実行関数
def conductMain():
    from selenium import webdriver
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.select import Select
    import time

    driver = webdriver.Chrome(r"\\126.0.0.42\個人(フル)\田中智也\1.田中\chromedriver_win32\chromedriver.exe")
    driver.maximize_window()
    driver.get("http://fs2/scripts/cbag/ag.exe?page=MailSend&fid=&cp=el&flagged=")
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, '_ID')))
    driver.find_element_by_xpath('/html/body/div[2]/div[1]/div/table/tbody/tr/td/center/div/table/tbody/tr[2]/td[2]/div/div/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/a').click()
    select_element = driver.find_element_by_name("Group")
    select_object = Select(select_element)
    select_object.select_by_value("170")
    driver.find_element_by_name("Submit").click()
    select_element = driver.find_element_by_name("_ID")
    select_object = Select(select_element)
    select_object.select_by_value("1157")
    driver.find_element_by_name("Password").send_keys("0117")
    driver.find_element_by_name("Submit").click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable( (By.NAME,"Submit")) )
    #buf[1]:宛先, buf[2]:CC, buf[3]:BCC
    ToAds = ['"gum@kambarakisen.com" <gum@kambarakisen.com>', '"Gaojianping" <gaojp@kambarakisen.com>', '"HD group" <MLKK.TEIKI@tsuneishi.com>', '"Chenchenggang" <chencg@kambarakisen.com>', '"CONTAINER" <container@kambarakisen.com>', '"Zhuli" <zhul@kambarakisen.com>']
    for ToAd in ToAds:
        driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[3]/td/div/ul/li/textarea").send_keys(ToAd)
        driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[3]/td/div/ul/li/textarea").send_keys(",")
    CCAds = ['kawahara@nagaij.co.jp']
    for CCAd in CCAds:
        driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[4]/td/div/ul/li/textarea").send_keys(CCAd)
        driver.find_element_by_xpath("/html/body/div[2]/div[4]/div/form/table/tbody/tr/td/table/tbody/tr[2]/td/div/div[1]/div/table/tbody/tr[4]/td/div/ul/li/textarea").send_keys(",")
    driver.find_element_by_name("Subject").send_keys("INVENTRY REPORT IMARI 16th")
    inputtext = "Dear Chen san\nDear Sandra san\n\nGood day,\nPls find attached file Inventry report IMARI\nPls advise me, Load quantity of KORA-0111W.\nIt may increase or decrease in number from these prospects.\nI will report these numbers after confirmation.\n\nTks & B.rgds\nIMARI AGENT/JAPAN\nTanaka"
    for part in inputtext.split('\n'):
        driver.find_element_by_id("Data").send_keys(part)
        driver.find_element_by_id("Data").send_keys("\n")
    select_element = driver.find_element_by_name("SigID")
    select_object = Select(select_element)
    select_object.select_by_visible_text('(なし)')
    AttachedFiles = [r"\\126.0.0.42\業務_共通\神原荷役作業\Booking_List_IMARI 2020.xls",r"\\126.0.0.42\業務_共通\神原荷役作業\IMARI INVENTRY REPORT.xlsx"]
    for File in AttachedFiles:
        driver.find_element_by_id("files1").send_keys(File)

if __name__ == "__main__":

    # rootの作成
    root = Tk()
    root.title("練習")

    # Frame1の作成
    frame1 = ttk.Frame(root, padding=10)
    # frame1.grid(row=0, column=1, sticky=E)

    # # 「フォルダ参照」ラベルの作成
    # IDirLabel = ttk.Label(frame1, text="フォルダ参照＞＞", padding=(5, 2))
    # IDirLabel.pack(side=LEFT)

    # # 「フォルダ参照」エントリーの作成
    # entry1 = StringVar()
    # IDirEntry = ttk.Entry(frame1, textvariable=entry1, width=30)
    # IDirEntry.pack(side=LEFT)

    # # 「フォルダ参照」ボタンの作成
    # IDirButton = ttk.Button(frame1, text="参照", command=dirdialog_clicked)
    # IDirButton.pack(side=LEFT)

    # Frame2の作成
    frame2 = ttk.Frame(root, padding=10)
    # frame2.grid(row=2, column=1, sticky=E)

    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid(row=5,column=1,sticky=W)

    # 実行ボタンの設置
    button1 = ttk.Button(frame3, text="実行", command=conductMain)
    button1.pack(fill = "x", padx=30, side = "left")

    # # キャンセルボタンの設置
    # button2 = ttk.Button(frame3, text=("閉じる"), command=quit)
    # button2.pack(fill = "x", padx=30, side = "left")

    root.mainloop()
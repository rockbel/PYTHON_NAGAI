# import requests
# from bs4 import BeautifulSoup

# html_doc = requests.get("https://review-of-my-life.blogspot.com").text
# soup = BeautifulSoup(html_doc, 'html.parser') # BeautifulSoupの初期化
# tags = soup.find_all("h3", {"class": "post-title"})
# for tag in tags:
#  print (tag.a.string)
#  print (tag.a.get("href"))

import pandas as pd
columns = ["Name", "Url"]
df1 = pd.DataFrame(columns=columns) # 列名を指定する# TODO1 以下の表のようになるように、データフレームを作成してください。
se = pd.Series(['データ解析の実務プロセス入門（あんちべ）』を読んで特に学びが多かったこと', 'https://review-of-my-life.blogspot.com/2018/03/moji-okosi-1.html'], columns) # 行を作成
df1 = df1.append(se, columns) # データフレームに行を追加
se = pd.Series(['sqlite3覚書 データベースに接続したり、中身のテーブル確認したり', 'https://review-of-my-life.blogspot.com/2018/04/sqlite3.html'], columns) # 行を作成
df1 = df1.append(se, columns)
se = pd.Series(['LINEから送った画像を文字起こししてくれるアプリを作るときのメモ①', '	https://review-of-my-life.blogspot.com/2018/03/moji-okosi-1.html'], columns) # 行を作成
df1 = df1.append(se, columns)
df1.to_csv("test2.csv", encoding="shift_jis")
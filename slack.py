# -*- coding: utf-8 -*-
## cron設定 https://fremilli.com/mac-cron-python/
# gspreadの使い方 https://zak-papa.com/python_gspread_cell_input_output
# google api https://tanuhack.com/operate-spreadsheet/

# chrome のバージョンアップ
# https://chromedriver.chromium.org/downloads から適したバージョンを入れる
#
#
# ####################


# scraiping
import os
import requests

# spreadsheet
import gspread
import json
import datetime
from oauth2client.service_account import ServiceAccountCredentials
import chromedriver_binary

# scraiping
TOKEN = "xoxb-2519179986917-2528733205924-zaiqaCjkkciN0IiJ9g4bxVgn"
CHANNEL = "#bot"

from selenium import webdriver
from time import sleep
from bs4 import BeautifulSoup
from webdriver_manager.chrome import ChromeDriverManager

# spreadsheet
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    "/Users/morishitaseiju/Documents/app/slack/bot-python/search-tracker-328413-2f7d16436ff4.json",
    scope,
)
gc = gspread.authorize(credentials)
SPREADSHEET_KEY = "1ziRNv4-S8ErU6MIy_6HIYQdb7sVFjc_-Kq4ieUdWiNM"
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

today = datetime.date.today()
year = today.year
month = today.month
day = today.day
date = str(year) + "/" + str(month) + "/" + str(day)


wait_time = 10

search_word_list = [
    "電動昇降デスク コスパ",
    "テレワーク サボってしまう",
    "在宅勤務 サボり",
    "スマセル 売れない",
    "メルカリ airbods 偽物",
    "リモートワーク サボってしまう",
    "ルームシェア お金の管理",
    "スタンディングデスク コスパ",
    "fenge 電動昇降デスク レビュー",
    "ルームシェア 共同口座",
]


def split_list(l, n):
    """
    リストをサブリストに分割する
    :param l: リスト
    :param n: サブリストの要素数
    :return:
    """
    for idx in range(0, len(l), n):
        yield l[idx : idx + n]


# driver = webdriver.Chrome('/opt/homebrew/bin/chromedriver')
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.implicitly_wait(wait_time)
driver.get("https://seocheki.net/")
search_bar = driver.find_element_by_name("u")
search_bar.send_keys("https://dinheiro-m.com/")
search_bar.submit()

word1 = driver.find_element_by_id("word1")
word2 = driver.find_element_by_id("word2")
word3 = driver.find_element_by_id("word3")

div_result = list(split_list(search_word_list, 3))
eternal_text = ""
ss_array = []

for search_div in div_result:
    try:
        word1.send_keys(search_div[0])
        word2.send_keys(search_div[1])
        word3.send_keys(search_div[2])
    except:
        pass

    click_btn = driver.find_element_by_xpath("//input[@value='チェック']")
    driver.execute_script("arguments[0].click();", click_btn)

    sleep(10)

    base_html = driver.page_source.encode("utf-8")
    soup = BeautifulSoup(base_html, "html.parser")
    rank_result_div = soup.find(id="rank-result")
    res_table = rank_result_div.find("table")

    ### slack連携
    td = res_table.find_all("td")
    res_list = []
    last_post_text = ""

    for m, item in enumerate(td):
        res_list.append(item.get_text())
        mod = m % 3
        post_text = ""

        if item.get_text() != "\xa0":
            ss_array.append(item.get_text())

        if mod == 0:
            res_text = ""
            res_text += ("検索ワード:" + item.get_text()) + "\n"
        elif mod == 1:
            res_text += ("Google順位:" + item.get_text()) + "\n"
        elif mod == 2:
            res_text += ("Yahoo順位:" + item.get_text()) + "\n"
            post_text += res_text

        last_post_text += post_text
    word1.clear()
    word2.clear()
    word3.clear()

    eternal_text += last_post_text

# sleep(10)
print(ss_array)

### spreadsheet出力
div_result = list(split_list(ss_array, 3))
for item in div_result:

    print(item[0])
    print(item[1])
    print(item[2])

    # キーワードをセルから探す
    cell = worksheet.find(item[0])
    col_array = worksheet.col_values(1)
    row_array = worksheet.row_values(1)
    last_col_index = len(row_array)
    last_row_index = len(col_array)
    this_rank = ""

    if item[1] == "圏外":
        this_rank = "11"
    else:
        this_rank = item[1][:-1]

    try:
        # 順位を追加
        worksheet.update_cell(last_row_index + 1, cell.col, this_rank)
    except:
        # 新たな検索ワードをヘッダーに追加
        worksheet.update_cell(1, last_col_index + 1, item[0])
        # 順位を追加
        worksheet.update_cell(last_row_index + 1, last_col_index + 1, this_rank)

## 日付を追加
worksheet.update_cell(last_row_index + 1, 1, date)
# ### ----------

driver.close()

url = "https://slack.com/api/chat.postMessage"
headers = {"Authorization": "Bearer " + TOKEN}
data = {"channel": CHANNEL, "text": eternal_text}

r = requests.post(url, headers=headers, data=data)
# print("return ", r.json())
### ----------

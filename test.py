# 参考
# gspreadの使い方 https://zak-papa.com/python_gspread_cell_input_output
# google api https://tanuhack.com/operate-spreadsheet/

import gspread
import json
import datetime
from oauth2client.service_account import ServiceAccountCredentials 

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('./search-tracker-328413-2f7d16436ff4.json', scope)
gc = gspread.authorize(credentials)
SPREADSHEET_KEY = '1ziRNv4-S8ErU6MIy_6HIYQdb7sVFjc_-Kq4ieUdWiNM'
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

today = datetime.date.today()
year = today.year
month = today.month
day = today.day
date = str(year) + '/' + str(month) + '/' + str(day)

def split_list(l, n):
    """
    リストをサブリストに分割する
    :param l: リスト
    :param n: サブリストの要素数
    :return: 
    """
    for idx in range(0, len(l), n):
        yield l[idx:idx + n]

res_table = [
    '電動昇降デスク コスパ', 
    '1位', 
    '2位', 
    'テレワーク サボってしまう', 
    '1位', 
    '1位', 
    'Google順位', 
    '圏外', 
    '11位以下',
    'telotp', 
    '圏外', 
    '11位以下', 
    'seo', 
    '3位', 
    '11位以下', 
    ]
    

div_result = list(split_list(res_table, 3))
for item in div_result:
    
    #キーワードをセルから探す
    cell = worksheet.find(item[0])
    col_array = worksheet.col_values(1)
    row_array = worksheet.row_values(1)
    last_col_index = len(row_array)
    last_row_index = len(col_array)
    this_rank = ''

    if item[1] == '圏外':
        this_rank = '11'
    else:
        this_rank = item[1][:-1]

    try:
        # 順位を追加
        worksheet.update_cell(last_row_index + 1, cell.col, this_rank)
    except:
        # 新たな検索ワードをヘッダーに追加
        worksheet.update_cell(1 ,last_col_index + 1, item[0])
        # 順位を追加
        worksheet.update_cell(last_row_index + 1, last_col_index + 1, this_rank)
## 日付を追加
worksheet.update_cell(last_row_index + 1, 1, date)

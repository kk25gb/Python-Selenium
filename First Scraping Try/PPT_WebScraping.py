import requests
from bs4 import BeautifulSoup

# Google 搜尋 URL
google_url = 'https://www.ptt.cc/bbs/MobileComm/index.html'

# 查詢參數
# my_params = {'q': '寒流'}

# 下載 Google 搜尋結果
r = requests.get(google_url)

# 確認是否下載成功
if r.status_code == requests.codes.ok:
    # 以 BeautifulSoup 解析 HTML 原始碼
    soup = BeautifulSoup(r.text, 'html.parser')
    # print(soup.prettify())

    # 以 CSS 的選擇器來抓取 Google 的搜尋結果
    items = soup.select('#main-container > div.r-list-container.action-bar-margin.bbs-screen > div > div.title > a')
    for i in items:
        print("標題：" + i.text)
        print("網址：" + i.get('href'))

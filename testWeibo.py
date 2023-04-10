import json
import time

import xlwt
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
excel_name = "统计结果.xlsx"
wb = Workbook()
ws1 = wb.active
ws1.title='统计结果'

def get_html(url):
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0'}
    html = requests.get(url, headers=header).json()
    return html



def main():
    url = 'https://weibo.com/ajax/statuses/hot_band'
    html = get_html(url)
    htmlData = html['data']['band_list']

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Sheet1')
    ws.write(0, 0, '排名')
    ws.write(0, 1, '新闻')
    ws.write(0, 2, '分类')
    ws.write(0, 3, '搜索量')

    i = 0
    for col in htmlData:
        i = i+1
        ws.write(i, 0, i)
        ws.write(i, 1, col.get('word_scheme'))
        ws.write(i, 2, col.get('category'))
        ws.write(i, 3, col.get('raw_hot'))

    #now_time = time.strftime("%Y"+"年"+"%m"+"月"+"%d"+"日"+" %H:%M:%S", time.localtime())
    now_time = time.strftime("%Y年%m月%d日%H时%M分%S秒的", time.localtime())

    wb.save(now_time+'微博热搜.xls')

if __name__ == '__main__':
    main()
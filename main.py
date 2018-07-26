#-*-encoding:utf8-*-
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import os

def get_bs_obejct_by_url(url):
    html = requests.get(url)
    # print(html.encoding) # ISO-8859-1 인코딩나와서
    html.encoding = 'euc-kr'  # 한글 인코딩으로 변환
    return BeautifulSoup(html.text, 'lxml')

if __name__=="__main__":
    #===    CONFIG
    FILENAME = 'sample.xlsx'
    #===    DECLARE & DEFINE
    bs4 = get_bs_obejct_by_url('http://www.wooticket.com/popup_price.php')
    headList = []
    titleList = []
    buyList = []
    now = time.localtime()
    s = "%04d-%02d-%02d %02d:%02d:%02d" % \
        (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

    #===    CODE
    if os.path.isfile(FILENAME): # 파일있는 경우
        pass
    else: # 파일 없는 경우

        trs = bs4.find_all('table')[3].find_all("tr")[1:]
        for tr in trs[1:-1]:
            try:
                name = tr.find_all('td')[1].get_text().strip()
                titleList.append(name.split(' ', maxsplit=1)[1])
                headList.append(name.split(' ',maxsplit=1)[0])
            except:
                headList.append('-')
                titleList.append(name.split(' ', maxsplit=1)[0])
            finally:
                buy = tr.find_all('td')[2].find('font').get_text().strip()
                buyList.append(buy)
                sell = tr.find_all('td')[3].find('font').get_text().strip()
        print(len(titleList))
        print(len(headList))

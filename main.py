#-*-encoding:utf8-*-
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import os
def auto_fit_width():
    dims = {}
    for row in sheet1.rows:
        for cell in row:
            if cell.value:
                dims[cell.column] = max((dims.get(cell.column, 0), len(cell.value)))
    for col, value in dims.items():
        sheet1.column_dimensions[col].width = value

def get_bs_obejct_by_url(url):
    html = requests.get(url)
    # print(html.encoding) # ISO-8859-1 인코딩나와서
    html.encoding = 'euc-kr'  # 한글 인코딩으로 변환
    return BeautifulSoup(html.text, 'lxml')

def remove_dup_data_at_list(tmp):
    returnList = []
    returnList.append(tmp[0])
    for idx in range(1,len(tmp)):
        if tmp[idx-1] != tmp[idx]:
            returnList.append(tmp[idx])
        else:
            returnList.append('')
    return returnList
if __name__=="__main__":
    #===    CONFIG
    FILENAME = 'DATA.xlsx'
    #===    DECLARE & DEFINE
    bs4 = get_bs_obejct_by_url('http://www.wooticket.com/popup_price.php')
    headList = []   #맨위
    titleList = []  #그아래
    buyList = []    #매입
    sellList = []   #판매
    now = time.localtime()
    nowDate = time.strftime("%x")
    nowTime = "%04d-%02d-%02d %02d:%02d:%02d" % \
        (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

    header1 = ['일자', '긁어온 시간', '']
    header2 = [' ', ' ', ' ']
    datas = [nowDate,nowTime,' ']

    #===    CODE
    #파싱
    trs = bs4.find_all('table')[3].find_all("tr")[1:]
    for tr in trs[1:-1]:
        try:
            name = tr.find_all('td')[1].get_text().strip()
            titleList.append(name.split(' ', maxsplit=1)[1])
            headList.append(name.split(' ', maxsplit=1)[0])
        except:
            headList.append('-')
            titleList.append(name.split(' ', maxsplit=1)[0])
        finally:
            buy = tr.find_all('td')[2].find('font').get_text().strip()
            buyList.append(buy)
            sell = tr.find_all('td')[3].find('font').get_text().strip()
    if os.path.isfile(FILENAME): # 파일있는 경우
        wb = load_workbook(filename=FILENAME)
        sheet1 = wb.get_sheet_by_name(wb.get_sheet_names()[0])
        nextRow = sheet1.max_row + 1
        sheet1.append(datas + buyList)
        wb.save(FILENAME)
    else: # 파일 없는 경우
        # 엑셀파일 초기설정
        book = Workbook()
        sheet1 = book.active
        sheet1.column_dimensions['A'].width = 10
        sheet1.column_dimensions['B'].width = 20
        sheet1.column_dimensions['C'].width = 2
        sheet1.title = 'RawData'

        # 저장
        sheet1.append(header1 + remove_dup_data_at_list(headList))
        sheet1.append(header2 + titleList)
        sheet1.append(datas + buyList)
        book.save(FILENAME)

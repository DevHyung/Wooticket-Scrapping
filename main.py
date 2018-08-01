#-*-encoding:utf8-*-
from bs4 import BeautifulSoup
import requests
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import os

def valid_user():
    # 20180730 10:03기준 6시간
    now = 1533091486.2744226
    terminTime = now + 60 * 60 * 6
    print("체험판 만료기간 : ", time.ctime(terminTime))
    if time.time() > terminTime:
        print('만료되었습니다.')
        exit(-1)
    else:
        print(">>> 프로그램이 실행되었습니다.")


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
    valid_user()

    #===    CONFIG
    #FILENAME = r'C:\Users\khuph\Desktop\DATA.xlsx'
    f = open("CONFIG.txt",encoding='utf8').readlines()
    targetDir = f[0].split(':',maxsplit=1)[1].strip()
    targetName = f[1].split(':')[1].strip()
    FILENAME = targetDir+'\\'+targetName
    print(FILENAME)
    #===    DECLARE & DEFINE
    bs4 = get_bs_obejct_by_url('http://www.wooticket.com/popup_price.php')
    headList = []   #맨위
    titleList = []  #그아래
    buyList = []    #매입
    buyPerList = []
    sellList = []   #판매
    sellPerList = []
    spreadList = []
    now = time.localtime()
    nowDate = time.strftime("%x")
    nowTime = "%04d-%02d-%02d %02d:%02d:%02d" % \
        (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

    header1 = ['일자', '긁어온 시간', '']
    header2 = [' ', ' ', ' ']
    datas = [nowDate, nowTime, ' ']


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
            buyPer = tr.find_all('td')[2].get_text().strip().split('(')[1].split(')')[0]
            buyList.append(buy)
            buyPerList.append(buyPer)

            sell = tr.find_all('td')[3].find('font').get_text().strip()
            sellPer = tr.find_all('td')[3].get_text().strip().split('(')[1].split(')')[0]
            sellPerList.append(sellPer)
            sellList.append(sell)
            spreadList.append(int(sell.replace(',',''))-int(buy.replace(',','')))
    # 엑셀 저장 부분
    if os.path.isfile(FILENAME): # 파일있는 경우
        wb = load_workbook(filename=FILENAME)
        sheet1 = wb[wb.sheetnames[0]]
        #sheet1 = wb.get_sheet_by_name(wb.get_sheet_names()[0])
        nextRow = sheet1.max_row + 1
        sheet1.append(datas + buyList + header2 + sellList)

        sheet2 = wb[wb.sheetnames[1]]
        #sheet2 = wb.get_sheet_by_name(wb.get_sheet_names()[1])
        nextRow = sheet2.max_row + 1
        sheet2.append(datas + buyPerList + header2 + sellPerList)

        sheet3 = wb[wb.sheetnames[2]]
        #sheet3 = wb.get_sheet_by_name(wb.get_sheet_names()[2])
        nextRow = sheet3.max_row + 1
        sheet3.append(datas + spreadList)
        wb.save(FILENAME)
    else: # 파일 없는 경우
        # 엑셀파일 초기설정
        book = Workbook()
        # 시트 설정
        sheet1 = book.active
        sheet1.column_dimensions['A'].width = 10
        sheet1.column_dimensions['B'].width = 20
        sheet1.column_dimensions['C'].width = 2
        sheet1.title = 'RawData'

        sheet2 = book.create_sheet(title="Chg")
        sheet2.column_dimensions['A'].width = 10
        sheet2.column_dimensions['B'].width = 20
        sheet2.column_dimensions['C'].width = 2

        sheet3 = book.create_sheet(title="Spread")
        sheet3.column_dimensions['A'].width = 10
        sheet3.column_dimensions['B'].width = 20
        sheet3.column_dimensions['C'].width = 2
        # 저장
        sheet1.append(header1 + remove_dup_data_at_list(headList) + header2 + remove_dup_data_at_list(headList))
        sheet1.append(header2 + titleList + header2 + titleList)
        sheet1.cell(row=3, column=len(header2)+1).value = '매입가(원)'
        sheet1.cell(row=3, column=len(remove_dup_data_at_list(headList)) + 7).value = '판매가(원)' # 6+1
        sheet1.append(datas + buyList + header2 + sellList)

        sheet2.append(header1 + remove_dup_data_at_list(headList) + header2 + remove_dup_data_at_list(headList))
        sheet2.append(header2 + titleList + header2 + titleList)
        sheet2.cell(row=3, column=len(header2) + 1).value = '매입가 Chg.'
        sheet2.cell(row=3, column=len(remove_dup_data_at_list(headList)) + 7).value = '판매가 Chg.'  # 6+1
        sheet2.append(datas + buyPerList + header2 + sellPerList)


        sheet3.append(header1 + remove_dup_data_at_list(headList))
        sheet3.append(header2 + titleList)
        sheet3.cell(row=3, column=len(header2) + 1).value = 'Spread (판매가-매입가)'
        sheet3.append(datas + spreadList)

        book.save(FILENAME)

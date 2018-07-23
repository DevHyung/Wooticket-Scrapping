#-*-encoding:utf8-*-
from bs4 import BeautifulSoup
import requests
import xlsxwriter
if __name__=="__main__":
    url = 'http://www.wooticket.com/popup_price.php'
    html = requests.get(url)
    #print(html.encoding) # ISO-8859-1 인코딩나와서
    html.encoding = 'euc-kr'  # 한글 인코딩으로 변환
    bs4 = BeautifulSoup(html.text,'lxml')
    trs = bs4.find_all('table')[3].find_all("tr")[1:]
    with xlsxwriter.Workbook('test.xlsx') as workbook:
        # head
        headList = []
        titleList = []
        buyList = []
        worksheet = workbook.add_worksheet()
        row = 2
        col = 0
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
        worksheet.write_row(0, 0, headList)
        worksheet.write_row(1, 0, titleList)
        worksheet.write_row(2, 0, buyList)


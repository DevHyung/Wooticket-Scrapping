#-*-encoding:utf8-*-
from bs4 import BeautifulSoup
import requests

if __name__=="__main__":
    url = 'http://www.wooticket.com/popup_price.php'
    html = requests.get(url)
    #print(html.encoding) # ISO-8859-1 인코딩나와서
    html.encoding = 'euc-kr'  # 한글 인코딩으로 변환
    bs4 = BeautifulSoup(html.text,'lxml')
    trs = bs4.find_all('table')[3].find_all("tr")[1:]
    for tr in trs:
        try:
            name = tr.find_all('td')[1].get_text().strip()
            buy = tr.find_all('td')[2].find('font').get_text().strip()
            sell = tr.find_all('td')[3].find('font').get_text().strip()
            print(name,buy,sell)
        except:
            pass

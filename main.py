#-*-encoding:utf8-*-
from bs4 import BeautifulSoup
import requests

if __name__=="__main__":
    url = 'http://www.wooticket.com/popup_price.php'
    html = requests.get(url)
    #print(html.encoding) # ISO-8859-1 인코딩나와서
    html.encoding = 'euc-kr'  # 한글 인코딩으로 변환
    bs4 = BeautifulSoup(html.text,'lxml')
    print(bs4.prettify())
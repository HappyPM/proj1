#-*- coding: utf-8 -*-
import requests
from BeautifulSoup import BeautifulSoup # 왜 난 from bs4가 안되나.....??
#from bs4 import BeautifulSoup
from pymongo import MongoClient

# requests HTML, get document 
url = "http://www.krx.co.kr/por_kor/popup/JHPKOR13008.jsp"
r = requests.post(url, data={'mkt_typ':'S', 'market_gubun':'allVal'})

# BeautifulSoup HTML document parsing
soup = BeautifulSoup(r.text)
table = soup.find('table', {'id':'tbl1'})
trs = table.findAll('tr')

# making list after extracting data
stock_list = []

for tr in trs[1:]:
    stock = {}
    cols = tr.findAll('td')
    stock['code'] = cols[0].text[1:]
    stock['name'] = cols[1].text.replace(";", "")
    #stock['full_code'] = cols[2].text
    stock_list.append(stock)


# find sector info and add to the list
sectorBaseUrl = "http://finance.naver.com/item/main.nhn?code="

def findSectorMarketTypeInfo(company):
    response = requests.get(sectorBaseUrl + company);
    soup = BeautifulSoup(response.text)
    #print(sectorBaseUrl + company)
    
    sector = ""
    h4 = soup.find('h4', {'class':'h_sub sub_tit7'})
    if h4 is not None:
        sector = h4.a.text
        
    marketType = ""
    marketTypeKospi = soup.find('img', {'class':'kospi'})
    marketTypeKosdaq = soup.find('img', {'class':'kosdaq'})
     
    if marketTypeKospi is not None:
        marketType = 'kospi'
    elif marketTypeKosdaq is not None:
        marketType = 'kosdaq'
    else:
        marketType = "none"
                
    return sector,marketType
    

stock_sector_list = []
for s in stock_list[1:]:
    sector,mtype = findSectorMarketTypeInfo(s['code'])
    s['sector'] = sector
    s['marketType'] = mtype

   
#for s in stock_list[1:15]:
#    print s

 
client = MongoClient()
db = client.hpm
coll = db.company

coll.insert(stock_list)

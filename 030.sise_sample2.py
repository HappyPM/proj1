#-*- coding: utf-8 -*-
#http://finance.naver.com/item/frgn.nhn?code=014530&page=1

from bs4 import BeautifulSoup
import urllib2
from pymongo import MongoClient

# page
gnStockCode             = '014530';         # 주식 코드 번호
gnEndDay                = '2010.01.01'      # 이 날짜까지 데이터 구함
gnPageLoopCount         = 100;              # 취합할 Page 개수, 100개 이네에서 날짜에 걸려서 종료됨.


gnStockPricDays = []                            # DB에 insert할 list
gnTableFirstDay = '';                       # 과거날짜 -> loop 종료용 


def MakeTable(nCode, nPageEntry):
    
    global gnTableFirstDay

    nPageIndex = nPageEntry + 1;

    opener = urllib2.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

    nUrl = "http://finance.naver.com/item/frgn.nhn?code=" + nCode + "&page=" + str(nPageIndex);
    response = opener.open(nUrl)
    page = response.read()

    stSoupDays =  BeautifulSoup(page).find('table', {'class':'type2', 'width':'680'}).findAll('tr', {'onmouseover':'mouseOver(this)'})

    if len(stSoupDays)  == 0:   # 항목 조회 안되면 루프 종료함
        return False
    
    stTableFirstDay = stSoupDays[0].findAll('td')[0].text;  # 더이상 페이지가 없을 경우 똑같은 날자가 나타남 loop 종료 체크함
    if gnTableFirstDay == stTableFirstDay:
        return False
    gnTableFirstDay = stTableFirstDay
    
    for stSoupDay in stSoupDays:
        stItems = stSoupDay.findAll('td')
        stStockPric = {}
        #stStockPric['code'] = nCode
        stStockPric['day']  = stItems[0].text
        stStockPric['pric'] = stItems[1].text.replace(',', '')
        if gnEndDay > stStockPric['day'] : # 최종날짜 체크함 
            return False        
        gnStockPricDays.append(stStockPric)

    return True        


for nPageIndex in range(3):
    if MakeTable(gnStockCode, nPageIndex) == False :
        break;


siseDoc = {}
siseDoc['code'] = gnStockCode
siseDoc['days'] = gnStockPricDays # embedded document 구조로 저장하기 


client = MongoClient()      # MongoDB 
db = client.hpm             # hpm dbs
coll = db.sise              # sise collection
coll.insert(siseDoc)


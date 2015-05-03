#-*- coding: utf-8 -*-
import requests
from BeautifulSoup import BeautifulSoup
from pymongo import MongoClient


url = 'http://www.krx.co.kr/por_kor/popup/JHPKOR13008.jsp'
r = requests.post(url, data={'mkt_typ':'S', 'market_gubun': 'allVal'})

soup = BeautifulSoup(r.text)
table = soup.find('table', {'id':'tbl1'})
trs = table.findAll('tr')

stock_list = []

for tr in trs[1:]:
    stock = {}
    cols = tr.findAll('td')
    stock['code'] = cols[0].text[0:]
    stock['name'] = cols[1].text.replace(";", "")
    stock_list.append(stock)
    
    

client = MongoClient()
db = client.hpm
coll = db.company

coll.insert(stock_list)


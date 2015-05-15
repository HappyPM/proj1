#-*- coding: utf-8 -*-
import requests
import json
from pymongo import MongoClient
try:
    from BeautifulSoup import BeautifulSoup 
except ImportError:
    from bs4 import BeautifulSoup              # exception  import 


def get_days_to_json(soup):
    script = soup.findAll('script')[4].string
    day = script.split("changeFin = ", 1)[1].split(";",1)[0]
    soup = BeautifulSoup(day)
    day = soup.text
    day = json.loads(day)    
    return day
    #print script


def get_data_to_json(soup):
    script = soup.findAll('script')[4].string
    data = script.split("changeFinData = ", 1)[1].split(";",1)[0]
    data = json.loads(data)    
    return data
    #print script    
    
code = '014530'


# page2
url = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd=' + code
r = requests.get(url)
soup = BeautifulSoup(r.text)


days = get_days_to_json(soup)
data = get_data_to_json(soup)

              

# embedded document 구조로 저장하기 
finance_list = [];
finance  = {};
finance["code"] = code
finance_list.append(finance);

#print finan_list
year_data_list = [];
quater_data_list = [];



year_day = days[0]
quater_day = days[1]

for data1 in data:
    
    yy_dat = data1[0]    
    qt_dat = data1[1]    

    jj = 0
    for yy_dat1 in yy_dat:
        
        dnam = yy_dat1[0]
        
        qt_dat1 = qt_dat[jj]
        jj = jj + 1

        ii = 0            
        for yy_dat2 in yy_dat1[1:]:
            #print len(qt_dat1[ii])
            qt_dat2 = qt_dat1[ii]                
            
            year_data = {}            
            year_data["day"] = year_day[ii]
            year_data["item_name"] = dnam
            year_data["item_value"] = yy_dat2.replace(',', '')
            year_data_list.append(year_data);
                
            quater_data = {}            
            quater_data["day"] = quater_day[ii]
            quater_data["item_name"] = dnam
            quater_data["item_value"] = qt_dat2.replace(',', '')
            quater_data_list.append(quater_data);
            
            ii = ii + 1;
            #print iid

 
client = MongoClient()
db = client.hpm
coll = db.finance

coll.insert(finance_list)
coll.update({"code": code}, { "$push": {"year_date" : year_data_list, "quater_data" : quater_data_list } })



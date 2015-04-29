#-*- coding: utf-8 -*-
import requests
import json
from BeautifulSoup import BeautifulSoup
from pymongo import MongoClient



def get_sector(soup):
    cominfo = soup.find('table', {'id':'comInfo'}).findAll('span', {'class':'exp'})[1].text[7:]
    return cominfo



def get_profit_per(soup):
   ss = soup.find('table', {'id':'cTB26'}).text
   return ss
   

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

# page1
url = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=' + code
r = requests.get(url)
soup1 = BeautifulSoup(r.text)

# page2
url = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd=' + code
r = requests.get(url)
soup2 = BeautifulSoup(r.text)



#print get_sector(soup1, code);
#print get_profit_per(soup2)

days = get_days_to_json(soup2)
data = get_data_to_json(soup2)


    

def set_json_data():
    
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
                jd = {"code": code, "period": "year", "dday": year_day[ii], "item_name": dnam, "item_value": yy_dat2}
                coll.insert_one(jd).inserted_id
                jd = {"code": code, "period": "quater", "dday": quater_day[ii], "item_name": dnam, "item_value": qt_dat2}
                coll.insert_one(jd).inserted_id
                ii = ii + 1;
                #print iid
    
        
        
client = MongoClient()
db = client.hpm
coll = db.finance




year_day = days[0]
quater_day = days[1]

set_json_data()


#print year_day
#print quater_day

#for day1 in days:
#    n = n + 1
#    if n == 1:
#        dbas = 'year'
#    else:
#        dbas = 'quarter'
#    for day2 in day1:
#       print dbas, day2

       #set_json_data(dbas, day2)
       # for data1 in data:
       
            
            
            #post_id = posts.insert_one(jd).inserted_id
            #print post_id
            
        
    
    
 
 

















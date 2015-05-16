#-*- coding: utf-8 -*-
import json
import urllib2
from pymongo import MongoClient
try:
    from BeautifulSoup import BeautifulSoup 
except ImportError:
    from bs4 import BeautifulSoup              # exception  import 


gnUrl = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd='
gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define


_client = MongoClient()
_db = _client.hpm
gnCompanyColl = _db.company
gnFinanceColl = _db.finance

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

def set_year_and_quater(days, data, year_data_list, quater_data_list) :
        
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
                
                #print quater_data
                quater_data_list.append(quater_data);
                
                ii = ii + 1;
                #print iid
    

def insert_finance(ncode):
    nUrl = gnUrl + ncode    
    nResponse = gnOpener.open(nUrl)
    nPage = nResponse.read()
    # page2
    # r = requests.get(gnUrl)
    nSoup = BeautifulSoup(nPage)

    days = get_days_to_json(nSoup)
    data = get_data_to_json(nSoup)

    #print days
    # embedded document 구조하기 위해서 root 에 code 저장
    finance_list = [];
    finance  = {};
    finance["code"] = ncode
    finance_list.append(finance);
    
    #print finance_list
    gnFinanceColl.insert(finance_list)

    # year, quater  1차 json 데이터를 finance 구조에 맞게 다시 저장함 
    year_data_list = [];
    quater_data_list = [];
   
    
    set_year_and_quater(days, data, year_data_list, quater_data_list)
    gnFinanceColl.update({"code": ncode}, { "$push": {"year" : year_data_list, "quater" : quater_data_list } })
    
    print ncode
     
     
def load_company_all():
    _companys = gnCompanyColl.find();
    return _companys;
    

#############3 main ################################

#insert_finance('192400')

nCompanys = load_company_all()
for company in nCompanys[:]
    insert_finance(company['code'][1:])
    
    








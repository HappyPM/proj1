#-*- coding: utf-8 -*-
import json
import urllib2
from pymongo import MongoClient
try:
    from BeautifulSoup import BeautifulSoup 
except ImportError:
    from bs4 import BeautifulSoup              # exception  import 
import re


# 제무제표 class
class Finance:
    #__url = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd=' # 과거 url
    __url = "http://companyinfo.stock.naver.com/v1/company/ajax/cF1001.aspx?fin_typ=0"
    __opener = urllib2.build_opener()
    __opener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define


    '''
    freq_type : A 모두 , Y : 연도, Q : 분기 
    code : 종목코드 (6자리)
    '''
    def __init__(self, freq_typ = "A", code = "000660"):
        
       
        self._freq_typ = freq_typ
        if self._freq_typ == "A" or self._freq_typ == "Y": 
            self._yeardata = []
            self._set_finance(code, "Y", self._yeardata)
        if self._freq_typ == "A" or self._freq_typ == "Q":
            self._quaterdata = []
            self._set_finance(code, "Q", self._quaterdata)            
                      

    
    def _set_finance(self, code, freq_typ, dataset):
        #print (code)
        url = Finance.__url + "&freq_typ" + freq_typ + "&cmp_cd=" +code   
        response = Finance.__opener.open(url)
        page = response.read()
        self._soup = BeautifulSoup(page)
        self._set_json_data(freq_typ, dataset)
        
    def _set_json_data(self, freq_typ, dataset):       
        # 날짜 (예상치 제외)
        days = self._soup.findAll("th", {"class":re.compile("r03c0[1-5] bg ")})
        # 재무정보 타이틀
        item_names = self._soup.findAll("th", {"class":"bg txt title "})        
        # 재무정보 값 
        item_values = self._soup.findAll("td", {"class":"num line "})        
        
        #print(type(vals))
        for day in days:
            kk = 0        
            for item_name in item_names:
                for ii in range(0,5):
                    item_value = item_values[kk*5+ii]
                    data = {}            
                    data["day"] = day.text
                    data["item_name"]  = item_name.text
                    data["item_value"] = item_value.text
                    dataset.append(data)
                    
    def get_data(self, freq_typ = "Y"):
        if freq_typ == "Y" :
            return self._yeardata
        elif freq_typ == "Q":
            return self._quaterdata

                

## main 

f = Finance("A", "000660")
print(f.get_data("Y")) # 연간데이터 
print(f.get_data("Q")) # 분기데이터 






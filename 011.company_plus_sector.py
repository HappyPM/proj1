import requests
import json
from bs4 import BeautifulSoup

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
    stock['full_code'] = cols[2].text
    stock_list.append(stock)

# list save to JSON format
j = json.dumps(stock_list)
with open('./krx_symbols.json', 'w') as f:
    f.write(j)

# review the saved file
fn = 'krx_symbols.json'
with open(fn, 'r') as f:
    stock_list = json.load(f)

# Test print
#for s in stock_list[:10]:
#    print s['full_code'], s['code'][1:], s['name']


# find sector info and add to the list
sectorBaseUrl = "http://finance.naver.com/item/main.nhn?code="

def findSectorInfo(company):
    response = requests.get(sectorBaseUrl + company);
    soup = BeautifulSoup(response.text)
    #print(sectorBaseUrl + company)
    
    sector = ""
    h4 = soup.find('h4', {'class':'h_sub sub_tit7'})
    if h4 is not None:
        sector = h4.a.text
    return sector


stock_sector_list = []
for s in stock_list[:10]:
    newStock = {}
    newStock['code'] = s['code']
    newStock['name'] = s['name']
    newStock['full_code'] = s['full_code']
    newStock['sector'] = findSectorInfo(s['code'])
    stock_sector_list.append(newStock)

# Test print
for s in stock_sector_list[:10]:
     print s['full_code'], s['code'][1:], s['name'], s['sector']




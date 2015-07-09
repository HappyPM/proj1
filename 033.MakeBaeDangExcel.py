#-*- coding: utf-8 -*-
import requests;
import copy;
import json;
import urllib2;
from bs4 import BeautifulSoup;
import xlsxwriter;

gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

# Stock 이름과 코드를 얻는 함수
gastStockList = {};
def COMPANY_GetStockCode(astStockList): # OUT (gastStockList: 종목 이름 / 코드)
    nUrl = 'http://www.krx.co.kr/por_kor/popup/JHPKOR13008.jsp';
    nRequest = requests.post(nUrl, data={'mkt_typ':'S', 'market_gubun': 'allVal'});

    nSoup = BeautifulSoup(nRequest.text);
    stTable = nSoup.find('table', {'id':'tbl1'});
    astTrs = stTable.findAll('tr');

    for stTr in astTrs[1:]:
        stStock = {};

        cols = stTr.findAll('td')
        nStockCode = cols[0].text[0:].split("A")[1];
        nStockName = cols[1].text.replace(";", "");
        gastStockList[nStockName] = nStockCode;

gastStockNameCode = [];
def COMPANY_GetNameToCode(astStockList, astStockName, astStockNameCode):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    stStockNameCode = {};
    nStockOffset = 0;
    nStockCount = len(astStockName);

    for nStockIndex in range(nStockCount):
        stStockNameCode['Name'] = astStockName[nStockIndex];
        if (astStockList.has_key(stStockNameCode['Name']) == False):
            continue;
        stStockNameCode['Code'] = astStockList[stStockNameCode['Name']];
        stStockNameCode['SISE'] = 0;
        stStockNameCode['Count'] = 0;

        astStockNameCode.append(0);
        astStockNameCode[nStockOffset] = copy.deepcopy(stStockNameCode);
        nStockOffset = nStockOffset + 1;

gastStockName = [];
def COMPANY_GetStockName(astStockName, nMaxStockCount):
    nMaxPageRange = 10;

    for nPageIndex in range(nMaxPageRange):
        if (len(astStockName) >= nMaxStockCount):
            break;

        anUrl = 'http://finance.naver.com/sise/dividend_list.nhn?sosok=KOSPI&fsq=20144&field=divd_rt&ordering=desc&page=' + str(nPageIndex + 1);
        stResponse = gnOpener.open(anUrl);
        stPage = stResponse.read();
        stSoup = BeautifulSoup(stPage);

        astTr = stSoup.findAll('tr');
        nPageTrLen = len(astTr);

        nSkipStrLen = len(astTr[2].text);

        for nTrIndex in range(nPageTrLen):
            if (nTrIndex <= 2):
                continue;

            nStrLen = len(astTr[nTrIndex].text);
            if (nStrLen <= nSkipStrLen):
                continue;

            astStockType = astTr[nTrIndex].text.split("\n");
            nStockName = astStockType[1];
            astStockName.append(nStockName);

            if ((len(astStockName) % 50) == 0):
                break;
            if (len(astStockName) >= nMaxStockCount):
                break;

def get_days_to_json(soup):
    script = soup.findAll('script')[4].string
    day = script.split("changeFin = ", 1)[1].split(";",1)[0]
    soup = BeautifulSoup(day)
    day = soup.text
    day = json.loads(day)    
    return day


def get_data_to_json(soup):
    script = soup.findAll('script')[4].string
    data = script.split("changeFinData = ", 1)[1].split(";",1)[0]
    data = json.loads(data)    
    return data

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
    

gastYearDataList = [];
gastQuaterDataList = [];
def COMPANY_GetFinance(ncode, astYearDataList, astQuaterDataList):
    nCodeUrl = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=';
    nCodeUrl = nCodeUrl + ncode;
    nResponse = gnOpener.open(nCodeUrl);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);
    tables = nSoup.findAll('table');
    astStrip = tables[1].text.split(' ');
    nStockCode = astStrip[1];

    nFinanceUrl = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd=';
    nUrl = nFinanceUrl + nStockCode;
    nResponse = gnOpener.open(nUrl);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);

    stDays = get_days_to_json(nSoup);
    stData = get_data_to_json(nSoup);

    set_year_and_quater(stDays, stData, astYearDataList, astQuaterDataList);

gstWorkBook = xlsxwriter.Workbook('BaeDangStockList.xlsx');
gstWorkSheet = gstWorkBook.add_worksheet();
def GetSplitTitle(stString):
#    stString = stString.split(u"(지배)")[0];
#    stString = stString.split(u"(비지배)")[0];
#    stString = stString.split(u"(%)")[0];
#    stString = stString.split(u"(원)")[0];
#    stString = stString.split(u"(배)")[0];
#    stString = stString.split(u"(보통주)")[0];
#    stString = stString.split(u"활동현금흐름")[0];
#    stString = stString.split(u"계속사업이익")[0];
#    stString = stString.split(u"발생부채")[0];
    stString = stString.split(u"(IFRS연결)")[0];
    stString = stString.split(u"(IFRS별도)")[0];
    return stString;

def SetXlsxTitle(astYearDataList, astQuaterDataList):
    nXlsxColumnOffset = 0;
    nRowOffset = 1;
    nColOffset = 0;
    nXlsxYear = 0;
    nXlsxQuarter= 0;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});

    gstWorkSheet.write(0, nColOffset, u"종목명", stPurpleFormat);
    nColOffset = nColOffset + 1;
    gstWorkSheet.write(0, nColOffset, u"코드번호", stGrayFormat);
    nColOffset = nColOffset + 1;

    nLength = len(astYearDataList);
    for nYearIndex in range(nLength):
        stYearDataList = astYearDataList[nYearIndex];
        stItemName = GetSplitTitle(stYearDataList["item_name"]);
        stDay = GetSplitTitle(stYearDataList["day"]);
        stThisYear = stDay.split('/')[0];
        if (nYearIndex == 0):
            nXlsxYear = stThisYear;
        if (nXlsxYear == stThisYear):
            gstWorkSheet.write(0, nColOffset, u"연간 " + stItemName, stRedTitleFormat);

        gstWorkSheet.write(nRowOffset, nColOffset, stThisYear, stTitleFormat);
        nColOffset = nColOffset + 1;

    nLength = len(astQuaterDataList);
    for nQuarterIndex in range(nLength):
        stQuaterDataList = astQuaterDataList[nQuarterIndex];
        stItemName = GetSplitTitle(stQuaterDataList["item_name"]);
        stDay = GetSplitTitle(stQuaterDataList["day"]);
        stThisQuarter = stDay.split('/')[1];
        if (nQuarterIndex == 0):
            nXlsxQuarter = stDay.split('/')[1];
        if (nXlsxQuarter == stThisQuarter):
            gstWorkSheet.write(0, nColOffset, u"분기 " + stItemName, stGreenTitleFormat);

        gstWorkSheet.write(nRowOffset, nColOffset, stThisQuarter, stTitleFormat);
        nColOffset = nColOffset + 1;

def SetXlsxData(nRowOffset, stStockNameCode, astYearDataList, astQuaterDataList):
    nColOffset = 0;

    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'font_color': 'gray'});

    gstWorkSheet.write(nRowOffset, nColOffset, stStockNameCode['Name'], stPurpleFormat);
    nColOffset = nColOffset + 1;
    gstWorkSheet.write(nRowOffset, nColOffset, stStockNameCode['Code'], stGrayFormat);
    nColOffset = nColOffset + 1;

    for stYearDataList in astYearDataList:
        gstWorkSheet.write(nRowOffset, nColOffset, stYearDataList["item_value"]);
        nColOffset = nColOffset + 1;

    for stQuaterDataList in astQuaterDataList:
        gstWorkSheet.write(nRowOffset, nColOffset, stQuaterDataList["item_value"]);
        nColOffset = nColOffset + 1;

def GetFinanceAndWriteExcel(astStockNameCode):
    gastYearDataList = [];
    gastQuaterDataList = [];

#    COMPANY_GetFinance('005935', gastYearDataList, gastQuaterDataList);
#    SetXlsxTitle(gastYearDataList, gastQuaterDataList);

    nStockLen = len(astStockNameCode);
    for nStockIndex in range(nStockLen):
        gastYearDataList = [];
        gastQuaterDataList = [];
        COMPANY_GetFinance(astStockNameCode[nStockIndex]['Code'], gastYearDataList, gastQuaterDataList);
        if (nStockIndex == 0):
            SetXlsxTitle(gastYearDataList, gastQuaterDataList);
        SetXlsxData(nStockIndex + 2, astStockNameCode[nStockIndex], gastYearDataList, gastQuaterDataList);
        nStockIndex = nStockIndex;

############# main #############

gnMaxBaeDangStockCount = 500;
#GetFinanceAndWriteExcel(gastStockNameCode);

COMPANY_GetStockName(gastStockName, gnMaxBaeDangStockCount);
COMPANY_GetStockCode(gastStockList);
COMPANY_GetNameToCode(gastStockList, gastStockName, gastStockNameCode);
GetFinanceAndWriteExcel(gastStockNameCode);
gstWorkBook.close();

gnMaxBaeDangStockCount = gnMaxBaeDangStockCount;

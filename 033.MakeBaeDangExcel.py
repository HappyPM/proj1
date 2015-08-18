#-*- coding: utf-8 -*-
import requests;
import pandas as pd;
import pandas.io.data as web;
from StringIO import StringIO;
import datetime;
import copy;
import json;
import urllib2;
from bs4 import BeautifulSoup;
import xlsxwriter;
from xlsxwriter.utility import xl_rowcol_to_cell;


gnMaxBaeDangStockCount = 500;
gnMaxGraphStockCount = 30;
gbPrintProgress = 1;


gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

def GetTodayString():
    stNow = datetime.datetime.now();

    if (stNow.year > 2015):
        exit();

    stDate = str(stNow.year)[2:];
    if (stNow.month < 10):
        stDate = stDate + '0';
    stDate = stDate + str(stNow.month);
    if (stNow.day < 10):
        stDate = stDate + '0';
    stDate = stDate + str(stNow.day);

    return stDate;

# Stock 이름과 코드를 얻는 함수
gastChangeStockNameCodeList = {};
def COMPANY_GetStockCode(astStockList): # OUT (gastChangeStockNameCodeList: 종목 이름 / 코드)
    PrintProgress(u"[시작] 종목 코드 취합");
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
        astStockList[nStockName] = nStockCode;
    PrintProgress(u"[완료] 종목 코드 취합");


gastStockNameCodeInfor = [];
def COMPANY_GetNameToCode(nStockCode, astStockList, astStockName, astStockNameCode):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    PrintProgress(u"[시작] " + nStockCode + u" 종목 코드 변환");
    stStockNameCode = {};
    nStockOffset = len(astStockNameCode);
    nStockCount = len(astStockName);

    for nStockIndex in range(nStockCount):
        stStockNameCode['Name'] = astStockName[nStockIndex];
        if (astStockList.has_key(stStockNameCode['Name']) == False):
            continue;
        stStockNameCode['Code'] = astStockList[stStockNameCode['Name']];
        stStockNameCode['Type'] = nStockCode;
        stStockNameCode['SISE'] = 0;
        stStockNameCode['Count'] = 0;

        astStockNameCode.append(0);
        astStockNameCode[nStockOffset] = copy.deepcopy(stStockNameCode);
        nStockOffset = nStockOffset + 1;
    PrintProgress(u"[완료] " + nStockCode + u" 종목 코드 변환");

gastKospiStockName = [];
gastKosdaqStockName = [];
def COMPANY_GetStockName(nStockCode, astStockName, nMaxStockCount):
    PrintProgress(u"[시작] " + nStockCode + u" 종목 리스트 취합");

    anBaseUrl = {};
    anBaseUrl['KOSPI']      = 'http://finance.naver.com/sise/dividend_list.nhn?sosok=KOSPI&fsq=20144&field=divd_rt&ordering=desc&page=';
    anBaseUrl['KOSDAQ']     = 'http://finance.naver.com/sise/dividend_list.nhn?sosok=KOSDAQ&fsq=20144&field=divd_rt&ordering=desc&page=';

    nMaxPageRange = int(nMaxStockCount / 50) + 1;

    for nPageIndex in range(nMaxPageRange):
        if (len(astStockName) >= nMaxStockCount):
            break;

        anUrl = anBaseUrl[nStockCode] + str(nPageIndex + 1);
        stResponse = gnOpener.open(anUrl, timeout=60);
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

            # 혹시라도 동일 종목이 존재하면 skip
            nStockLen = len(astStockName);
            if ((nStockLen > 0) and (astStockName[nStockLen - 1] == nStockName)):
                continue;

            astStockName.append(nStockName);
            PrintProgress(u"[진행] 종목 리스트 취합: " + nStockName);

            if ((len(astStockName) % 50) == 0):
                break;
            if (len(astStockName) >= nMaxStockCount):
                break;

    PrintProgress(u"[완료] " + nStockCode + u" 종목 리스트 취합");


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

gastYearDataList = [];
gastQuaterDataList = [];
def COMPANY_GetFinance(ncode, astYearDataList, astQuaterDataList):
    nFinanceUrl = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd=';
    nUrl = nFinanceUrl + ncode;
    nResponse = gnOpener.open(nUrl, timeout=60);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);

    stDays = get_days_to_json(nSoup);
    stData = get_data_to_json(nSoup);

    set_year_and_quater(stDays, stData, astYearDataList, astQuaterDataList);

def GetSplitTitle(stString):
    stString = stString.split(u"(IFRS연결)")[0];
    stString = stString.split(u"(IFRS별도)")[0];
    return stString;

def COMPANY_SetDummyStockInfor(stStockInfor, tables, nType, nName, nCode):
    stStockInfor['Name'] = nName;
    stStockInfor['Code'] = nCode;
    stStockInfor['Type'] = nType;
    stStockInfor['WebCode'] = "0";
    stStockInfor['시세'] = {};

def COMPANY_SetStockInfor(stStockInfor, tables, nType, nName, nCode):
    astTable4 = tables[4].text.split('/');
    astTable42 = astTable4[2].split(u'\n');
    stStockInfor['CurPrice'] = astTable42[1].replace(',', '');
    stStockInfor['CurPrice'] = stStockInfor['CurPrice'].replace(' ', '');

    stStockInfor['Name'] = nName;
    stStockInfor['Code'] = nCode;
    stStockInfor['Type'] = nType;
    astSplit = tables[1].text.split(' | ');

    stSplit0 = astSplit[0].split(' ');
    nSplit0Len = len(stSplit0);
    stStockInfor['WebCode'] = stSplit0[nSplit0Len - 1];

#    stSplit2 = astSplit[2].split(' : ');
#    stStockInfor['종목Type'] = stSplit2[0];

    stSplit3 = astSplit[3].split('|');
    stSplit30 = stSplit3[0].split('EPS');
    stSplit300 = stSplit30[0].split(':');
    stStockInfor['WICS'] = stSplit300[1];

    stSplit301 = stSplit30[1].split(u'\xa0');
    stStockInfor['EPS'] = stSplit301[0].replace(',', '');
    stStockInfor['EPS'] = stStockInfor['EPS'].replace(' ', '');

    stSplit31 = stSplit3[1].split(u'\xa0');
    stSplit311 = stSplit31[1].split(' ');
    stStockInfor['BPS'] = stSplit311[1].replace(',', '');
    stStockInfor['BPS'] = stStockInfor['BPS'].replace(' ', '');

    stSplit32 = stSplit3[2].split(u'\xa0');
    stSplit321 = stSplit32[1].split(' ');
    stStockInfor['PER'] = stSplit321[1].replace(',', '');
    stStockInfor['PER'] = stStockInfor['PER'].replace(' ', '');

    stSplit34 = stSplit3[4].split(u'\xa0');
    stSplit341 = stSplit34[1].split(' ');
    stStockInfor['PBR'] = stSplit341[1].replace(',', '');
    stStockInfor['PBR'] = stStockInfor['PBR'].replace(' ', '');

    stSplit35 = stSplit3[5].split(u'\xa0');
    stSplit351 = stSplit35[1].split(' ');
    stSplit3511 = stSplit351[1].split(u'결산기');
    stSplit35110 = stSplit3511[0].replace('%', '');
    stStockInfor['배당률'] = stSplit35110.replace(',', '');
    stStockInfor['배당률'] = stStockInfor['배당률'].replace(' ', '');

    astTable4Len = len(astTable4);
    astTable4_1M = astTable4[astTable4Len - 4].split(u'\n');
    astTable4_3M = astTable4[astTable4Len - 3].split(u'\n');
    astTable4_6M = astTable4[astTable4Len - 2].split(u'\n');
    astTable4_1Y = astTable4[astTable4Len - 1].split(u'\n');

    stStockInfor['1M'] = astTable4_1M[len(astTable4_1M) - 1].replace('\r', '').replace(' ', '').replace(',', '');
    stStockInfor['3M'] = astTable4_3M[len(astTable4_3M) - 1].replace('\r', '').replace(' ', '').replace(',', '');
    stStockInfor['6M'] = astTable4_6M[len(astTable4_6M) - 1].replace('\r', '').replace(' ', '').replace(',', '');
    stStockInfor['1Y'] = astTable4_1Y[0].replace('\r', '').replace(' ', '').replace(',', '');

    astYearDataList = [];
    astQuaterDataList = [];
    COMPANY_GetFinance(stStockInfor['WebCode'], astYearDataList, astQuaterDataList);
    stStockInfor['YearDataList'] = astYearDataList;
    stStockInfor['QuaterDataList'] = astQuaterDataList;

    stStockInfor['시세'] = {};
    SISE_GetStockInfor(nCode, nType, stStockInfor['시세']);
    
def COMPANY_GetStockFinanceInfor(nType, nName, nCode, astStockInfor):
    stStockInfor = {};
    nCodeUrl = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=';
    nCodeUrl = nCodeUrl + nCode;
    nResponse = gnOpener.open(nCodeUrl, timeout=60);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);
    tables = nSoup.findAll('table');

    if (len(tables) > 0):
        COMPANY_SetStockInfor(stStockInfor, tables, nType, nName, nCode);
    else:
        COMPANY_SetDummyStockInfor(stStockInfor, tables, nType, nName, nCode);

    astStockInfor.append(stStockInfor);

gastStockInfor = [];
def COMPANY_GetFinanceInfor(astStockNameCode, astStockInfor):
    nStockLen = len(astStockNameCode);
    PrintProgress(u"[시작] 종목 정보 취합: " + str(0) + " / " + str(nStockLen));

    for nStockIndex in range(nStockLen):
        COMPANY_GetStockFinanceInfor(astStockNameCode[nStockIndex]['Type'],
                                        astStockNameCode[nStockIndex]['Name'],
                                        astStockNameCode[nStockIndex]['Code'],
                                        astStockInfor);
        PrintProgress(u"[진행] 종목 정보 취합: " + str(nStockIndex + 1) + " / " + str(nStockLen) + " - " + astStockNameCode[nStockIndex]['Name']);
    PrintProgress(u"[완료] 종목 정보 취합: " + str(nStockLen) + " / " + str(nStockLen));

def SetFnXlsxTitle(astStockInfor):
    stStockInfor = astStockInfor[0];
    nStockLen = len(astStockInfor);
    nXlsxColumnOffset = 0;
    nRowOffset = 1;
    nColOffset = 0;
    nXlsxYear = 0;
    nXlsxQuarter= 0;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue', 'align':'center'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stChoiceFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'black', 'align':'center'});

    gstFnSheet.write(0, nColOffset, u"종목선정", stChoiceFormat);
    stStartCell = xl_rowcol_to_cell(2, nColOffset);
    stEndCell = xl_rowcol_to_cell(2 + nStockLen - 1, nColOffset);
    gstFnSheet.write(1, nColOffset, "=count(" + stStartCell + ":" + stEndCell + ")", stChoiceFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(0, nColOffset, u"종목명", stPurpleFormat);
    nColOffset = nColOffset + 1;

#    gstFnSheet.write(0, nColOffset, u"종목Type", stNavyFormat);
#    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"코드번호", stGrayFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"시세연결", stGrayFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"WICS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"현재가격", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"PER", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"PBR", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"BPS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"EPS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"배당률", stNavyFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(0, nColOffset, u"최근수익률", stNavyFormat);
    gstFnSheet.write(1, nColOffset, u"1M", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(1, nColOffset, u"3M", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(1, nColOffset, u"6M", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(1, nColOffset, u"1Y", stTitleFormat);
    nColOffset = nColOffset + 1;

    nLength = len(stStockInfor['YearDataList']);
    for nYearIndex in range(nLength):
        stYearDataList = stStockInfor['YearDataList'][nYearIndex];
        stItemName = GetSplitTitle(stYearDataList["item_name"]);
        stDay = GetSplitTitle(stYearDataList["day"]);
        stThisYear = stDay.split('/')[0];
        if (nYearIndex == 0):
            nXlsxYear = stThisYear;
        if (nXlsxYear == stThisYear):
            gstFnSheet.write(0, nColOffset, u"연간 " + stItemName, stRedTitleFormat);

        gstFnSheet.write(nRowOffset, nColOffset, stThisYear, stTitleFormat);
        nColOffset = nColOffset + 1;

    nLength = len(stStockInfor['QuaterDataList']);
    for nQuarterIndex in range(nLength):
        stQuaterDataList = stStockInfor['QuaterDataList'][nQuarterIndex];
        stItemName = GetSplitTitle(stQuaterDataList["item_name"]);
        stDay = GetSplitTitle(stQuaterDataList["day"]);
        stThisQuarter = stDay.split('/')[1];
        if (nQuarterIndex == 0):
            nXlsxQuarter = stDay.split('/')[1];
        if (nXlsxQuarter == stThisQuarter):
            gstFnSheet.write(0, nColOffset, u"분기 " + stItemName, stGreenTitleFormat);

        gstFnSheet.write(nRowOffset, nColOffset, stThisQuarter, stTitleFormat);
        nColOffset = nColOffset + 1;

def SetSiseXlsxTitle(astStockInfor):
    nXlsxColumnOffset = 0;
    nRowOffset = 0;
    nColOffset = 0;
    nXlsxYear = 0;
    nXlsxQuarter= 0;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});

    gstSiseSheet.write(0, nColOffset, u'날짜', stPurpleBoldFormat);
    nRowOffset = nRowOffset + 1;
    nRowOffset = nRowOffset + 1;

    nLength = len(astStockInfor);
    for nDayIndex in range(nLength):
        stStockInfor = astStockInfor[nDayIndex];
        gstSiseSheet.write(nRowOffset, nColOffset, stStockInfor['Date'], stPurpleFormat);
        nRowOffset = nRowOffset + 1;

def SetFnXlsxData(nRowOffset, astStockInfor, nStockIndex):
    stStockInfor = astStockInfor[nStockIndex];
    nColOffset = 1;
    nCodeUrl = 'http://finance.naver.com/item/main.nhn?code=';
    nSiseUrl = u'internal:' + gstSiseSheetName + u'!';
    nStartSiseColOffset = 5;

    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stOrangeFormat = gstWorkBook.add_format({'font_color': 'orange'});
    stGrayFormat = gstWorkBook.add_format({'font_color': 'gray', 'underline':  1});

    # 종목명
    if (stStockInfor['Type'] == 'KOSPI'):
        gstFnSheet.write(nRowOffset, nColOffset, stStockInfor['Name'], stPurpleFormat);
    else:
        gstFnSheet.write(nRowOffset, nColOffset, stStockInfor['Name'], stOrangeFormat);
    nColOffset = nColOffset + 1;

    # 코드번호
    stCell = xl_rowcol_to_cell(nRowOffset, nColOffset);
    gstFnSheet.write_url(stCell, nCodeUrl + stStockInfor['Code'], stGrayFormat, stStockInfor['Code']);
    nColOffset = nColOffset + 1;

    # 시세연결
    stLinkCell = xl_rowcol_to_cell(0, nStartSiseColOffset + (2 * nStockIndex));
    stTargetColOffset = str(nStartSiseColOffset + 1 + (2 * nStockIndex));
    gstFnSheet.write(nRowOffset, nColOffset, nSiseUrl + stLinkCell, stGrayFormat, stTargetColOffset);
    nColOffset = nColOffset + 1;

    # 재무 Page 비정상 예외 처리
    if (stStockInfor['WebCode'] == "0"):
        return;

    # WICS
    gstFnSheet.write(nRowOffset, nColOffset, stStockInfor['WICS']);
    nColOffset = nColOffset + 1;

    if (stStockInfor['CurPrice'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['CurPrice']));
    nColOffset = nColOffset + 1;

    if (stStockInfor['PER'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['PER']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['PBR'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['PBR']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['BPS'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['BPS']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['EPS'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['EPS']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['배당률'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['배당률']));
    nColOffset = nColOffset + 1;

    if (stStockInfor['1M'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['1M']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['3M'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['3M']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['6M'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['6M']));
    nColOffset = nColOffset + 1;
    if (stStockInfor['1Y'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['1Y']));
    nColOffset = nColOffset + 1;

    nLength = len(stStockInfor['YearDataList']);
    for nYearIndex in range(nLength):
        stYearDataList = stStockInfor['YearDataList'][nYearIndex];
        if (stYearDataList["item_value"] != u''):
            gstFnSheet.write(nRowOffset, nColOffset, float(stYearDataList["item_value"]));
        nColOffset = nColOffset + 1;

    nLength = len(stStockInfor['QuaterDataList']);
    for nQuaterIndex in range(nLength):
        stQuaterDataList = stStockInfor['QuaterDataList'][nQuaterIndex];
        if (stQuaterDataList["item_value"] != u''):
            gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]));
        nColOffset = nColOffset + 1;

def SetKospiXlsxData(nColOffset, nType, astStockInfor, astBaseInfor):
    nRowOffset = 0;
    bFirstPrice = 0;
    nCurPrice = 0;
    nCurRate = 0;
    nPrevPrice = 1;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stSiseFormat = gstWorkBook.add_format({'num_format':'0.00'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstSiseSheet.write(nRowOffset, nColOffset, nType, stTitleFormat);
    nRowOffset = nRowOffset + 1;

    gstSiseSheet.write(nRowOffset, nColOffset, u'증감율', stRedTitleFormat);
    gstSiseSheet.write(nRowOffset, nColOffset + 1, u'시세', stGreenTitleFormat);
    nRowOffset = nRowOffset + 1;

    # 시세 출력
    nBaseLength = len(astBaseInfor);
    nStockLength = len(astStockInfor);
    for nBaseIndex in range(nBaseLength):
        for nDayIndex in range(nBaseIndex, nStockLength):
            if (astBaseInfor[nBaseIndex]['Date'] == astStockInfor[nDayIndex]['Date']):
                break;

        stStockInfor = astStockInfor[nDayIndex];
        nCurPrice = stStockInfor['Price'];

        if (bFirstPrice == 0):
            bFirstPrice = 1;
        else:
            nCurRate = float((float(nCurPrice) * 100) / float(nPrevPrice)) - 100;
            gstSiseSheet.write(nRowOffset, nColOffset, nCurRate, stRateFormat);

        gstSiseSheet.write(nRowOffset, nColOffset + 1, nCurPrice, stSiseFormat);

        nPrevPrice = nCurPrice;
        nRowOffset = nRowOffset + 1;

def SetSiseXlsxData(nColOffset, astKospiInfor, stStockInfor):
    nRowOffset = 0;
    nKospiIndex = 0;
    astSiseStockInfor = stStockInfor['시세'];
    bFirstPrice = 0;
    nCurPrice = 0;
    nCurRate = 0;
    nPrevPrice = 1;
    nImpossibleRate = 30;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstSiseSheet.write(nRowOffset, nColOffset, stStockInfor['Name'], stNavyFormat);
    nRowOffset = nRowOffset + 1;

    gstSiseSheet.write(nRowOffset, nColOffset, u'증감율', stRedTitleFormat);
    gstSiseSheet.write(nRowOffset, nColOffset + 1, u'시세', stGreenTitleFormat);
    nRowOffset = nRowOffset + 1;

    # 재무 Page 비정상 예외 처리
    if (stStockInfor['WebCode'] == "0"):
        return;

    # 시세 출력
    nKospiLength = len(astKospiInfor);
    for nKospiIndex in range(nKospiLength):
        if (astSiseStockInfor.has_key(astKospiInfor[nKospiIndex]['Date'])):
            nCurPrice = astSiseStockInfor[astKospiInfor[nKospiIndex]['Date']];

            if (bFirstPrice == 0):
                bFirstPrice = 1;
            else:
                nCurRate = float((float(nCurPrice) * 100) / float(nPrevPrice)) - 100;
                if (nCurRate > nImpossibleRate) or (nCurRate < (nImpossibleRate * -1)):
                    nCurRate = 0;

                gstSiseSheet.write(nRowOffset, nColOffset, nCurRate, stRateFormat);

            gstSiseSheet.write(nRowOffset, nColOffset + 1, nCurPrice);

            nPrevPrice = nCurPrice;
        nRowOffset = nRowOffset + 1;


# 누적승리 출력
def PrintWinningRate(nColOffset, nTitle, nMaxDateCount):
    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstGraphSheet.write(0, nColOffset, nTitle, stTitleBoldFormat);
    gstGraphSheet.write(1, nColOffset, u"누적 승리", stRedTitleBoldFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + 2;

        stAccumulatedCell = xl_rowcol_to_cell(nDateRowOffset - 1, nColOffset);
        if (nTitle == u"KOSPI"):
            stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 1);
        else:
            stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 2);
        stKospiRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 3);

        stString = "=IFERROR(" + stAccumulatedCell + " + (" + stAvgStockRate + " - " + stKospiRate + "), \"\")";

        gstGraphSheet.write(nDateRowOffset, nColOffset, stString, stRateFormat);

def SetGraphXlsxData(nMaxDateCount, nMaxStockCount):
    nMaxRowOffset = nMaxDateCount + 2;
    nStartFnRowOffset = 3;
    nEndFnRowOffset = nStartFnRowOffset + nMaxStockCount - 1;
    stStartFnRowOffset = str(nStartFnRowOffset);
    stEndFnRowOffset = str(nEndFnRowOffset);

    stSiseCell = gstSiseSheetName + u'!';
    nRowOffset = 0;
    nColOffset = 0;

    nDateColOffset = 0;
    nKospiColOffset = 1;
    nKosdaqColOffset = 2;
    nAvgStockColOffset = 3;
    nKospiVsColOffset = 4;
    nKosdaqVsColOffset = 5;
    nStockColOffset = 6;

    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stTitleFormat = gstWorkBook.add_format({'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'font_color': 'green'});
    stPurpleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    # 날짜 / KOSPI
    for nRowOffset in range(nMaxRowOffset):
        stTransCell = xl_rowcol_to_cell(nRowOffset, nDateColOffset);
        stString = stSiseCell + stTransCell;
        stDateString = u'=' + "IF(" + stString + " > 0," + stString + ", \"\")";
        if (nRowOffset < 2):
            gstGraphSheet.write(nRowOffset, nDateColOffset, stDateString, stPurpleBoldFormat);
        else:
            gstGraphSheet.write(nRowOffset, nDateColOffset, stDateString, stPurpleFormat);


        stTransCell = xl_rowcol_to_cell(nRowOffset, nKospiColOffset);
        stKospiString = u'=' + stSiseCell + stTransCell;
        if (nRowOffset < 1):
            gstGraphSheet.write(nRowOffset, nKospiColOffset, stKospiString, stTitleFormat);
        elif (nRowOffset < 2):
            gstGraphSheet.write(nRowOffset, nKospiColOffset, stKospiString, stGreenTitleFormat);
        else:
            gstGraphSheet.write(nRowOffset, nKospiColOffset, stKospiString, stRateFormat);

        stTransCell = xl_rowcol_to_cell(nRowOffset, nKosdaqColOffset + 1);
        stKospiString = u'=' + stSiseCell + stTransCell;
        if (nRowOffset < 1):
            gstGraphSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stTitleFormat);
        elif (nRowOffset < 2):
            gstGraphSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stGreenTitleFormat);
        else:
            gstGraphSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stRateFormat);

    # 평균 증감
    gstGraphSheet.write(0, nAvgStockColOffset, u"종목 평균", stNavyFormat);
    gstGraphSheet.write(1, nAvgStockColOffset, u"증감율", stGreenTitleFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + 2;
        stStartTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset);
        stEndTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nMaxStockCount - 1);
        stString = "=IFERROR(AVERAGE(" + stStartTransCell + ":" + stEndTransCell + "), \"\")";
        gstGraphSheet.write(nDateRowOffset, nAvgStockColOffset, stString, stRateFormat);


    # KOSPI 누적승리
    PrintWinningRate(nKospiVsColOffset, u"KOSPI", nMaxDateCount);
    PrintWinningRate(nKosdaqVsColOffset, u"KOSDAQ", nMaxDateCount);


    # 선정 종목 (그래프 취합 50개 제한)
    nStockCount = nMaxStockCount;
    if (nStockCount > gnMaxGraphStockCount):
        nStockCount = gnMaxGraphStockCount;
    for nStockIndex in range(nStockCount):
        for nRowOffset in range(nMaxRowOffset):
            stStockColOffset = str(nStockIndex + 1);
            stSiseRowOffset = str(nRowOffset + 1);

            stString = "=IFERROR(INDIRECT(ADDRESS(" + stSiseRowOffset + ", INDIRECT(ADDRESS(2 + MATCH(" + stStockColOffset + ", ";
            stString += gstFnSheetName + "!$A$" + stStartFnRowOffset + ":$A$" + stEndFnRowOffset + ", 0), 4, 4, 5, \"" + gstFnSheetName + "\")), ";
            stString += "4, 5, \"" + gstSiseSheetName + "\")), \"\")";

            if (nRowOffset >= 2):
                gstGraphSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString, stRateFormat);
            else:
                gstGraphSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

    # 차트 출력
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartFnRowOffset - 1, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(2, nKospiVsColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKospiVsColOffset);
    stKospiData = '=' + gstGraphSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartTransCell = xl_rowcol_to_cell(2, nKosdaqVsColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKosdaqVsColOffset);
    stKosdaqData = '=' + gstGraphSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartDateCell = xl_rowcol_to_cell(2, nDateColOffset);
    stEndDateCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nDateColOffset);
    stDate = '=' + gstGraphSheetName + '!' + stStartDateCell + ":" + stEndDateCell;

    stTitle = xl_rowcol_to_cell(1, nKospiVsColOffset);
    stChart.set_title({'name':u"KOSPI / KOSDAQ 대비 누적 승리율"});
    stChart.set_x_axis({'name':u'날짜'});
    stChart.set_y_axis({'name':u'승리율(%)', 'min':0, 'max':100});

    stChart.add_series({'name':u"KOSPI",'categories':stDate, 'values':stKospiData});
    stChart.add_series({'name':u"KOSDAQ",'categories':stDate, 'values':stKosdaqData});
    stChart.set_size({'width':720, 'height':504});
    gstGraphSheet.insert_chart(stGraphCell, stChart);

def COMPANY_WriteExcelFile(astKospiInfor, astKosdaqInfor, astStockInfor):
    PrintProgress(u"[시작] 엑셀 취합");
    nColOffset = 0;
    nRowOffset = 0;

    # 시세 Title 출력
    SetSiseXlsxTitle(astKospiInfor);
    nColOffset = nColOffset + 1;
    SetKospiXlsxData(nColOffset, 'KOSPI', astKospiInfor, astKospiInfor);
    nColOffset = nColOffset + 2;
    SetKospiXlsxData(nColOffset, 'KOSDAQ', astKosdaqInfor, astKospiInfor);
    nColOffset = nColOffset + 2;
    PrintProgress(u"[진행] 시세 Title 출력");

    # 시세 데이터 출력
    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        SetSiseXlsxData(nColOffset, astKospiInfor, astStockInfor[nStockIndex]);
        nColOffset = nColOffset + 2;
        PrintProgress(u"[진행] 시세 데이터 출력: " + str(nStockIndex + 1) + " / " + str(nStockLen) + " - " + astStockInfor[nStockIndex]['Name']);

    # 재무 Title 출력
    SetFnXlsxTitle(astStockInfor);
    nRowOffset = nRowOffset + 2;
    PrintProgress(u"[진행] 재무 Title 출력");

    # 재무 데이터 출력
    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        SetFnXlsxData(nRowOffset, astStockInfor, nStockIndex);
        nRowOffset = nRowOffset + 1;
        PrintProgress(u"[진행] 재무 데이터 출력: " + str(nStockIndex + 1) + " / " + str(nStockLen) + " - " + astStockInfor[nStockIndex]['Name']);

    # 그래프 출력
    SetGraphXlsxData(len(astKospiInfor), len(astStockInfor));
    PrintProgress(u"[진행] 그래프 출력");
    PrintProgress(u"[완료] 엑셀 취합");

# Date & 가격을 얻는 함수 (코스피 / 코스닥 / 일반종목)
# 코스피 or 코스닥 or 일반 종목 선택
#gnStockCode             = 'KOSPI';      # '1997-07-01' ~
#gnStockCode             = 'KOSDAQ';     # '2013-03-04' ~
#gnStockCode             = '014530';     # '2000-0101' ~
gastStockInfor          = [];
def SISE_GetNonStockInfor(nStockCode, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    stDataInfor = {};

    anReqCode               = {};
    anReqCode['KOSPI']      = '^KS11';
    anReqCode['KOSDAQ']     = '^KQ11';

    # nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=0&b=1&c=1900';

    # Month = a + 1 / Day = b / Year = c
    nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=11&b=30&c=2013';
    stRequest               = requests.get(nUrl);
    stDataInfor             = pd.read_csv(StringIO(stRequest.content), index_col='Date', parse_dates={'Date'});

    for nIndex in range(stDataInfor.shape[0]):
        stStock             = {};
        stStock['Date']     = stDataInfor.index[nIndex]._date_repr[2:]; # 날짜
        stStock['Price']    = stDataInfor.values[nIndex][3];            # 종가: 'Close'
        stStockInfor.append(stStock);

def SISE_GetErrorStockInfor(nStockCode, nStockType, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    stDataInfor = {};

    anReqCode               = {};
    anReqCode['KOSPI']      = '.KS';
    anReqCode['KOSDAQ']     = '.KQ';

#        stStartDate             = datetime.datetime(1900, 1, 1);
    stStartDate             = datetime.datetime(2013, 12, 30);
    stDataInfor             = web.DataReader(nStockCode + anReqCode[nStockType], "yahoo", stStartDate);

    for nIndex in range(stDataInfor.shape[0]):
        stStockInfor[stDataInfor.index[nIndex]._date_repr[2:]]   = stDataInfor.values[nIndex][3];

def SISE_GetStockInfor(nStockCode, nStockType, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    nCodeUrl = 'http://www.etomato.com/home/itemAnalysis/ItemPrice.aspx?item_code=';
    nCodeUrl = nCodeUrl + nStockCode;
    nResponse = gnOpener.open(nCodeUrl, timeout=60);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);
    tables = nSoup.findAll('table');

    if ((len(tables) <= 13) or (len(tables[13].text.split(u'창출')) <= 1)):
        SISE_GetErrorStockInfor(nStockCode, nStockType, stStockInfor);
        return;

    astTable = tables[20].contents;
    nTableLen = len(astTable);
    for nIndex in range(3, nTableLen):
        nLen = len(astTable[nIndex]);
        if (nLen <= 1):
            continue;

        astSplit = astTable[nIndex].text.split(u'\n');
        stDate = astSplit[1][2:].replace(".", "-");
        stStockInfor[stDate] = int(astSplit[2].replace(",", ""));
        if (stDate == u"13-12-30"):
            break;

gastKospiInfor      = [];
gastKosdaqInfor      = [];
def SISE_GetKospiInfor(astKospiInfor, astKosdaqInfor):
    PrintProgress(u"[시작] KOSPI / KOSDAQ 정보 취합");
    SISE_GetNonStockInfor('KOSPI', astKospiInfor);
    PrintProgress(u"[진행] KOSPI 정보 취합");
    SISE_GetNonStockInfor('KOSDAQ', astKosdaqInfor);
    PrintProgress(u"[진행] KOSDAQ 정보 취합");
    astKospiInfor.sort();
    astKosdaqInfor.sort();
    PrintProgress(u"[완료] KOSPI / KOSDAQ 정보 취합");

def PrintProgress(stString):
    if (gbPrintProgress > 0):
        print (stString);

############# main #############

gstDate = GetTodayString();

# Kospi / Kosdaq 정보 취합
SISE_GetKospiInfor(gastKospiInfor, gastKosdaqInfor);

# 종목 정보 취합
COMPANY_GetStockName(u'KOSPI', gastKospiStockName, gnMaxBaeDangStockCount);
COMPANY_GetStockName(u'KOSDAQ', gastKosdaqStockName, gnMaxBaeDangStockCount);
COMPANY_GetStockCode(gastChangeStockNameCodeList);
COMPANY_GetNameToCode(u'KOSPI', gastChangeStockNameCodeList, gastKospiStockName, gastStockNameCodeInfor);
COMPANY_GetNameToCode(u'KOSDAQ', gastChangeStockNameCodeList, gastKosdaqStockName, gastStockNameCodeInfor);
COMPANY_GetFinanceInfor(gastStockNameCodeInfor, gastStockInfor);

# 종목 정보 출력
gstWorkBookName     = u'BaeDangStockList_' + gstDate + u'.xlsx';
gstFnSheetName      = u'FN' + gstDate;
gstSiseSheetName    = u'시세' + gstDate;
gstGraphSheetName   = u'그래프' + gstDate;
gstWorkBook         = xlsxwriter.Workbook(gstWorkBookName);
gstFnSheet          = gstWorkBook.add_worksheet(gstFnSheetName);
gstSiseSheet        = gstWorkBook.add_worksheet(gstSiseSheetName);
gstGraphSheet       = gstWorkBook.add_worksheet(gstGraphSheetName);

COMPANY_WriteExcelFile(gastKospiInfor, gastKosdaqInfor, gastStockInfor);

gstFnSheet.autofilter('A2:JG2');
gstFnSheet.freeze_panes('C3');
gstSiseSheet.freeze_panes('F3');
gstGraphSheet.freeze_panes('G3');

PrintProgress(u"[시작] 엑셀 출력");
gstWorkBook.close();
PrintProgress(u"[완료] 엑셀 출력");
PrintProgress(u"Complete all process");

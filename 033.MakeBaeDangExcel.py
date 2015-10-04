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
import time;


gnMaxBaeDangStockCount = 1000;


gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

gnGetBaeDangStockCount = int(gnMaxBaeDangStockCount + (gnMaxBaeDangStockCount * 0.2) + 1);
if (gnGetBaeDangStockCount <= 50):
    gnGetBaeDangStockCount = 50;
gnMaxGraphStockCount = gnMaxBaeDangStockCount;
gnMaxKospiStockCount = gnMaxBaeDangStockCount / 2;
gnMaxKosdaqStockCount = gnMaxBaeDangStockCount / 2;
if ((gnMaxBaeDangStockCount % 2) > 0):
    gnMaxKospiStockCount = gnMaxKospiStockCount + 1;
gbPrintProgress = 1;

ganYear = [1];
ganMonth = [1];
ganDay = [1];
def GetTodayString(anYear, anMonth, anDay):
    stNow = datetime.datetime.now();

    if (stNow.year > 2015):
        exit();

    stDate = str(stNow.year)[2:];
    anYear[0] = stNow.year;

    if (stNow.month < 10):
        stDate = stDate + '0';
    stDate = stDate + str(stNow.month);
    anMonth[0] = stNow.month;

    if (stNow.day < 10):
        stDate = stDate + '0';
    stDate = stDate + str(stNow.day);
    anDay[0] = stNow.day;

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

def GetUrlOpen(anUrl):
    stResponse = {};

    while True:
        try:
            stResponse = gnOpener.open(anUrl, timeout=60);
        except:
            time.sleep(1);
        else:
            break;

    return stResponse;

gastStockName = {};
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
        stResponse = GetUrlOpen(anUrl);
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
            if (len(nStockName) == 0):
                continue;
            elif ((nStockLen > 0) and (gastStockName.has_key(nStockName))):
                continue;

            astStockName.append(nStockName);
            gastStockName[nStockName] = 0;
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

            nYearIndicator = 0;
            nQuaterIndicator = 0;
            nMultipleIndicator = 1;

            ii = 0            

            astYearIndicatorData = [];
            astQuaterIndicatorData = [];
            stYearAppendData = {};
            stQuaterAppendData = {};

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

                if (ii > 0):
                    nIndicator = nMultipleIndicator;
                    for stYearIndicatorData in astYearIndicatorData:
                        if ((year_data["item_value"] != '') and (stYearIndicatorData["item_value"] != '') and (float(year_data["item_value"]) > 0) and (float(year_data["item_value"]) >= float(stYearIndicatorData["item_value"]))):
                            nYearIndicator = nYearIndicator + nIndicator;
                        nIndicator = nIndicator * 2;
                    nIndicator = nMultipleIndicator;
                    for stQuaterIndicatorData in astQuaterIndicatorData:
                        if ((quater_data["item_value"] != '') and (stQuaterIndicatorData["item_value"] != '') and (float(quater_data["item_value"]) > 0) and (float(quater_data["item_value"]) >= float(stQuaterIndicatorData["item_value"]))):
                            nQuaterIndicator = nQuaterIndicator + nIndicator;
                        nIndicator = nIndicator * 2;
                        
                    nMultipleIndicator = nMultipleIndicator * 10;
                astYearIndicatorData.append(year_data);
                astQuaterIndicatorData.append(quater_data);

                ii = ii + 1;
            stYearAppendData["day"] = u"지표/";
            stYearAppendData["item_name"] = dnam;
            stYearAppendData["item_value"] = unicode(nYearIndicator);
            year_data_list.append(stYearAppendData);
            stQuaterAppendData["day"] = u"/지표";
            stQuaterAppendData["item_name"] = dnam;
            stQuaterAppendData["item_value"] = unicode(nQuaterIndicator);
            quater_data_list.append(stQuaterAppendData);

gastYearDataList = [];
gastQuaterDataList = [];
def COMPANY_GetFinance(ncode, astYearDataList, astQuaterDataList):
    nFinanceUrl = 'http://companyinfo.stock.naver.com/v1/company/cF1001.aspx?finGubun=MAIN&cmp_cd=';
    nUrl = nFinanceUrl + ncode;
    nResponse = GetUrlOpen(nUrl);
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
    if (stStockInfor['WebCode'] != nCode):
        return 0;

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
    stStockInfor['수익률지표'] = int(0);
    if (stStockInfor['1M'] != ''):
        if (stStockInfor['3M'] != ''):
            if (float(stStockInfor['3M']) >= float(stStockInfor['1M'])):
                stStockInfor['수익률지표'] = stStockInfor['수익률지표'] + 1;
                
            if (stStockInfor['6M'] != ''):
                if (float(stStockInfor['6M']) >= float(stStockInfor['1M'])):
                    stStockInfor['수익률지표'] = stStockInfor['수익률지표'] + 10;
                if (float(stStockInfor['6M']) >= float(stStockInfor['3M'])):
                    stStockInfor['수익률지표'] = stStockInfor['수익률지표'] + 20;
                    
                if (stStockInfor['1Y'] != ''):
                    if (float(stStockInfor['1Y']) >= float(stStockInfor['1M'])):
                        stStockInfor['수익률지표'] = stStockInfor['수익률지표'] + 100;
                    if (float(stStockInfor['1Y']) >= float(stStockInfor['3M'])):
                        stStockInfor['수익률지표'] = stStockInfor['수익률지표'] + 200;
                    if (float(stStockInfor['1Y']) >= float(stStockInfor['6M'])):
                        stStockInfor['수익률지표'] = stStockInfor['수익률지표'] + 400;

    astYearDataList = [];
    astQuaterDataList = [];
    COMPANY_GetFinance(stStockInfor['WebCode'], astYearDataList, astQuaterDataList);
    stStockInfor['YearDataList'] = astYearDataList;
    stStockInfor['QuaterDataList'] = astQuaterDataList;

    stStockInfor['시세'] = {};
    bRet = SISE_GetStockInfor(nCode, nType, stStockInfor['시세']);
    return bRet;

def COMPANY_GetStockFinanceInfor(nType, nName, nCode, astStockInfor):
    stStockInfor = {};
    nCodeUrl = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=';
    nCodeUrl = nCodeUrl + nCode;
    nResponse = GetUrlOpen(nCodeUrl);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);
    tables = nSoup.findAll('table');

    if (len(tables) == 0):
        return 0;

    bRet = COMPANY_SetStockInfor(stStockInfor, tables, nType, nName, nCode);

    if (bRet > 0):
        astStockInfor.append(stStockInfor);

    return bRet;

gastStockInfor = [];
def COMPANY_GetFinanceInfor(astStockNameCode, astStockInfor):
    nMaxGettingCount = gnMaxBaeDangStockCount;
    nKospiCount = 0;
    nKosdaqCount = 0;

    nStockLen = len(astStockNameCode);
    PrintProgress(u"[시작] 종목 정보 취합: " + str(0) + " / " + str(nMaxGettingCount));

    for nStockIndex in range(nStockLen):
        if ((astStockNameCode[nStockIndex]['Type'] == u'KOSPI') and (nKospiCount >= gnMaxKospiStockCount)):
            continue;
        elif ((astStockNameCode[nStockIndex]['Type'] == u'KOSDAQ') and (nKosdaqCount >= gnMaxKosdaqStockCount)):
            continue;

        bRet = COMPANY_GetStockFinanceInfor(astStockNameCode[nStockIndex]['Type'],
                                        astStockNameCode[nStockIndex]['Name'],
                                        astStockNameCode[nStockIndex]['Code'],
                                        astStockInfor);
        # astStockInfor에 종목이 추가 안됨.
        if (bRet == 0):
            continue;

        if (astStockNameCode[nStockIndex]['Type'] == u'KOSPI'):
            nKospiCount = nKospiCount + 1;
        elif (astStockNameCode[nStockIndex]['Type'] == u'KOSDAQ'):
            nKosdaqCount = nKosdaqCount + 1;

        PrintProgress(u"[진행] 종목 정보 취합: " + str(nKospiCount + nKosdaqCount) + " / " + str(nMaxGettingCount) + " - " + astStockNameCode[nStockIndex]['Name']);

    PrintProgress(u"[완료] 종목 정보 취합: " + str(nStockLen) + " / " + str(nStockLen));

gstAutoFilterStartCell  = 'A2';
gstAutoFilterEndCell    = 'A2';
def SetFnXlsxTitle(astStockInfor):
    stStockInfor = astStockInfor[0];
    nStockLen = len(astStockInfor);
    nXlsxColumnOffset = 0;
    nRowOffset = 0;
    nColOffset = 0;
    nXlsxYear = 0;
    nXlsxQuarter= 0;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue', 'align':'center'});
    stIndicatorFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'brown', 'align':'center'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stChoiceFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'black', 'align':'center'});

    gstFnSheet.write(nRowOffset, nColOffset, u"종목매핑", stPurpleFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(nRowOffset, nColOffset, u"종목선정", stChoiceFormat);
    stStartCell = xl_rowcol_to_cell(2, nColOffset);
    stEndCell = xl_rowcol_to_cell(2 + nStockLen - 1, nColOffset);
    gstFnSheet.write(nRowOffset + 1, nColOffset, "=count(" + stStartCell + ":" + stEndCell + ")", stChoiceFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(nRowOffset, nColOffset, u"종목명", stPurpleFormat);
    nColOffset = nColOffset + 1;

#    gstFnSheet.write(nRowOffset, nColOffset, u"종목Type", stNavyFormat);
#    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"코드번호", stGrayFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"시세연결", stGrayFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"WICS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"현재가격", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"PER", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"PBR", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"BPS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"EPS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset, nColOffset, u"배당률", stNavyFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(nRowOffset, nColOffset, u"최근수익률", stNavyFormat);
    gstFnSheet.write(nRowOffset + 1, nColOffset, u"1M", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset + 1, nColOffset, u"3M", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset + 1, nColOffset, u"6M", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset + 1, nColOffset, u"1Y", stTitleFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(nRowOffset + 1, nColOffset, u"지표", stIndicatorFormat);
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
            gstFnSheet.write(nRowOffset, nColOffset, u"연간 " + stItemName, stRedTitleFormat);

        if (stThisYear == u'지표'):
            gstFnSheet.write(nRowOffset + 1, nColOffset, stThisYear, stIndicatorFormat);
        else:
            gstFnSheet.write(nRowOffset + 1, nColOffset, stThisYear, stTitleFormat);

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
            gstFnSheet.write(nRowOffset, nColOffset, u"분기 " + stItemName, stGreenTitleFormat);

        if (stThisQuarter == u'지표'):
            gstFnSheet.write(nRowOffset + 1, nColOffset, stThisQuarter, stIndicatorFormat);
        else:
            gstFnSheet.write(nRowOffset + 1, nColOffset, stThisQuarter, stTitleFormat);
        nColOffset = nColOffset + 1;

    stAutoFilterCell = xl_rowcol_to_cell(1, nColOffset - 1);
    return stAutoFilterCell;

def SetSiseXlsxTitle(astStockInfor):
    nXlsxColumnOffset = 0;
    nRowOffset = 1;
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

    gstSiseSheet.write(nRowOffset, nColOffset, u'날짜', stPurpleBoldFormat);
    nRowOffset = nRowOffset + 1;
    nRowOffset = nRowOffset + 1;

    nLength = len(astStockInfor);
    for nDayIndex in range(nLength):
        stStockInfor = astStockInfor[nDayIndex];
        gstSiseSheet.write(nRowOffset, nColOffset, stStockInfor['Date'], stPurpleFormat);
        nRowOffset = nRowOffset + 1;

def SetFnXlsxMapping(nRowOffset, nColOffset):
    nStartRow = 3;
    nEndRow = gnMaxBaeDangStockCount + nStartRow - 1;
    stStockChoiceLocation = u'B';
    stStockNameLocation = u'C';
    nTargetRowOffset = nRowOffset + 1;

    stString = u'';
    stString = stString + u'=IF(ISNUMBER(' + stStockChoiceLocation + str(nTargetRowOffset) + u'),';
    stString = stString + u'count(' + stStockChoiceLocation + str(nStartRow) + u':' + stStockChoiceLocation + str(nTargetRowOffset) + u'),';
    stString = stString + u'IFERROR(MATCH(' + stStockNameLocation + str(nTargetRowOffset) + u',' + stStockChoiceLocation + str(nStartRow) + u':' + stStockChoiceLocation + str(nEndRow) + u',0), \"\"))';
    gstFnSheet.write(nRowOffset, nColOffset, stString);

def SetFnXlsxData(nRowOffset, astStockInfor, nStockIndex):
    stStockInfor = astStockInfor[nStockIndex];
    nColOffset = 0;
    nCodeUrl = 'http://finance.naver.com/item/main.nhn?code=';
    nSiseUrl = u'internal:' + gstSiseSheetName + u'!';
    nStartSiseColOffset = 5;
    nLinkRowOffset = 1;

    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stOrangeFormat = gstWorkBook.add_format({'font_color': 'orange'});
    stGrayFormat = gstWorkBook.add_format({'font_color': 'gray', 'underline':  1});
    stIndicator1Format = gstWorkBook.add_format({'bg_color': '#FFFF00'});
    stIndicator2Format = gstWorkBook.add_format({'bg_color': '#FFFFDF'});
    stIndicator3Format = gstWorkBook.add_format({'bg_color': '#FFFFEF'});

    # 종목매핑
    SetFnXlsxMapping(nRowOffset, nColOffset);
    nColOffset = nColOffset + 1;
    nColOffset = nColOffset + 1;

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
    stLinkCell = xl_rowcol_to_cell(nLinkRowOffset, nStartSiseColOffset + (2 * nStockIndex));
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
    if (stStockInfor['수익률지표'] != u''):
        if (float(stStockInfor['수익률지표']) >= 700):
            if (float(stStockInfor['수익률지표']) >= 730):
                gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['수익률지표']), stIndicator1Format);
            elif (float(stStockInfor['수익률지표']) >= 710):
                gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['수익률지표']), stIndicator2Format);
            else:
                gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['수익률지표']), stIndicator3Format);
        else:
            gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['수익률지표']));
    nColOffset = nColOffset + 1;

    nLength = len(stStockInfor['YearDataList']);
    for nYearIndex in range(nLength):
        stYearDataList = stStockInfor['YearDataList'][nYearIndex];
        if (stYearDataList["item_value"] != u''):
            stDay = GetSplitTitle(stYearDataList["day"]);
            stThisYear = stDay.split('/')[0];        
            if ((stThisYear == u'지표') and (float(stYearDataList["item_value"]) >= 700)):
                if (float(stYearDataList["item_value"]) >= 730):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stYearDataList["item_value"]), stIndicator1Format);
                elif (float(stYearDataList["item_value"]) >= 710):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stYearDataList["item_value"]), stIndicator2Format);
                else:
                    gstFnSheet.write(nRowOffset, nColOffset, float(stYearDataList["item_value"]), stIndicator3Format);
            else:
                gstFnSheet.write(nRowOffset, nColOffset, float(stYearDataList["item_value"]));
        nColOffset = nColOffset + 1;

    nLength = len(stStockInfor['QuaterDataList']);
    for nQuaterIndex in range(nLength):
        stQuaterDataList = stStockInfor['QuaterDataList'][nQuaterIndex];
        if (stQuaterDataList["item_value"] != u''):
            stDay = GetSplitTitle(stQuaterDataList["day"]);
            stThisQuarter = stDay.split('/')[1];
            if ((stThisQuarter == u'지표') and (float(stQuaterDataList["item_value"]) >= 700)):
                if (float(stQuaterDataList["item_value"]) >= 730):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]), stIndicator1Format);
                elif (float(stQuaterDataList["item_value"]) >= 710):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]), stIndicator2Format);
                else:
                    gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]), stIndicator3Format);
            else:
                gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]));
        nColOffset = nColOffset + 1;

def SetKospiXlsxData(nColOffset, nType, astStockInfor, astBaseInfor):
    nRowOffset = 1;
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
    nLastDayIndex = 0;
    for nBaseIndex in range(nBaseLength):
        bFound = 0;
        for nDayIndex in range(nLastDayIndex, nStockLength):
            if (astBaseInfor[nBaseIndex]['Date'] == astStockInfor[nDayIndex]['Date']):
                bFound = 1;
                break;

        if (bFound == 1):
            stStockInfor = astStockInfor[nDayIndex];
            nCurPrice = stStockInfor['Price'];

            if (bFirstPrice == 0):
                bFirstPrice = 1;
            else:
                nCurRate = 0;
                if (float(nPrevPrice) > 0):
                    nCurRate = float((float(nCurPrice) * 100) / float(nPrevPrice)) - 100;

                gstSiseSheet.write(nRowOffset, nColOffset, nCurRate, stRateFormat);

            gstSiseSheet.write(nRowOffset, nColOffset + 1, nCurPrice, stSiseFormat);

            nPrevPrice = nCurPrice;
            nLastDayIndex = nLastDayIndex + 1;
        nRowOffset = nRowOffset + 1;

def SetSiseXlsxData(nColOffset, astKospiInfor, stStockInfor):
    nRowOffset = 1;
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
                nCurRate = 0;
                if (float(nPrevPrice) > 0):
                    nCurRate = float((float(nCurPrice) * 100) / float(nPrevPrice)) - 100;
                if (nCurRate > nImpossibleRate) or (nCurRate < (nImpossibleRate * -1)):
                    nCurRate = 0;

                gstSiseSheet.write(nRowOffset, nColOffset, nCurRate, stRateFormat);

            gstSiseSheet.write(nRowOffset, nColOffset + 1, float(nCurPrice));

            nPrevPrice = nCurPrice;
        nRowOffset = nRowOffset + 1;


# 누적승리 출력
def PrintWinningRate(nRowOffset, nColOffset, nTitle, nMaxDateCount):
    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstGraphSheet.write(nRowOffset, nColOffset, nTitle, stTitleBoldFormat);
    gstGraphSheet.write(nRowOffset + 1, nColOffset, u"누적 승리", stRedTitleBoldFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + (nRowOffset + 2);

        stAccumulatedCell = xl_rowcol_to_cell(nDateRowOffset - 1, nColOffset);
        if (nTitle == u"KOSPI"):
            stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 1);
        else:
            stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 2);
        stKospiRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 3);

        stString = "=IFERROR(" + stAccumulatedCell + " + (" + stAvgStockRate + " - " + stKospiRate + "), \"\")";

        gstGraphSheet.write(nDateRowOffset, nColOffset, stString, stRateFormat);

def SetGraphXlsxData(nMaxDateCount, nMaxStockCount):
    nStartGraphRowOffset = 3;
    nMaxRowOffset = nMaxDateCount + nStartGraphRowOffset;

    nStartFnRowOffset = 3;
    nEndFnRowOffset = nStartFnRowOffset + nMaxStockCount - 1;
    stStartFnRowOffset = str(nStartFnRowOffset);
    stEndFnRowOffset = str(nEndFnRowOffset);

    stSiseCell = gstSiseSheetName + u'!';
    nStockChoiceRowOffset = 0;
    nGraphRowOffset = 1;
    nRowOffset = 0;

    nKospiOffset = 2;
    nKosdaqOffset = 4;

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

        if (nRowOffset == nStockChoiceRowOffset):   # 0
            continue;
        elif (nRowOffset < nStartGraphRowOffset):      # 0 or 1
            gstGraphSheet.write(nRowOffset, nDateColOffset, stDateString, stPurpleBoldFormat);
        else:                                       # > 1
            gstGraphSheet.write(nRowOffset, nDateColOffset, stDateString, stPurpleFormat);

        stTransCell = xl_rowcol_to_cell(nRowOffset, nKospiColOffset);
        stKospiString = u'=' + stSiseCell + stTransCell;
        if (nRowOffset == (nStartGraphRowOffset - 2)):
            gstGraphSheet.write(nRowOffset, nKospiColOffset, stKospiString, stTitleFormat);
        elif (nRowOffset == (nStartGraphRowOffset - 1)):
            gstGraphSheet.write(nRowOffset, nKospiColOffset, stKospiString, stGreenTitleFormat);
        else:
            gstGraphSheet.write(nRowOffset, nKospiColOffset, stKospiString, stRateFormat);

        stTransCell = xl_rowcol_to_cell(nRowOffset, nKosdaqColOffset + 1);
        stKospiString = u'=' + stSiseCell + stTransCell;
        if (nRowOffset == (nStartGraphRowOffset - 2)):
            gstGraphSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stTitleFormat);
        elif (nRowOffset == (nStartGraphRowOffset - 1)):
            gstGraphSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stGreenTitleFormat);
        else:
            gstGraphSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stRateFormat);

    # 평균 증감
    gstGraphSheet.write(nGraphRowOffset, nAvgStockColOffset, u"종목 평균", stNavyFormat);
    gstGraphSheet.write(nGraphRowOffset + 1, nAvgStockColOffset, u"증감율", stGreenTitleFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + nStartGraphRowOffset;
        stStartTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset);
        stEndTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nMaxStockCount - 1);
        stString = "=IFERROR(AVERAGE(" + stStartTransCell + ":" + stEndTransCell + "), \"\")";
        gstGraphSheet.write(nDateRowOffset, nAvgStockColOffset, stString, stRateFormat);


    # KOSPI 누적승리
    PrintWinningRate(nGraphRowOffset, nKospiVsColOffset, u"KOSPI", nMaxDateCount);
    PrintWinningRate(nGraphRowOffset, nKosdaqVsColOffset, u"KOSDAQ", nMaxDateCount);

    # 선정 종목 (그래프 취합 50개 제한)
    nStockCount = nMaxStockCount;
    if (nStockCount > gnMaxGraphStockCount):
        nStockCount = gnMaxGraphStockCount;
    for nStockIndex in range(nStockCount):
        for nRowOffset in range(nMaxRowOffset):
            stStockColOffset = str(nStockIndex + 1);
            stSiseRowOffset = str(nRowOffset + 1);

            stTransCell = xl_rowcol_to_cell(0, nStockColOffset + nStockIndex);
            stString = "=IFERROR(";
            stString += "INDIRECT(ADDRESS(" + stSiseRowOffset + ", " + stTransCell + ", ";
#            stString += "INDIRECT(ADDRESS(" + stSiseRowOffset + ", INDIRECT(ADDRESS(2 + MATCH(" + stStockColOffset + ", ";
#            stString += gstFnSheetName + "!$A$" + stStartFnRowOffset + ":$A$" + stEndFnRowOffset + ", 0), 5, 4, 5, \"" + gstFnSheetName + "\")), ";
            stString += "4, 5, \"" + gstSiseSheetName + "\"))";
            stString += ", \"\")";

            # 일반 선정 종목 증감율 값
            if (nRowOffset >= nStartGraphRowOffset):
                gstGraphSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString, stRateFormat);
            # 선정 종목 Title
            elif (nRowOffset >= nGraphRowOffset):
                gstGraphSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);
            # 선정 종목 매핑 정보
            else:
                stString = "=IFERROR(";
                stString += "INDIRECT(ADDRESS(2 + MATCH(" + stStockColOffset + ", ";
                stString += gstFnSheetName + "!$A$" + stStartFnRowOffset + ":$A$" + stEndFnRowOffset + ", 0), 5, 4, 5, \"" + gstFnSheetName + "\"))";
                stString += ", \"\")";
                gstGraphSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

    # 차트 출력
    # 누적 승리율
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartGraphRowOffset, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKospiVsColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKospiVsColOffset);
    stKospiData = '=' + gstGraphSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKosdaqVsColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKosdaqVsColOffset);
    stKosdaqData = '=' + gstGraphSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartDateCell = xl_rowcol_to_cell(nStartGraphRowOffset, nDateColOffset);
    stEndDateCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nDateColOffset);
    stDate = '=' + gstGraphSheetName + '!' + stStartDateCell + ":" + stEndDateCell;

    stTitle = xl_rowcol_to_cell(1, nKospiVsColOffset);
    stChart.set_title({'name':u"KOSPI / KOSDAQ 대비 누적 승리율"});
    stChart.set_x_axis({'name':u'날짜'});
    stChart.set_y_axis({'name':u'승리율(%)', 'min':0, 'max':100, 'major_unit':10});

    stChart.add_series({'name':u"KOSPI",  'categories':stDate, 'text_axis':True, 'values':stKospiData});
    stChart.add_series({'name':u"KOSDAQ", 'categories':stDate, 'text_axis':True, 'values':stKosdaqData});

    stChart.set_size({'width':720, 'height':504});
    gstGraphSheet.insert_chart(stGraphCell, stChart);


    # KOSPI / KOSDAQ 지수
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartGraphRowOffset + 25, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKospiOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKospiOffset);
    stKospiSise = '=' + gstSiseSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKosdaqOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKosdaqOffset);
    stKosdaqSise = '=' + gstSiseSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stTitle = xl_rowcol_to_cell(1, nKospiVsColOffset);
    stChart.set_title({'name':u"KOSPI / KOSDAQ 지수"});
    stChart.set_x_axis({'name':u'날짜'});
    stChart.set_y_axis({'name':u'KOSPI지수',   'num_format':'0'});
    stChart.set_y2_axis({'name':u'KOSDAQ지수', 'num_format':'0'});

    stChart.add_series({'name':u"KOSPI",  'categories':stDate, 'values':stKospiSise});
    stChart.add_series({'name':u"KOSDAQ", 'categories':stDate, 'values':stKosdaqSise, 'y2_axis':1});

    stChart.set_size({'width':770, 'height':504});
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
    stAutoFilter = SetFnXlsxTitle(astStockInfor);
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
    return stAutoFilter;

def SISE_GetNonStockInfor(nStockCode, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    anUrl = "http://vip.mk.co.kr/newSt/rate/kospikosdaq_2.php?sty=2010&stm=5&std=1";
    stDataInfor = {};

    anReqCode               = {};
    anReqCode['KOSPI']      = "&stCode=KPS001";
    anReqCode['KOSDAQ']     = "&stCode=KDS001";

    anUrl                   = anUrl + "&eny=" + str(ganYear[0]) + "&enm=" + str(ganMonth[0]) + "&end=" + str(ganDay[0]);
    anUrl                   = anUrl + anReqCode[nStockCode];
    stResponse              = GetUrlOpen(anUrl);
    stPage                  = stResponse.read();
    stSoup                  = BeautifulSoup(stPage);
    astTables               = stSoup.findAll('table')[7].contents;

    nTableLen = len(astTables);
    for nIndex in range(5, nTableLen, 2):
        stStock             = {};

        astSplit = astTables[nIndex].text.split(u'\n');

        stStock['Date']     = astSplit[1].replace(".", "-");
        stStock['Price']    = float(astSplit[2].replace(",", ""));
        stStockInfor.append(stStock);

def SISE_GetStockInfor(nStockCode, nStockType, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    nCodeUrl = 'http://www.etomato.com/home/itemAnalysis/ItemPrice.aspx?item_code=';
    nCodeUrl = nCodeUrl + nStockCode;
    nResponse = GetUrlOpen(nCodeUrl);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);
    tables = nSoup.findAll('table');

    # 해당 Page로부터 시세 정보 확인 불가
    if ((len(tables) <= 13) or (len(tables[13].text.split(u'창출')) <= 1)):
        return 0;

    astTable = tables[20].contents;
    nTableLen = len(astTable);
    for nIndex in range(3, nTableLen):
        nLen = len(astTable[nIndex]);
        if (nLen <= 1):
            continue;

        astSplit = astTable[nIndex].text.split(u'\n');
        stDate = astSplit[1][2:].replace(".", "-");
        stStockInfor[stDate] = int(astSplit[2].replace(",", ""));
        if (stDate == u"10-04-30"):
            break;
    return 1;

gastKospiInfor      = [];
gastKosdaqInfor     = [];
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

gstDate = GetTodayString(ganYear, ganMonth, ganDay);

# Kospi / Kosdaq 정보 취합
SISE_GetKospiInfor(gastKospiInfor, gastKosdaqInfor);

# 종목 정보 취합
if (gnMaxKospiStockCount > 0):
    COMPANY_GetStockName(u'KOSPI', gastKospiStockName, gnGetBaeDangStockCount);
if (gnMaxKosdaqStockCount > 0):
    COMPANY_GetStockName(u'KOSDAQ', gastKosdaqStockName, gnGetBaeDangStockCount);
COMPANY_GetStockCode(gastChangeStockNameCodeList);
if (gnMaxKospiStockCount > 0):
    COMPANY_GetNameToCode(u'KOSPI', gastChangeStockNameCodeList, gastKospiStockName, gastStockNameCodeInfor);
if (gnMaxKosdaqStockCount > 0):
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

gstAutoFilterEndCell = COMPANY_WriteExcelFile(gastKospiInfor, gastKosdaqInfor, gastStockInfor);

gstFnSheet.autofilter(gstAutoFilterStartCell + ':' + gstAutoFilterEndCell);
gstFnSheet.freeze_panes('D3');
gstFnSheet.set_column('A:A', None, None, {'hidden': 1});
gstSiseSheet.freeze_panes('F4');
#gstSiseSheet.set_row(0, None, None, {'hidden': True})
gstGraphSheet.freeze_panes('G4');
#gstGraphSheet.set_row(0, None, None, {'hidden': True})

PrintProgress(u"[시작] 엑셀 출력");
gstWorkBook.close();
PrintProgress(u"[완료] 엑셀 출력");
PrintProgress(u"Complete all process");

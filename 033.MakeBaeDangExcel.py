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
gbPrintProgress = 1;


gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

def GetTodayString():
    stNow = datetime.datetime.now();

    stDate = str(stNow.year)[2:];
    if (stNow.month < 10):
        stDate = stDate + '0';
    stDate = stDate + str(stNow.month);
    if (stNow.day < 10):
        stDate = stDate + '0';
    stDate = stDate + str(stNow.day);

    return stDate;

# Stock �̸��� �ڵ带 ��� �Լ�
gastStockList = {};
def COMPANY_GetStockCode(astStockList): # OUT (gastStockList: ���� �̸� / �ڵ�)
    PrintProgress(u"[����] ���� �ڵ� ����");
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
    PrintProgress(u"[�Ϸ�] ���� �ڵ� ����");


gastStockNameCode = [];
def COMPANY_GetNameToCode(astStockList, astStockName, astStockNameCode):   # IN (nStock: �����ڵ�), OUT (stStockInfor: ���� ����)
    PrintProgress(u"[����] ���� �ڵ� ��ȯ");
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
    PrintProgress(u"[�Ϸ�] ���� �ڵ� ��ȯ");

gastStockName = [];
def COMPANY_GetStockName(astStockName, nMaxStockCount):
    PrintProgress(u"[����] ���� ����Ʈ ����");
    nMaxPageRange = 10;

    for nPageIndex in range(nMaxPageRange):
        if (len(astStockName) >= nMaxStockCount):
            break;

        anUrl = 'http://finance.naver.com/sise/dividend_list.nhn?sosok=KOSPI&fsq=20144&field=divd_rt&ordering=desc&page=' + str(nPageIndex + 1);
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
            astStockName.append(nStockName);
            PrintProgress(u"[����] ���� ����Ʈ ����: " + nStockName);

            if ((len(astStockName) % 50) == 0):
                break;
            if (len(astStockName) >= nMaxStockCount):
                break;

    PrintProgress(u"[�Ϸ�] ���� ����Ʈ ����");


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
    stString = stString.split(u"(IFRS����)")[0];
    stString = stString.split(u"(IFRS����)")[0];
    return stString;

def COMPANY_GetStockFinanceInfor(nName, nCode, astStockInfor):
    stStockInfor = {};
    nCodeUrl = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=';
    nCodeUrl = nCodeUrl + nCode;
    nResponse = gnOpener.open(nCodeUrl, timeout=60);
    nPage = nResponse.read();
    nSoup = BeautifulSoup(nPage);
    tables = nSoup.findAll('table');

    astTable4 = tables[4].text.split('/');
    astTable42 = astTable4[2].split(u'\n');
    stStockInfor['CurPrice'] = astTable42[1].replace(',', '');
    stStockInfor['CurPrice'] = stStockInfor['CurPrice'].replace(' ', '');

    astSplit = tables[1].text.split(' | ');

    stSplit0 = astSplit[0].split(' ');
    stStockInfor['Name'] = nName;
    stStockInfor['Code'] = nCode;
    nSplit0Len = len(stSplit0);
    stStockInfor['WebCode'] = stSplit0[nSplit0Len - 1];

#    stSplit2 = astSplit[2].split(' : ');
#    stStockInfor['����Type'] = stSplit2[0];

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
    stSplit3511 = stSplit351[1].split(u'����');
    stSplit35110 = stSplit3511[0].replace('%', '');
    stStockInfor['����'] = stSplit35110.replace(',', '');
    stStockInfor['����'] = stStockInfor['����'].replace(' ', '');

    astYearDataList = [];
    astQuaterDataList = [];
    COMPANY_GetFinance(stStockInfor['WebCode'], astYearDataList, astQuaterDataList);
    stStockInfor['YearDataList'] = astYearDataList;
    stStockInfor['QuaterDataList'] = astQuaterDataList;

    stStockInfor['�ü�'] = {};
    SISE_GetStockInfor(nCode, stStockInfor['�ü�']);

    astStockInfor.append(stStockInfor);

gastStockInfor = [];
def COMPANY_GetFinanceInfor(astStockNameCode, astStockInfor):
    PrintProgress(u"[����] ���� ���� ����: " + str(0) + " / " + str(nStockLen));
    nStockLen = len(astStockNameCode);
    for nStockIndex in range(nStockLen):
        COMPANY_GetStockFinanceInfor(astStockNameCode[nStockIndex]['Name'],
                                        astStockNameCode[nStockIndex]['Code'],
                                        astStockInfor);
        PrintProgress(u"[����] ���� ���� ����: " + str(nStockIndex + 1) + " / " + str(nStockLen) + " - " + astStockNameCode[nStockIndex]['Name']);
    PrintProgress(u"[�Ϸ�] ���� ���� ����: " + str(nStockLen) + " / " + str(nStockLen));

def SetFnXlsxTitle(astStockInfor):
    stStockInfor = astStockInfor[0];
    nStockLen = len(astStockInfor);
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
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stChoiceFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'black'});

    gstFnSheet.write(0, nColOffset, u"������", stChoiceFormat);
    stStartCell = xl_rowcol_to_cell(2, nColOffset);
    stEndCell = xl_rowcol_to_cell(2 + nStockLen - 1, nColOffset);
    gstFnSheet.write(1, nColOffset, "=count(" + stStartCell + ":" + stEndCell + ")", stChoiceFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(0, nColOffset, u"�����", stPurpleFormat);
    nColOffset = nColOffset + 1;

#    gstFnSheet.write(0, nColOffset, u"����Type", stNavyFormat);
#    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"�ڵ��ȣ", stGrayFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"�ü�����", stGrayFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"WICS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"���簡��", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"PER", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"PBR", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"BPS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"EPS", stNavyFormat);
    nColOffset = nColOffset + 1;
    gstFnSheet.write(0, nColOffset, u"����", stNavyFormat);
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
            gstFnSheet.write(0, nColOffset, u"���� " + stItemName, stRedTitleFormat);

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
            gstFnSheet.write(0, nColOffset, u"�б� " + stItemName, stGreenTitleFormat);

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

    gstSiseSheet.write(0, nColOffset, u'��¥', stPurpleBoldFormat);
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

    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'font_color': 'gray', 'underline':  1});

    # �����
    gstFnSheet.write(nRowOffset, nColOffset, stStockInfor['Name'], stPurpleFormat);
    nColOffset = nColOffset + 1;

    # �ڵ��ȣ
    stCell = xl_rowcol_to_cell(nRowOffset, nColOffset);
    gstFnSheet.write_url(stCell, nCodeUrl + stStockInfor['Code'], stGrayFormat, stStockInfor['Code']);
    nColOffset = nColOffset + 1;

    # �ü�����
    stLinkCell = xl_rowcol_to_cell(0, 3 + (2 * nStockIndex));
    stTargetColOffset = str(4 + (2 * nStockIndex));
    gstFnSheet.write(nRowOffset, nColOffset, nSiseUrl + stLinkCell, stGrayFormat, stTargetColOffset);
    nColOffset = nColOffset + 1;

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
    if (stStockInfor['����'] != u''):
        gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['����']));
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

def SetKospiXlsxData(nColOffset, astKospiInfor):
    nRowOffset = 0;
    bFirstPrice = 0;
    nCurPrice = 0;
    nCurRate = 0;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stSiseFormat = gstWorkBook.add_format({'num_format':'0.00'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstSiseSheet.write(nRowOffset, nColOffset, u'KOSPI', stTitleFormat);
    nRowOffset = nRowOffset + 1;

    gstSiseSheet.write(nRowOffset, nColOffset, u'������', stRedTitleFormat);
    gstSiseSheet.write(nRowOffset, nColOffset + 1, u'�ü�', stGreenTitleFormat);
    nRowOffset = nRowOffset + 1;

    # �ü� ���
    nKospiLength = len(astKospiInfor);
    for nDayIndex in range(nKospiLength):
        stStockInfor = astKospiInfor[nDayIndex];
        nCurPrice = stStockInfor['Price'];

        if (bFirstPrice == 0):
            bFirstPrice = 1;
        else:
            nCurRate = ((nCurPrice * 100) / nPrevPrice) - 100;
            gstSiseSheet.write(nRowOffset, nColOffset, nCurRate, stRateFormat);

        gstSiseSheet.write(nRowOffset, nColOffset + 1, nCurPrice, stSiseFormat);

        nPrevPrice = nCurPrice;
        nRowOffset = nRowOffset + 1;

def SetSiseXlsxData(nColOffset, astKospiInfor, stStockInfor):
    nRowOffset = 0;
    nKospiIndex = 0;
    astSiseStockInfor = stStockInfor['�ü�'];
    bFirstPrice = 0;
    nCurPrice = 0;
    nCurRate = 0;
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

    gstSiseSheet.write(nRowOffset, nColOffset, u'������', stRedTitleFormat);
    gstSiseSheet.write(nRowOffset, nColOffset + 1, u'�ü�', stGreenTitleFormat);
    nRowOffset = nRowOffset + 1;

    # �ü� ���
    nKospiLength = len(astKospiInfor);
    for nKospiIndex in range(nKospiLength):
        if (astSiseStockInfor.has_key(astKospiInfor[nKospiIndex]['Date'])):
            nCurPrice = astSiseStockInfor[astKospiInfor[nKospiIndex]['Date']];

            if (bFirstPrice == 0):
                bFirstPrice = 1;
            else:
                nCurRate = ((nCurPrice * 100) / nPrevPrice) - 100;
                if (nCurRate > nImpossibleRate) or (nCurRate < (nImpossibleRate * -1)):
                    nCurRate = 0;

                gstSiseSheet.write(nRowOffset, nColOffset, nCurRate, stRateFormat);

            gstSiseSheet.write(nRowOffset, nColOffset + 1, nCurPrice);

            nPrevPrice = nCurPrice;
        nRowOffset = nRowOffset + 1;

def SetGraphXlsxData(nMaxDateCount, nMaxsStockCount):
    nMaxRowOffset = nMaxDateCount + 2;
    nStartFnRowOffset = 3;
    nEndFnRowOffset = nStartFnRowOffset + nMaxsStockCount - 1;
    stStartFnRowOffset = str(nStartFnRowOffset);
    stEndFnRowOffset = str(nEndFnRowOffset);

    stSiseCell = gstSiseSheetName + u'!';
    nRowOffset = 0;
    nColOffset = 0;

    nDateColOffset = 0;
    nKospiColOffset = 1;
    nAvgStockColOffset = 2;
    nKospiVsColOffset = 3;
    nStockColOffset = 4;

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

    # ��¥ / KOSPI
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


    # ��� ����
    gstGraphSheet.write(0, nAvgStockColOffset, u"���� ���", stTitleFormat);
    gstGraphSheet.write(1, nAvgStockColOffset, u"������", stGreenTitleFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + 2;
        stStartTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset);
        stEndTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nMaxsStockCount - 1);
        stString = "=IFERROR(AVERAGE(" + stStartTransCell + ":" + stEndTransCell + "), \"\")";
        gstGraphSheet.write(nDateRowOffset, nAvgStockColOffset, stString, stRateFormat);


    # �����¸�
    gstGraphSheet.write(0, nKospiVsColOffset, u"���� ���", stTitleBoldFormat);
    gstGraphSheet.write(1, nKospiVsColOffset, u"���� �¸�", stRedTitleBoldFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + 2;

        stAccumulatedCell = xl_rowcol_to_cell(nDateRowOffset - 1, nKospiVsColOffset);
        stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nKospiVsColOffset - 1);
        stKospiRate = xl_rowcol_to_cell(nDateRowOffset, nKospiVsColOffset - 2);

        stString = "=IFERROR(" + stAccumulatedCell + " + (" + stAvgStockRate + " - " + stKospiRate + "), \"\")";

        gstGraphSheet.write(nDateRowOffset, nKospiVsColOffset, stString, stRateFormat);


    # ���� ����
    for nStockIndex in range(nMaxsStockCount):
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

    # ���� ����
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartFnRowOffset - 1, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(2, nKospiVsColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKospiVsColOffset);
    stData = '=' + gstGraphSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartDateCell = xl_rowcol_to_cell(2, nDateColOffset);
    stEndDateCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nDateColOffset);
    stDate = '=' + gstGraphSheetName + '!' + stStartDateCell + ":" + stEndDateCell;

    stTitle = xl_rowcol_to_cell(1, nKospiVsColOffset);
    stChart.set_title({'name':u"KOSPI ��� ���� �¸���"});
    stChart.set_x_axis({'name':u'��¥'});
    stChart.set_y_axis({'name':u'�¸���(%)', 'min':0, 'max':100});

    stChart.add_series({'name':u"���� �¸�",'categories':stDate, 'values':stData});
    stChart.set_size({'width':720, 'height':504});
    gstGraphSheet.insert_chart(stGraphCell, stChart);

def COMPANY_WriteExcelFile(astKospiInfor, astStockInfor):
    PrintProgress(u"[����] ���� ����");
    nColOffset = 0;
    nRowOffset = 0;

    # �ü� Title ���
    SetSiseXlsxTitle(astKospiInfor);
    nColOffset = nColOffset + 1;
    SetKospiXlsxData(nColOffset, astKospiInfor);
    nColOffset = nColOffset + 2;

    # �ü� ������ ���
    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        SetSiseXlsxData(nColOffset, astKospiInfor, astStockInfor[nStockIndex]);
        nColOffset = nColOffset + 2;

    # �繫 Title ���
    SetFnXlsxTitle(astStockInfor);
    nRowOffset = nRowOffset + 2;

    # �繫 ������ ���
    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        SetFnXlsxData(nRowOffset, astStockInfor, nStockIndex);
        nRowOffset = nRowOffset + 1;

    # �׷��� ���
    SetGraphXlsxData(len(astKospiInfor), len(astStockInfor));
    PrintProgress(u"[�Ϸ�] ���� ����");

# Date & ������ ��� �Լ� (�ڽ��� / �ڽ��� / �Ϲ�����)
# �ڽ��� or �ڽ��� or �Ϲ� ���� ����
#gnStockCode             = 'KOSPI';      # '1997-07-01' ~
#gnStockCode             = 'KOSDAQ';     # '2013-03-04' ~
#gnStockCode             = '014530';     # '2000-0101' ~
gastStockInfor          = [];
def SISE_GetNonStockInfor(nStockCode, stStockInfor):   # IN (nStock: �����ڵ�), OUT (stStockInfor: ���� ����)
    stDataInfor = {};

    anReqCode               = {};
    anReqCode['KOSPI']      = '^KS11';
    anReqCode['KOSDAQ']     = '^KQ11';

    # nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=0&b=1&c=1900';

    # Month = a + 1 / Day = b / Year = c
    nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=11&b=30&c=2012';
    stRequest               = requests.get(nUrl);
    stDataInfor             = pd.read_csv(StringIO(stRequest.content), index_col='Date', parse_dates={'Date'});

    for nIndex in range(stDataInfor.shape[0]):
        stStock             = {};
        stStock['Date']     = stDataInfor.index[nIndex]._date_repr[2:]; # ��¥
        stStock['Price']    = stDataInfor.values[nIndex][3];            # ����: 'Close'
        stStockInfor.append(stStock);

def SISE_GetStockInfor(nStockCode, stStockInfor):   # IN (nStock: �����ڵ�), OUT (stStockInfor: ���� ����)
    stDataInfor = {};

#        stStartDate             = datetime.datetime(1900, 1, 1);
    stStartDate             = datetime.datetime(2012, 12, 30);
    stDataInfor             = web.DataReader(nStockCode + ".KS", "yahoo", stStartDate);

    for nIndex in range(stDataInfor.shape[0]):
        stStockInfor[stDataInfor.index[nIndex]._date_repr[2:]]   = stDataInfor.values[nIndex][3];

gastKospiInfor      = [];
def SISE_GetKospiInfor(astKospiInfor):
    PrintProgress(u"[����] KOSPI ���� ����");
    SISE_GetNonStockInfor('KOSPI', astKospiInfor);
    astKospiInfor.sort();
    PrintProgress(u"[�Ϸ�] KOSPI ���� ����");

def PrintProgress(stString):
    if (gbPrintProgress > 0):
        print stString;

############# main #############

gstDate = GetTodayString();

# Kospi ���� ����
SISE_GetKospiInfor(gastKospiInfor);

# ���� ���� ����
COMPANY_GetStockName(gastStockName, gnMaxBaeDangStockCount);
COMPANY_GetStockCode(gastStockList);
COMPANY_GetNameToCode(gastStockList, gastStockName, gastStockNameCode);
COMPANY_GetFinanceInfor(gastStockNameCode, gastStockInfor);

# ���� ���� ���
gstWorkBook = xlsxwriter.Workbook('BaeDangStockList.xlsx');
gstFnSheetName = u'FN' + gstDate;
gstFnSheet = gstWorkBook.add_worksheet(gstFnSheetName);
gstSiseSheetName = u'�ü�' + gstDate;
gstSiseSheet = gstWorkBook.add_worksheet(gstSiseSheetName);
gstGraphSheetName = u'�׷���' + gstDate;
gstGraphSheet = gstWorkBook.add_worksheet(gstGraphSheetName);

COMPANY_WriteExcelFile(gastKospiInfor, gastStockInfor);

gstFnSheet.autofilter('E2:JG2');
gstFnSheet.freeze_panes('C3');
gstSiseSheet.freeze_panes('D3');
gstGraphSheet.freeze_panes('E3');

PrintProgress(u"[����] ���� ���");
gstWorkBook.close();
PrintProgress(u"[�Ϸ�] ���� ���");
PrintProgress(u"Complete all process");

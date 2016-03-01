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
import re;

gnMaxBaeDangStockCount  = 10000;     # 종목 추출 개수 - 10000 절반씩 코스피(5000), 코스닥(5000)
gnMaxGraphStockCount    = 100;      # 엑셀에서 선택 가능한 종목 개수


gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

gnGetBaeDangStockCount = int(gnMaxBaeDangStockCount + (gnMaxBaeDangStockCount * 0.2) + 1);
if (gnGetBaeDangStockCount <= 50):
    gnGetBaeDangStockCount = 50;
gnGetMaxKospiStockCount = gnGetBaeDangStockCount / 2;
gnGetMaxKosdaqStockCount = gnGetBaeDangStockCount / 2;
gnMaxKospiStockCount = gnMaxBaeDangStockCount / 2;
gnMaxKosdaqStockCount = gnMaxBaeDangStockCount / 2;
if ((gnMaxBaeDangStockCount % 2) > 0):
    gnMaxKospiStockCount = gnMaxKospiStockCount + 1;
gbPrintProgress = 1;
gbGetLastQuarter = 1;
gbWinningSheet = 1;
gbBenefitSheet = 1;

ganYear = [1];
ganMonth = [1];
ganDay = [1];
def GetTodayString(anYear, anMonth, anDay):
    stNow = datetime.datetime.now();

    if (stNow.year > 2016):
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

def COMPANY_CheckExpectSum(stItemName):
    astExpectSumName = [u"매출액", u"영업이익", u"세전계속사업이익", u"당기순이익", u"당기순이익(지배)", u"당기순이익(비지배)", u"당기순이익(비지배)",
                        u"영업활동현금흐름", u"투자활동현금흐름", u"재무활동현금흐름", u"CAPEX", u"FCF", u"현금DPS(원)"];

    nLen = len(astExpectSumName);
    for nIndex in range(nLen):
        if (astExpectSumName[nIndex] == stItemName):
            return True;

    return False;

def COMPANY_SetJsonData(stSoup, eFreq_typ, astDataSet, stExpectDataSet):
    astDays = stSoup.findAll("th", {"class":re.compile("r03c0[1-5]")})
    astDaysExpect = stSoup.findAll("th", {"class":re.compile("r03c0")})

    # 재무정보 타이틀
    astItemNames = stSoup.findAll("th", {"class":"bg txt title "})
    # 재무정보 값 
    astItemValues = stSoup.findAll("td", {"class":"num line "})

    nItemLen = len(astItemNames);
    nDayLen = len(astDays);
    for nItemIndex in range(nItemLen):
        astIndicatorData = [];
        nMultipleIndicator = 1;
        nSumIndicator = 0;

        for nDayIndex in range(nDayLen):
            # 주요 Factor 추가
            data = {};
            data["day"] = astDays[nDayIndex].text.replace('\r', '').replace('\t', '').replace('\n', '').split(u"(")[0];
            data["item_name"] = astItemNames[nItemIndex].text.replace(u'\xa0', '');
            data["item_value"] = astItemValues[nItemIndex * nDayLen + nDayIndex].text.replace(',', '');
            astDataSet.append(data);

            # 지표 계산
            if (nDayIndex > 0):
                nIndicator = nMultipleIndicator;
                for stIndicatorData in astIndicatorData:
                    if ((data["item_value"] != '') and (stIndicatorData["item_value"] != '') and (float(data["item_value"]) > 0) and (float(data["item_value"]) >= float(stIndicatorData["item_value"]))):
                        nSumIndicator = nSumIndicator + nIndicator;
                    nIndicator = nIndicator * 2;
                if (nMultipleIndicator <= 100):
                    nMultipleIndicator = nMultipleIndicator * 10;
                else:
                    nMultipleIndicator = nMultipleIndicator * 100;
            astIndicatorData.append(data);

        if (eFreq_typ == 'Y'):
            data = {};
            nStartIndex = ((nDayLen + 2) * nItemIndex) + 2;
            nQuarterCount = 3;
            nSum = 0;
            nAvg = 0;

            if (gbGetLastQuarter):
                nQuarterCount = 4;
                nStartIndex = ((nDayLen + 2) * nItemIndex) + 1;

            # Next 연간 Factor 예측
            for nExIndex in range(nQuarterCount):
                nCurIndex = nStartIndex + nExIndex;
                if (stExpectDataSet[nCurIndex]['item_value'] != u''):
                    nSum = nSum + float(stExpectDataSet[nCurIndex]['item_value']);
            nAvg = nSum / nQuarterCount;

            # Next 연간 Factor 추가
            data["day"] = astDaysExpect[nDayLen + 1].text.replace('\r', '').replace('\t', '').replace('\n', '').split(u"(")[0];
            data["item_name"] = astItemNames[nItemIndex].text.replace(u'\xa0', '');
            if (COMPANY_CheckExpectSum(data["item_name"])):
                if (gbGetLastQuarter):
                    data["item_value"] = round(nSum, 2);
                else:
                    data["item_value"] = round(nSum + nAvg, 2);
            else:
                data["item_value"] = round(nAvg, 2);
            astDataSet.append(data);

            # Next 연간 지표 계산
            nIndicator = nMultipleIndicator;
            for stIndicatorData in astIndicatorData:
                if ((data["item_value"] != '') and (stIndicatorData["item_value"] != '') and (float(data["item_value"]) > 0) and (float(data["item_value"]) >= float(stIndicatorData["item_value"]))):
                    nSumIndicator = nSumIndicator + nIndicator;
                nIndicator = nIndicator * 2;

        # 보조 지표 추가 (예측 제외)
        stAppendData = {};
        stAppendData["day"] = u"지표/지표";
        stAppendData["item_name"] = astItemNames[nItemIndex].text;
        stAppendData["item_value"] = unicode(nSumIndicator % 100000);
        astDataSet.append(stAppendData);
        # 지표 추가
        stAppendData = {};
        stAppendData["day"] = u"지표/지표";
        stAppendData["item_name"] = astItemNames[nItemIndex].text;
        stAppendData["item_value"] = unicode(nSumIndicator);
        astDataSet.append(stAppendData);

def COMPANY_SetFinance(nCode, eFreq_typ, stDataSet, stExpectDataSet):
    anUrl = "http://companyinfo.stock.naver.com/v1/company/ajax/cF1001.aspx?fin_typ=0"
    strUrl = anUrl + "&freq_typ=" + eFreq_typ + "&cmp_cd=" + str(nCode)
    stResponse = GetUrlOpen(strUrl);
    stPage = stResponse.read();
    stSoup = BeautifulSoup(stPage);
    COMPANY_SetJsonData(stSoup, eFreq_typ, stDataSet, stExpectDataSet);

gastYearDataList = [];
gastQuaterDataList = [];
def COMPANY_GetFinance(ncode, stStockInfor):
    stQuaterData = [];
    COMPANY_SetFinance(ncode, "Q", stQuaterData, 0);
    stStockInfor['QuaterDataList'] = copy.deepcopy(stQuaterData);

    stYearData = [];
    COMPANY_SetFinance(ncode, "Y", stYearData, stQuaterData);
    stStockInfor['YearDataList'] = copy.deepcopy(stYearData);

def GetSplitTitle(stString):
    stString = stString.split(u"(IFRS연결)")[0];
    stString = stString.split(u"(IFRS별도)")[0];
    return stString;

def COMPANY_CheckBestStockInfor(stItemName):
    astBestItemName = [u"매출액", u"영업이익", u"영업이익률"];
    astBestDebtName = [u"부채비율"];
    astBestAllocName = [u"현금배당수익률"];
    
    nLen = len(astBestItemName);
    for nIndex in range(nLen):
        if (astBestItemName[nIndex] == stItemName):
            return 1;

    nLen = len(astBestDebtName);
    for nIndex in range(nLen):
        if (astBestDebtName[nIndex] == stItemName):
            return 2;

    nLen = len(astBestAllocName);
    for nIndex in range(nLen):
        if (astBestAllocName[nIndex] == stItemName):
            return 3;

    return 0;

def COMPANY_SetBestStockInfor(stStockInfor):
    nYearFieldCount = 8;
    nQuaterFieldCount = 7;
    nAttrCount = len(stStockInfor['YearDataList']) / nYearFieldCount;
    nYearIndicatorOffset = nYearFieldCount - 1;
    nQuaterIndicatorOffset = nQuaterFieldCount - 1;
    nYearDebtOffset = nYearFieldCount - 4;
    nQuaterDebtOffset = nQuaterFieldCount - 3;
    nYearAllocOffset = nYearFieldCount - 4;
    nQuaterAllocOffset = nQuaterFieldCount - 4;
    nYearBestIndicator = 3115000;
    nQuaterBestIndicator = 15000;
    nBestDebt = 150;

    stStockInfor['BestStock'] = 0;
    if (stStockInfor['Type'] != u'KOSPI'):
        return;

    for nIndex in range(nAttrCount):
        nCurOffset = (nIndex * nQuaterFieldCount) + nQuaterIndicatorOffset;
        nBestStockType = COMPANY_CheckBestStockInfor(stStockInfor['QuaterDataList'][nCurOffset]["item_name"]);
        if (nBestStockType == 1):
            if (float(stStockInfor['QuaterDataList'][nCurOffset]["item_value"]) <= nQuaterBestIndicator):
                return;
            continue;

        nCurOffset = (nIndex * nQuaterFieldCount) + nQuaterDebtOffset;
        nBestStockType = COMPANY_CheckBestStockInfor(stStockInfor['QuaterDataList'][nCurOffset]["item_name"]);
        if (nBestStockType == 2):
            if (float(stStockInfor['QuaterDataList'][nCurOffset]["item_value"]) >= nBestDebt):
                return;
            continue;

        nCurOffset = (nIndex * nYearFieldCount) + nYearAllocOffset;
        nBestStockType = COMPANY_CheckBestStockInfor(stStockInfor['YearDataList'][nCurOffset]["item_name"]);
        if (nBestStockType == 3):
            if (float(stStockInfor['YearDataList'][nCurOffset]["item_value"]) <= 0):
                return;
            continue;

    stStockInfor['BestStock'] = 1;

def COMPANY_SetStockInfor(stStockInfor, tables, nType, nName, nCode):
    nIndexStock = 0;
    nIndexStockSise = 1;
    nIndexPrice = 2;
    nIndexCode = 8;
    nIndexWICS = 12;
    nIndexEPS = 24;
    nIndexBPS = 25;

    nIndexPER = 26;
    nIndexPBR = 28;
    nIndexBaeDang = 29;

    astTable4 = tables[nIndexStockSise].text.split('/');
    astTable42 = astTable4[2].replace('\r', '').replace('\t', '');
    astTable42 = astTable42.split(u'\n');
    stStockInfor['CurPrice'] = astTable42[nIndexPrice].replace(',', '');
    stStockInfor['CurPrice'] = stStockInfor['CurPrice'].replace(u'원', '');

    stStockInfor['Name'] = nName;
    stStockInfor['Code'] = nCode;
    stStockInfor['Type'] = nType;
    astSplit = tables[nIndexStock].text.split('\n');

    stStockInfor['WebCode'] = astSplit[nIndexCode];
    if (stStockInfor['WebCode'] != nCode):
        return 0;

    stSplitWICS = astSplit[nIndexWICS].split(' : ');
    stStockInfor['WICS'] = stSplitWICS[1];

    stSplitEPS = astSplit[nIndexEPS].split(' ');
    stStockInfor['EPS'] = stSplitEPS[1].replace(',', '');

    stSplitBPS = astSplit[nIndexBPS].split(' ');
    stStockInfor['BPS'] = stSplitBPS[1].replace(',', '');

    stSplitPER = astSplit[nIndexPER].split(' ');
    stStockInfor['PER'] = stSplitPER[1].replace(',', '');

    stSplitPBR = astSplit[nIndexPBR].split(' ');
    stStockInfor['PBR'] = stSplitPBR[1].replace(',', '');

    stSplitBaeDang = astSplit[nIndexBaeDang].split(' ');
    stStockInfor['배당률'] = stSplitBaeDang[1].replace('%', '');

    astTable4Len = len(astTable4);
    astTable4_1M = astTable4[astTable4Len - 4].split(u'\n');
    astTable4_3M = astTable4[astTable4Len - 3].split(u'\n');
    astTable4_6M = astTable4[astTable4Len - 2].split(u'\n');
    astTable4_1Y = astTable4[astTable4Len - 1].split(u'\n');

    stStockInfor['1M'] = astTable4_1M[len(astTable4_1M) - 1].replace('\r', '').replace(' ', '').replace(',', '').replace('%', '').replace('\t', '');
    stStockInfor['3M'] = astTable4_3M[len(astTable4_3M) - 1].replace('\r', '').replace(' ', '').replace(',', '').replace('%', '').replace('\t', '');
    stStockInfor['6M'] = astTable4_6M[len(astTable4_6M) - 1].replace('\r', '').replace(' ', '').replace(',', '').replace('%', '').replace('\t', '');
    stStockInfor['1Y'] = astTable4_1Y[0].replace('\r', '').replace(' ', '').replace(',', '').replace(',', '').replace('%', '').replace('\t', '');
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

    COMPANY_GetFinance(stStockInfor['WebCode'], stStockInfor);

    stStockInfor['시세'] = {};
    bRet = SISE_GetStockInfor(nCode, nType, stStockInfor['시세']);

    if (bRet > 0):
        COMPANY_SetBestStockInfor(stStockInfor);

    return bRet;

def COMPANY_GetStockFinanceInfor(nType, nName, nCode, astStockInfor):
    stStockInfor = {};
    nCodeUrl = 'http://companyinfo.stock.naver.com/v1/company/c1010001.aspx?cmp_cd=';
    nCodeUrl = nCodeUrl + nCode;
    stResponse = GetUrlOpen(nCodeUrl);
    stPage = stResponse.read();
    stSoup = BeautifulSoup(stPage);
    tables = stSoup.findAll('table');
    nMinTableSize = 16;

    if (len(tables) < nMinTableSize):
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

        PrintProgress(u"[진행] " + astStockNameCode[nStockIndex]['Type'] + u" 정보 취합: " + str(nKospiCount + nKosdaqCount) + " / " + str(nMaxGettingCount) + " - " + "<" + astStockNameCode[nStockIndex]['Code'] + "> " + astStockNameCode[nStockIndex]['Name']);

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

    PrintProgress(u"[완료] 종목 정보 취합: " + str(nMaxGettingCount) + " / " + str(nMaxGettingCount));

def CheckHighlightField(stItemName):
    astHighlightItemName = [u"영업이익", u"영업이익률", u"부채비율", u"PER(배)", u"PBR(배)", u"현금배당수익률"];

    nLen = len(astHighlightItemName);
    for nIndex in range(nLen):
        if (astHighlightItemName[nIndex] == stItemName):
            return True;

    return False;

gstAutoFilterStartCell  = 'A2';
gstAutoFilterEndCell    = 'A2';
def EXCEL_SetFnXlsxTitle(astStockInfor):
    stStockInfor = astStockInfor[0];
    nStockLen = len(astStockInfor);
    nRowOffset = 0;
    nColOffset = 0;
    nXlsxYear = 0;
    nXlsxQuarter= 0;

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue', 'align':'center'});
    stIndicatorFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'brown', 'align':'center'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPinkTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'pink'});
    stPurpleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stChoiceFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'black', 'align':'center'});

    gstFnSheet.write(nRowOffset, nColOffset, u"종목매핑", stPurpleFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(nRowOffset, nColOffset, u"이름매칭", stChoiceFormat);
    stStartCell = xl_rowcol_to_cell(2, nColOffset);
    stEndCell = xl_rowcol_to_cell(2 + nStockLen - 1, nColOffset);
    gstFnSheet.write(nRowOffset + 1, nColOffset, "=count(" + stStartCell + ":" + stEndCell + ")", stChoiceFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(nRowOffset, nColOffset, u"종목선정", stChoiceFormat);
    stStartCell = xl_rowcol_to_cell(2, nColOffset);
    stEndCell = xl_rowcol_to_cell(2 + nStockLen - 1, nColOffset);
    gstFnSheet.write(nRowOffset + 1, nColOffset, "=count(" + stStartCell + ":" + stEndCell + ")", stChoiceFormat);
    nColOffset = nColOffset + 1;

    gstFnSheet.write(nRowOffset, nColOffset, u"종목명", stPurpleFormat);
    nColOffset = nColOffset + 1;

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
            nXlsxYear = stDay;
        if (nXlsxYear == stDay):
            if (CheckHighlightField(stItemName)):
                gstFnSheet.write(nRowOffset, nColOffset, u"연간 " + stItemName, stPinkTitleFormat);
            else:
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
            nXlsxQuarter = stDay;
        if (nXlsxQuarter == stDay):
            if (CheckHighlightField(stItemName)):
                gstFnSheet.write(nRowOffset, nColOffset, u"분기 " + stItemName, stPinkTitleFormat);
            else:
                gstFnSheet.write(nRowOffset, nColOffset, u"분기 " + stItemName, stGreenTitleFormat);

        if (stThisQuarter == u'지표'):
            gstFnSheet.write(nRowOffset + 1, nColOffset, stThisQuarter, stIndicatorFormat);
        else:
            gstFnSheet.write(nRowOffset + 1, nColOffset, stThisQuarter, stTitleFormat);
        nColOffset = nColOffset + 1;

    stAutoFilterCell = xl_rowcol_to_cell(1, nColOffset - 1);
    return stAutoFilterCell;

def EXCEL_SetSiseXlsxTitle(astStockInfor):
    nRowOffset = 1;
    nColOffset = 0;

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

def EXCEL_SetFnXlsxMapping(nRowOffset, nColOffset):
    nStartRow = 3;
    stNameChoiceLocation = u'B';
    stStockChoiceLocation = u'C';
    nTargetRowOffset = nRowOffset + 1;

    stString = u'';
    stString = stString + u'=IF(ISNUMBER(' + stNameChoiceLocation + str(nTargetRowOffset) + u'),';
    stString = stString + u'COUNT(' + stNameChoiceLocation + str(nStartRow) + u':' + stStockChoiceLocation + str(nTargetRowOffset) + u'), ';
    stString = stString + u'IF(ISNUMBER(' + stStockChoiceLocation + str(nTargetRowOffset) + u'),';
    stString = stString + u'COUNT(' + stNameChoiceLocation + str(nStartRow) + u':' + stStockChoiceLocation + str(nTargetRowOffset) + u'), ';
    stString = stString + u'\"\"))';
    gstFnSheet.write(nRowOffset, nColOffset, stString);

def EXCEL_SetFnNameMapping(nRowOffset, nColOffset):
    nStartRow = 3;
    nEndRow = gnMaxBaeDangStockCount + nStartRow - 1;
    stStockChoiceLocation = u'C';
    stStockNameLocation = u'D';
    nTargetRowOffset = nRowOffset + 1;

    stString = u'';
    stString = stString + u'=IFERROR(MATCH(' + stStockNameLocation + str(nTargetRowOffset) + u',' + stStockChoiceLocation + str(nStartRow) + u':' + stStockChoiceLocation + str(nEndRow) + u',0), \"\")';
    gstFnSheet.write(nRowOffset, nColOffset, stString);

def EXCEL_SetFnXlsxData(nRowOffset, astStockInfor, nStockIndex):
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
    EXCEL_SetFnXlsxMapping(nRowOffset, nColOffset);
    nColOffset = nColOffset + 1;

    # 이름매핑
    EXCEL_SetFnNameMapping(nRowOffset, nColOffset);
    nColOffset = nColOffset + 1;

    # 종목선정
    if (stStockInfor['BestStock'] > 0):
        gstFnSheet.write(nRowOffset, nColOffset, float(0));
    nColOffset = nColOffset + 1;

    # 종목명
    if (stStockInfor['Type'] == 'KOSPI'):
        gstFnSheet.write(nRowOffset, nColOffset, stStockInfor['Name'], stPurpleFormat);
    else:
        gstFnSheet.write(nRowOffset, nColOffset, stStockInfor['Name'], stOrangeFormat);
    nColOffset = nColOffset + 1;

    # 코드번호
    stCell = xl_rowcol_to_cell(nRowOffset, nColOffset);
    gstFnSheet.write(stCell, nCodeUrl + stStockInfor['Code'], stGrayFormat, stStockInfor['Code']);
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

    nLevel1 = 3100000;
    nLevel2 = 3115000;
    nLevel3 = 3107000;

    if (stStockInfor['수익률지표'] != u''):
        if (float(stStockInfor['수익률지표']) >= nLevel1):
            if (float(stStockInfor['수익률지표']) >= nLevel2):
                gstFnSheet.write(nRowOffset, nColOffset, float(stStockInfor['수익률지표']), stIndicator1Format);
            elif (float(stStockInfor['수익률지표']) >= nLevel3):
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
            if ((stThisYear == u'지표') and (float(stYearDataList["item_value"]) >= nLevel1)):
                if (float(stYearDataList["item_value"]) >= nLevel2):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stYearDataList["item_value"]), stIndicator1Format);
                elif (float(stYearDataList["item_value"]) >= nLevel3):
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
            stThisQuarter = stDay.split('/')[0];
            if ((stThisQuarter == u'지표') and (float(stQuaterDataList["item_value"]) >= nLevel1)):
                if (float(stQuaterDataList["item_value"]) >= nLevel2):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]), stIndicator1Format);
                elif (float(stQuaterDataList["item_value"]) >= nLevel3):
                    gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]), stIndicator2Format);
                else:
                    gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]), stIndicator3Format);
            else:
                gstFnSheet.write(nRowOffset, nColOffset, float(stQuaterDataList["item_value"]));
        nColOffset = nColOffset + 1;

def EXCEL_SetKospiXlsxData(nColOffset, nType, astStockInfor, astBaseInfor):
    nRowOffset = 1;
    nStartRowOffset = nRowOffset;
    nChangeRowOffset = 2;
    bFirstPrice = 0;
    nCurPrice = 0;
    nCurRate = 0;
    nPrevPrice = 1;
    nInternalUrl = u'internal:';
    nBaseLength = len(astBaseInfor);
    nStockLength = len(astStockInfor);

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

    stTransCell = xl_rowcol_to_cell(nBaseLength + nChangeRowOffset, nColOffset);
    gstSiseSheet.write(nRowOffset, nColOffset, nInternalUrl + stTransCell, stRedTitleFormat, u'증감율');
    gstSiseSheet.write(nRowOffset, nColOffset + 1, u'시세', stGreenTitleFormat);
    nRowOffset = nRowOffset + 1;

    # 시세 출력
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
    stTransCell = xl_rowcol_to_cell(nStartRowOffset + nChangeRowOffset, nColOffset);
    gstSiseSheet.write(nRowOffset, nColOffset, nInternalUrl + stTransCell, stRedTitleFormat, u'위로');

def EXCEL_SetSiseXlsxData(nColOffset, astKospiInfor, stStockInfor):
    nRowOffset = 1;
    nStartRowOffset = nRowOffset;
    nChangeRowOffset = 2;
    nKospiIndex = 0;
    astSiseStockInfor = stStockInfor['시세'];
    bFirstPrice = 0;
    nCurPrice = 0;
    nCurRate = 0;
    nPrevPrice = 1;
    nImpossibleRate = 30;
    nInternalUrl = u'internal:';
    nFnUrl = nInternalUrl + gstFnSheetName + u'!';
    nKospiLength = len(astKospiInfor);

    stTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'green'});
    stPurpleFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    stTransCell = xl_rowcol_to_cell(int((nColOffset - 5) / 2) + 2, 3);
    gstSiseSheet.write(nRowOffset, nColOffset, nFnUrl + stTransCell, stNavyFormat, stStockInfor['Name']);
    nRowOffset = nRowOffset + 1;

    stTransCell = xl_rowcol_to_cell(nKospiLength + nChangeRowOffset, nColOffset);
    gstSiseSheet.write(nRowOffset, nColOffset, nInternalUrl + stTransCell, stRedTitleFormat, u'증감율');
    gstSiseSheet.write(nRowOffset, nColOffset + 1, u'시세', stGreenTitleFormat);
    nRowOffset = nRowOffset + 1;

    # 재무 Page 비정상 예외 처리
    if (stStockInfor['WebCode'] == "0"):
        return;

    # 시세 출력
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
    stTransCell = xl_rowcol_to_cell(nStartRowOffset + nChangeRowOffset, nColOffset);
    gstSiseSheet.write(nRowOffset, nColOffset, nInternalUrl + stTransCell, stRedTitleFormat, u'위로');

# 승리 출력
def EXCEL_PrintWinningDayRate(nRowOffset, nColOffset, nTitle, nMaxDateCount):
    nBaseColOffset = 1;
    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'magenta'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'brown'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstWinningSheet.write(nRowOffset, nColOffset, nTitle, stTitleBoldFormat);
    gstWinningSheet.write(nRowOffset + 1, nColOffset, u"오늘 승리", stRedTitleBoldFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + (nRowOffset + 2);

        if (nTitle == u"KOSPI"):
            stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - (nBaseColOffset + 0));
        else:
            stAvgStockRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - (nBaseColOffset + 1));
        stKospiRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - (nBaseColOffset + 2));

        stString = "=IFERROR(" + stAvgStockRate + " - " + stKospiRate + ", \"\")";

        gstWinningSheet.write(nDateRowOffset, nColOffset, stString, stRateFormat);

# 누적승리 출력
def EXCEL_PrintWinningSumRate(nRowOffset, nColOffset, nTitle, nMaxDateCount):
    nChoiceDateColOffset = 0;
    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    gstWinningSheet.write(nRowOffset, nColOffset, nTitle, stTitleBoldFormat);
    gstWinningSheet.write(nRowOffset + 1, nColOffset, u"누적 승리", stRedTitleBoldFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + (nRowOffset + 2);

        stDateChoiceCell = xl_rowcol_to_cell(nDateRowOffset, nChoiceDateColOffset);
        stAccumulatedCell = xl_rowcol_to_cell(nDateRowOffset - 1, nColOffset);
        stTodayRate = xl_rowcol_to_cell(nDateRowOffset, nColOffset - 2);

        stString = "=IF(" + stDateChoiceCell + " > 0, IFERROR(" + stAccumulatedCell + " + " + stTodayRate + ", 0), 0)";

        gstWinningSheet.write(nDateRowOffset, nColOffset, stString, stRateFormat);

def EXCEL_SetWinningRateGraphXlsxData(nMaxDateCount, nMaxStockCount):
    nStartGraphRowOffset = 3;
    nMaxRowOffset = nMaxDateCount + nStartGraphRowOffset;
    stMaxRowOffset = str(nMaxRowOffset);
    stChoiceDate = u'10-05-03';
    stDateString = 'B';
    stChoiceDateCell = 'B3';
    nInternalUrl = u'internal:';

    nDateRowOffset = 1;
    nChoiceDateRowOffset = 2;
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

    nChoiceDateColOffset = 0;
    nDateColOffset = 1;
    nKospiColOffset = 2;
    nKosdaqColOffset = 3;
    nAvgStockColOffset = 4;
    nKospiVsDayColOffset = 5;
    nKosdaqVsDayColOffset = 6;
    nKospiVsSumColOffset = 7;
    nKosdaqVsSumColOffset = 8;
    nStockColOffset = 9;
    nRoundUp = 3;

    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stChoiceDateFormat = gstWorkBook.add_format({'num_format': '@', 'font_color': 'blue'});
    stTitleFormat = gstWorkBook.add_format({'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'font_color': 'green'});
    stPurpleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});

    stBaseTransCell = xl_rowcol_to_cell(nGraphRowOffset + 1, nChoiceDateColOffset);
    
    # 선정 날짜 기준 현재 수익률 동작 여부
    stString = "=IFERROR(MATCH(" + stChoiceDateCell + ", " + stDateString + str(nStartFnRowOffset+1) + ":" + stDateString + str(stMaxRowOffset) + ", 0) + 3";
    stString += ", 0)";
    gstWinningSheet.write(nGraphRowOffset, nChoiceDateColOffset, u"선정 날짜", stGrayFormat);
    gstWinningSheet.write(nGraphRowOffset + 1, nChoiceDateColOffset, stString, stGrayFormat);
    for nDateIndex in range(nMaxDateCount):
        nDateRowCellOffset = nDateIndex + nStartGraphRowOffset;
        stCurDateTransCell = xl_rowcol_to_cell(nDateRowCellOffset, nDateColOffset);
        stChoiceDateTransCell = xl_rowcol_to_cell(nChoiceDateRowOffset, nDateColOffset);

        stString = "=IF(AND(";
        stString += stCurDateTransCell + " >= " + stChoiceDateTransCell;
        stString += ", " + stBaseTransCell + " > 0";
        stString += "), 1, 0)";

        gstWinningSheet.write(nDateRowCellOffset, nChoiceDateColOffset, stString);

    nDateChangeOffset = -1;
    nKospiChangeOffset = -1;
    nKosdaqChangeOffset = 0;
    # 날짜 / KOSPI
    for nRowOffset in range(nMaxRowOffset):
        stTransCell = xl_rowcol_to_cell(nRowOffset, nDateColOffset + nDateChangeOffset);
        stString = stSiseCell + stTransCell;
        stDateCellString = u'=' + "IF(" + stString + " > 0," + stString + ", \"\")";

        if (nRowOffset == nStockChoiceRowOffset):   # 0
            continue;
        elif (nRowOffset == nDateRowOffset):        # 1
            stTransCell = xl_rowcol_to_cell(nMaxRowOffset, nDateColOffset);
            gstWinningSheet.write(nRowOffset, nDateColOffset, nInternalUrl + stTransCell, stPurpleBoldFormat, u"날짜");
        elif (nRowOffset == nChoiceDateRowOffset):  # 2
            gstWinningSheet.write(nRowOffset, nDateColOffset, stChoiceDate, stChoiceDateFormat);
        else:                                       # > 2
            gstWinningSheet.write(nRowOffset, nDateColOffset, stDateCellString, stPurpleFormat);

        stTransCell = xl_rowcol_to_cell(nRowOffset, nKospiColOffset + nKospiChangeOffset);
        stKospiString = u'=' + stSiseCell + stTransCell;
        if (nRowOffset == (nStartGraphRowOffset - 2)):
            gstWinningSheet.write(nRowOffset, nKospiColOffset, stKospiString, stTitleFormat);
        elif (nRowOffset == (nStartGraphRowOffset - 1)):
            gstWinningSheet.write(nRowOffset, nKospiColOffset, stKospiString, stGreenTitleFormat);
        else:
            gstWinningSheet.write(nRowOffset, nKospiColOffset, stKospiString, stRateFormat);

        stTransCell = xl_rowcol_to_cell(nRowOffset, nKosdaqColOffset + nKosdaqChangeOffset);
        stKospiString = u'=' + stSiseCell + stTransCell;
        if (nRowOffset == (nStartGraphRowOffset - 2)):
            gstWinningSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stTitleFormat);
        elif (nRowOffset == (nStartGraphRowOffset - 1)):
            gstWinningSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stGreenTitleFormat);
        else:
            gstWinningSheet.write(nRowOffset, nKosdaqColOffset, stKospiString, stRateFormat);
    stTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nDateColOffset);
    gstWinningSheet.write(nMaxRowOffset, nDateColOffset, nInternalUrl + stTransCell, stPurpleBoldFormat, u"위로");

    # 평균 증감
    gstWinningSheet.write(nGraphRowOffset, nAvgStockColOffset, u"종목 평균", stNavyFormat);
    gstWinningSheet.write(nGraphRowOffset + 1, nAvgStockColOffset, u"증감율", stGreenTitleFormat);
    for nDateIndex in range(nMaxDateCount):
        if (nDateIndex == 0):
            continue;

        nDateRowOffset = nDateIndex + nStartGraphRowOffset;
        stStartTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset);
        stEndTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nMaxStockCount - 1);
        stString = "=IFERROR(AVERAGE(" + stStartTransCell + ":" + stEndTransCell + "), \"\")";
        gstWinningSheet.write(nDateRowOffset, nAvgStockColOffset, stString, stRateFormat);

    # KOSPI 승리
    EXCEL_PrintWinningDayRate(nGraphRowOffset, nKospiVsDayColOffset, u"KOSPI", nMaxDateCount);
    EXCEL_PrintWinningDayRate(nGraphRowOffset, nKosdaqVsDayColOffset, u"KOSDAQ", nMaxDateCount);

    # KOSPI 누적승리
    EXCEL_PrintWinningSumRate(nGraphRowOffset, nKospiVsSumColOffset, u"KOSPI", nMaxDateCount);
    EXCEL_PrintWinningSumRate(nGraphRowOffset, nKosdaqVsSumColOffset, u"KOSDAQ", nMaxDateCount);

    # 선정 종목 (그래프 취합 100개 제한)
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
            stString += "4, 5, \"" + gstSiseSheetName + "\"))";
            stString += ", \"\")";

            # 일반 선정 종목 증감율 값
            if (nRowOffset >= nStartGraphRowOffset):
                gstWinningSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString, stRateFormat);
            # 선정 종목 Title
            elif (nRowOffset >= nGraphRowOffset):
                gstWinningSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

                # 년도별 누적 승리율 출력
                if (nRowOffset == 2) and (nStockIndex < 7):
                    stTmpPreString = "=IFERROR(ROUNDUP(INDIRECT(ADDRESS(";
                    stTmpString = stDateString + str(nStartFnRowOffset+1) + ":" + stDateString + str(stMaxRowOffset) + ", 0)+" + str(nStartFnRowOffset) + "," + str(nKospiVsSumColOffset+1) + "))," + str(nRoundUp) + "),\"\")";
                    if (nStockIndex == 0): stString = stTmpPreString + "MATCH(\"11-01-03\", " + stTmpString;
                    if (nStockIndex == 1): stString = stTmpPreString + "MATCH(\"12-01-02\", " + stTmpString;
                    if (nStockIndex == 2): stString = stTmpPreString + "MATCH(\"13-01-02\", " + stTmpString;
                    if (nStockIndex == 3): stString = stTmpPreString + "MATCH(\"14-01-02\", " + stTmpString;
                    if (nStockIndex == 4): stString = stTmpPreString + "MATCH(\"15-01-02\", " + stTmpString;
                    if (nStockIndex == 5): stString = stTmpPreString + "MATCH(\"16-01-04\", " + stTmpString;
                    if (nStockIndex == 6): stString = stTmpPreString + str(stMaxRowOffset) + "," + str(nKospiVsSumColOffset+1) + "))," + str(nRoundUp) + "),\"\")";
                    gstWinningSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

            # 선정 종목 매핑 정보
            else:
                stString = "=IFERROR(";
                stString += "INDIRECT(ADDRESS(2 + MATCH(" + stStockColOffset + ", ";
                stString += gstFnSheetName + "!$A$" + stStartFnRowOffset + ":$A$" + stEndFnRowOffset + ", 0), 6, 4, 5, \"" + gstFnSheetName + "\"))";
                stString += ", \"\")";
                gstWinningSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

    # 차트 출력
    # 누적 승리율
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartGraphRowOffset, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKospiVsSumColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKospiVsSumColOffset);
    stKospiData = '=' + gstWinningSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKosdaqVsSumColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKosdaqVsSumColOffset);
    stKosdaqData = '=' + gstWinningSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartDateCell = xl_rowcol_to_cell(nStartGraphRowOffset, nDateColOffset);
    stEndDateCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nDateColOffset);
    stDate = '=' + gstWinningSheetName + '!' + stStartDateCell + ":" + stEndDateCell;

    stTitle = xl_rowcol_to_cell(1, nKospiVsSumColOffset);
    stChart.set_title({'name':u"누적 승리율"});
    stChart.set_x_axis({'name':u'날짜'});
    stChart.set_y_axis({'name':u'승리율(%)', 'min':0, 'max':200, 'major_unit':10});

    stChart.add_series({'name':u"KOSPI",  'categories':stDate, 'text_axis':True, 'values':stKospiData});
    stChart.add_series({'name':u"KOSDAQ", 'categories':stDate, 'text_axis':True, 'values':stKosdaqData});
    stChart.show_hidden_data();

    stChart.set_size({'width':1080, 'height':720});
    gstWinningSheet.insert_chart(stGraphCell, stChart);


    # KOSPI / KOSDAQ 지수
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartGraphRowOffset + 36, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKospiOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKospiOffset);
    stKospiSise = '=' + gstSiseSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nKosdaqOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nKosdaqOffset);
    stKosdaqSise = '=' + gstSiseSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stChart.set_title({'name':u"지수"});
    stChart.set_x_axis({'name':u'날짜'});
    stChart.set_y_axis({'name':u'KOSPI지수',   'num_format':'0'});
    stChart.set_y2_axis({'name':u'KOSDAQ지수', 'num_format':'0'});

    stChart.add_series({'name':u"KOSPI",  'categories':stDate, 'values':stKospiSise});
    stChart.add_series({'name':u"KOSDAQ", 'categories':stDate, 'values':stKosdaqSise, 'y2_axis':1});
    stChart.show_hidden_data();

    stChart.set_size({'width':1080, 'height':504});
    gstWinningSheet.insert_chart(stGraphCell, stChart);

def EXCEL_SetBenefitGraphXlsxData(nMaxDateCount, nMaxStockCount):
    nStockChoiceRowOffset = 0;
    nDateRowOffset = 1;
    nChoiceDateRowOffset = 2;
    nStartGraphRowOffset = 3;
    stDateString = 'B';
    stChoiceDateCell = 'B3';
    nMaxRowOffset = nMaxDateCount + nStartGraphRowOffset;
    stMaxRowOffset = str(nMaxRowOffset);
    stChoiceDate = u'16-01-04';
    nInternalUrl = u'internal:';

    nStartFnRowOffset = 3;
    nEndFnRowOffset = nStartFnRowOffset + nMaxStockCount - 1;
    stStartFnRowOffset = str(nStartFnRowOffset);
    stEndFnRowOffset = str(nEndFnRowOffset);

    stSiseCell = gstSiseSheetName + u'!';
    nGraphRowOffset = 1;
    nRowOffset = 0;

    nChoiceDateColOffset = 0;
    nDateColOffset = 1;
    nAvgStockColOffset = 2;
    nBuyCurColOffset = 3;
    nStockColOffset = 4;
    nRoundUp = 3;
    stBaseTransCell = xl_rowcol_to_cell(nGraphRowOffset + 1, nChoiceDateColOffset);

    stTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'blue'});
    stRedTitleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'red'});
    stChoiceDateFormat = gstWorkBook.add_format({'num_format': '@', 'font_color': 'blue'});
    stRedTitleFormat = gstWorkBook.add_format({'font_color': 'red'});
    stGreenTitleFormat = gstWorkBook.add_format({'font_color': 'green'});
    stPurpleBoldFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'purple'});
    stPurpleFormat = gstWorkBook.add_format({'font_color': 'purple'});
    stGrayFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'gray'});
    stNavyFormat = gstWorkBook.add_format({'bold': True, 'font_color': 'navy'});
    stRateFormat = gstWorkBook.add_format({'num_format':'0.000'});
    
    # 선정 날짜 기준 현재 수익률 동작 여부
    stString = "=IFERROR(MATCH(" + stChoiceDateCell + ", " + stDateString + str(nStartFnRowOffset+1) + ":" + stDateString + str(stMaxRowOffset) + ", 0) + 3";
    stString += ", 0)";
    gstBenefitSheet.write(nGraphRowOffset, nChoiceDateColOffset, u"선정 날짜", stGrayFormat);
    gstBenefitSheet.write(nGraphRowOffset + 1, nChoiceDateColOffset, stString, stGrayFormat);
    for nDateIndex in range(nMaxDateCount):
        nDateRowCellOffset = nDateIndex + nStartGraphRowOffset;
        stCurDateTransCell = xl_rowcol_to_cell(nDateRowCellOffset, nDateColOffset);
        stChoiceDateTransCell = xl_rowcol_to_cell(nChoiceDateRowOffset, nDateColOffset);

        stString = "=IF(AND(";
        stString += stCurDateTransCell + " >= " + stChoiceDateTransCell;
        stString += ", " + stBaseTransCell + " > 0";
        stString += "), 1, 0)";

        gstBenefitSheet.write(nDateRowCellOffset, nChoiceDateColOffset, stString);

    # 날짜 / KOSPI
    for nRowOffset in range(nMaxRowOffset):
        stTransCell = xl_rowcol_to_cell(nRowOffset, nDateColOffset - 1);
        stString = stSiseCell + stTransCell;
        stCellString = u'=' + "IF(" + stString + " > 0," + stString + ", \"\")";

        if (nRowOffset == nStockChoiceRowOffset):   # 0
            continue;
        elif (nRowOffset == nDateRowOffset):        # 1
            stTransCell = xl_rowcol_to_cell(nMaxRowOffset, nDateColOffset);
            gstBenefitSheet.write(nRowOffset, nDateColOffset, nInternalUrl + stTransCell, stPurpleBoldFormat, u"날짜");
        elif (nRowOffset == nChoiceDateRowOffset):  # 2
            gstBenefitSheet.write(nRowOffset, nDateColOffset, stChoiceDate, stChoiceDateFormat);
        else:                                       # >= 3
            gstBenefitSheet.write(nRowOffset, nDateColOffset, stCellString, stPurpleFormat);
    stTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nDateColOffset);
    gstBenefitSheet.write(nMaxRowOffset, nDateColOffset, nInternalUrl + stTransCell, stPurpleBoldFormat, u"위로");

    # 매입 시점별 수익률 증감
    gstBenefitSheet.write(nGraphRowOffset, nAvgStockColOffset, u"시점 수익", stNavyFormat);
    gstBenefitSheet.write(nGraphRowOffset + 1, nAvgStockColOffset, u"", stGreenTitleFormat);
    for nDateIndex in range(nMaxDateCount):
        nDateRowOffset = nDateIndex + nStartGraphRowOffset;
        stStartTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset);
        stEndTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nMaxStockCount - 1);
        stTransCellAgo = xl_rowcol_to_cell(nDateRowOffset - 1, nAvgStockColOffset);
        
        stString = "=IFERROR(";
        stString += "IF(COUNT(" + stStartTransCell + ":" + stEndTransCell + ") > 0, ";
        stString += "(";
        for nStockIndex in range (gnMaxGraphStockCount):
            stMaxTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nStockColOffset + nStockIndex);
            stCurTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nStockIndex);
            stString += "IFERROR(" + stMaxTransCell + " * 100 / " + stCurTransCell + ", 0)";
            if (nStockIndex < gnMaxGraphStockCount - 1):
                stString += " + ";
        stString += ") / ";
        stString += "COUNT(" + stStartTransCell + ":" + stEndTransCell + ")";
        stString += "- 100";
        stString += ", " + stTransCellAgo + ")";    # IF 문 종료
        
        stString += ", \"\")";
        gstBenefitSheet.write(nDateRowOffset, nAvgStockColOffset, stString, stRateFormat);

    # 선정 날짜 기준 현재 수익률
    gstBenefitSheet.write(nGraphRowOffset, nBuyCurColOffset, u"현재 수익", stNavyFormat);
    gstBenefitSheet.write(nGraphRowOffset + 1, nBuyCurColOffset, u"", stGreenTitleFormat);
    for nDateIndex in range(nMaxDateCount):
        nDateRowOffset = nDateIndex + nStartGraphRowOffset;
        stStartTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset);
        stEndTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nMaxStockCount - 1);
        stTransCellAgo = xl_rowcol_to_cell(nDateRowOffset - 1, nBuyCurColOffset);
        stChoiceDateTransCell = xl_rowcol_to_cell(nDateRowOffset, nChoiceDateColOffset);

        stString = "=IFERROR(";
        stString += "IF(AND(COUNT(" + stStartTransCell + ":" + stEndTransCell + ") > 0, " + stChoiceDateTransCell + " > 0, " + stBaseTransCell + " > 0), ";
        stString += "(";
        for nStockIndex in range (gnMaxGraphStockCount):
            stBaseAddrTransCell = "INDIRECT(ADDRESS(" + stBaseTransCell + ", " + str(nStockColOffset + nStockIndex + 1) + "))";
            stCurTransCell = xl_rowcol_to_cell(nDateRowOffset, nStockColOffset + nStockIndex);
            stString += "IFERROR(" + stCurTransCell + " * 100 / " + stBaseAddrTransCell + ", 0)";
            if (nStockIndex < gnMaxGraphStockCount - 1):
                stString += " + ";
        stString += ") / ";
        stString += "COUNT(" + stStartTransCell + ":" + stEndTransCell + ")";
        stString += " - 100";
        stString += ", " + stTransCellAgo + ")";    # IF 문 종료

        stString += ", \"\")";
        gstBenefitSheet.write(nDateRowOffset, nBuyCurColOffset, stString, stRateFormat);

    # 선정 종목 (그래프 취합 100개 제한)
    nStockCount = nMaxStockCount;
    if (nStockCount > gnMaxGraphStockCount):
        nStockCount = gnMaxGraphStockCount;
    for nStockIndex in range(nStockCount):
        for nRowOffset in range(nMaxRowOffset):
            stStockColOffset = str(nStockIndex + 1);
            stSiseRowOffset = str(nRowOffset + 1);
            stTransCellAgo = xl_rowcol_to_cell(nRowOffset - 1, nStockColOffset + nStockIndex);
            stTransCellAfter = xl_rowcol_to_cell(nRowOffset + 1, nStockColOffset + nStockIndex);

            stTransCell = xl_rowcol_to_cell(0, nStockColOffset + nStockIndex);
            stString = "=IFERROR(";
            stString += "IF(";
            stString += "INDIRECT(ADDRESS(" + stSiseRowOffset + ", " + stTransCell + "+1, ";
            stString += "4, 5, \"" + gstSiseSheetName + "\"))";
            stString += " > 0, ";
            stString += "INDIRECT(ADDRESS(" + stSiseRowOffset + ", " + stTransCell + "+1, ";
            stString += "4, 5, \"" + gstSiseSheetName + "\")), ";
            stString += "IF(" + stTransCellAfter + " > 0, " + stTransCellAfter + ", " + "\"\"))";  #IF
            stString += ", \"\")";  #IFERROR

            # 일반 선정 종목 증감율 값
            if (nRowOffset >= nStartGraphRowOffset):
                gstBenefitSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);
            # 선정 종목 Title
            elif (nRowOffset >= nGraphRowOffset):
                if (nRowOffset == 1):
                    stString = "=IFERROR(";
                    stString += "INDIRECT(ADDRESS(" + stSiseRowOffset + ", " + stTransCell + ", ";
                    stString += "4, 5, \"" + gstSiseSheetName + "\"))";
                    stString += ", \"\")";
                
                gstBenefitSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

                # 년도별 누적 승리율 출력
                if (nRowOffset == 2) and (nStockIndex < 7):
                    stTmpPreString = "=IFERROR(ROUNDUP(INDIRECT(ADDRESS(";
                    stTmpString = stDateString + str(nStartFnRowOffset+1) + ":" + stDateString + str(stMaxRowOffset) + ", 0)+" + str(nStartFnRowOffset) + "," + str(nAvgStockColOffset+1) + "))," + str(nRoundUp) + "),\"\")";
                    if (nStockIndex == 0): stString = stTmpPreString + "MATCH(\"11-01-03\", " + stTmpString;
                    if (nStockIndex == 1): stString = stTmpPreString + "MATCH(\"12-01-02\", " + stTmpString;
                    if (nStockIndex == 2): stString = stTmpPreString + "MATCH(\"13-01-02\", " + stTmpString;
                    if (nStockIndex == 3): stString = stTmpPreString + "MATCH(\"14-01-02\", " + stTmpString;
                    if (nStockIndex == 4): stString = stTmpPreString + "MATCH(\"15-01-02\", " + stTmpString;
                    if (nStockIndex == 5): stString = stTmpPreString + "MATCH(\"16-01-04\", " + stTmpString;
                    if (nStockIndex == 6): stString = stTmpPreString + str(stMaxRowOffset) + "," + str(nAvgStockColOffset+1) + "))," + str(nRoundUp) + "),\"\")";
                    gstBenefitSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

            # 선정 종목 매핑 정보
            else:
                stString = "=IFERROR(";
                stString += "INDIRECT(ADDRESS(2 + MATCH(" + stStockColOffset + ", ";
                stString += gstFnSheetName + "!$A$" + stStartFnRowOffset + ":$A$" + stEndFnRowOffset + ", 0), 6, 4, 5, \"" + gstFnSheetName + "\"))";
                stString += ", \"\")";
                gstBenefitSheet.write(nRowOffset, nStockColOffset + nStockIndex, stString);

    # 차트 출력
    # 누적 수익율
    stChart = gstWorkBook.add_chart({'type':'line'});
    stGraphCell = xl_rowcol_to_cell(nStartGraphRowOffset, nStockColOffset);

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nAvgStockColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nAvgStockColOffset);
    stBuyTimeData = '=' + gstBenefitSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartTransCell = xl_rowcol_to_cell(nStartGraphRowOffset, nBuyCurColOffset);
    stEndTransCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nBuyCurColOffset);
    stBuyAfterData = '=' + gstBenefitSheetName + '!' + stStartTransCell + ":" + stEndTransCell;

    stStartDateCell = xl_rowcol_to_cell(nStartGraphRowOffset, nDateColOffset);
    stEndDateCell = xl_rowcol_to_cell(nMaxRowOffset - 1, nDateColOffset);
    stDate = '=' + gstBenefitSheetName + '!' + stStartDateCell + ":" + stEndDateCell;

    stChart.set_title({'name':u"누적 수익율"});
    stChart.set_x_axis({'name':u'날짜'});
    stChart.set_y_axis({'name':u'수익율(%)', 'min':-10, 'max':400, 'major_unit':10});

    stChart.add_series({'name':u"매입 시점 수익률", 'categories':stDate, 'text_axis':True, 'values':stBuyTimeData});
    stChart.add_series({'name':u"매입 이후 수익률", 'categories':stDate, 'text_axis':True, 'values':stBuyAfterData});
    stChart.show_hidden_data();

    stChart.set_size({'width':1080, 'height':720});
    gstBenefitSheet.insert_chart(stGraphCell, stChart);

def EXCEL_WriteExcelFile(astKospiInfor, astKosdaqInfor, astStockInfor):
    PrintProgress(u"[시작] 엑셀 취합");
    nColOffset = 0;
    nRowOffset = 0;

    # 시세 Title 출력
    PrintProgress(u"[진행] 시세 Title 출력");
    EXCEL_SetSiseXlsxTitle(astKospiInfor);
    nColOffset = nColOffset + 1;
    EXCEL_SetKospiXlsxData(nColOffset, 'KOSPI', astKospiInfor, astKospiInfor);
    nColOffset = nColOffset + 2;
    EXCEL_SetKospiXlsxData(nColOffset, 'KOSDAQ', astKosdaqInfor, astKospiInfor);
    nColOffset = nColOffset + 2;
    PrintProgress(u"[완료] 시세 Title 출력");

    # 시세 데이터 출력
    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        PrintProgress(u"[진행] 시세 데이터 출력: " + str(nStockIndex + 1) + " / " + str(nStockLen) + " - " + astStockInfor[nStockIndex]['Name']);
        EXCEL_SetSiseXlsxData(nColOffset, astKospiInfor, astStockInfor[nStockIndex]);
        nColOffset = nColOffset + 2;
    PrintProgress(u"[완료] 시세 데이터 출력");

    # 재무 Title 출력
    PrintProgress(u"[진행] 재무 Title 출력");
    stAutoFilter = EXCEL_SetFnXlsxTitle(astStockInfor);
    nRowOffset = nRowOffset + 2;
    PrintProgress(u"[완료] 재무 Title 출력");

    # 재무 데이터 출력
    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        PrintProgress(u"[진행] 재무 데이터 출력: " + str(nStockIndex + 1) + " / " + str(nStockLen) + " - " + astStockInfor[nStockIndex]['Name']);
        EXCEL_SetFnXlsxData(nRowOffset, astStockInfor, nStockIndex);
        nRowOffset = nRowOffset + 1;
    PrintProgress(u"[완료] 재무 데이터 출력");

    # 승리율 그래프 출력
    if (gbWinningSheet > 0):
        PrintProgress(u"[진행] 승리율 그래프 출력");
        EXCEL_SetWinningRateGraphXlsxData(len(astKospiInfor), len(astStockInfor));
        PrintProgress(u"[완료] 승리율 그래프 출력");
    
    # 수익률 그래프 출력
    if (gbBenefitSheet > 0):
        PrintProgress(u"[진행] 수익률 그래프 출력");
        EXCEL_SetBenefitGraphXlsxData(len(astKospiInfor), len(astStockInfor));
        PrintProgress(u"[완료] 수익률 그래프 출력");
    
    PrintProgress(u"[완료] 엑셀 취합");
    return stAutoFilter;

def SISE_GetNonStockInfor(nStockCode, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    anUrl = "http://vip.mk.co.kr/newSt/rate/kospikosdaq_2.php?sty=2010&stm=5&std=1";

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
    stResponse = GetUrlOpen(nCodeUrl);
    stPage = stResponse.read();
    stSoup = BeautifulSoup(stPage);
    tables = stSoup.findAll('table');

    # 해당 Page로부터 시세 정보 확인 불가
    nTableIndex = 0;
    if (len(tables) <= 13):
        return 0;
    else:
        for nTableIndex in range(len(tables)):
            if (len(tables[nTableIndex].text.split(u'창출')) > 1):
                break;
            if (len(tables[nTableIndex].text.split(u'내용이 존재하지 않습니다')) > 1):
                return 0;
        if (nTableIndex == (len(tables) - 1)):
            return 0;

    astTable = tables[nTableIndex + 7].contents;
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

def COMPANY_GetNaverStockPageCount(nStockType):
    strType = 'sosok=';
    nUrl = 'http://finance.naver.com/sise/sise_market_sum.nhn?';

    strUrl = nUrl + strType + str(nStockType);
    stResponse = GetUrlOpen(strUrl);
    stPage = stResponse.read();
    stSoup = BeautifulSoup(stPage);

    astTrs = stSoup.findAll('tr');
    nTrLen = len(astTrs);
    astTrsContents = astTrs[nTrLen - 1];

    nTrsContentsLen = len(astTrsContents.contents);
    astTrsContentsContents = astTrsContents.contents[nTrsContentsLen - 2];
    nContentLen = len(astTrsContentsContents.contents);
    stContent = astTrsContentsContents.contents[nContentLen - 2];
    astValues = dict.values(stContent.attrs);
    nMaxPage  = astValues[0].split('page=')[1];

    return int(nMaxPage);

# Stock 이름과 코드를 얻는 함수
gastChangeStockNameCodeList = [];
def COMPANY_GetStockCode(astStockList): # OUT (gastChangeStockNameCodeList: 종목 이름 / 코드)
    astrStockType = ['KOSPI', 'KOSDAQ'];
    strType = 'sosok=';
    strPage = '&page=';
    nUrl = 'http://finance.naver.com/sise/sise_market_sum.nhn?';

    for nStockType in range(len(astrStockType)):
        PrintProgress(u"[시작] " + astrStockType[nStockType] + u" List 취합");

        nMaxPage = COMPANY_GetNaverStockPageCount(nStockType);

        for nPageIndex in range(1, nMaxPage + 1):
            strUrl = nUrl + strType + str(nStockType) + strPage + str(nPageIndex);
            stResponse = GetUrlOpen(strUrl);
            stPage = stResponse.read();
            stSoup = BeautifulSoup(stPage);

            astAhref = stSoup.findAll('a', {'class':'tltle'});
            nAhrefLen = len(astAhref);
            for nStockIndex in range(nAhrefLen):
                stStockNameCode = {};
                stStockNameCode['Name'] = astAhref[nStockIndex].text;

                stStockNameCode['Code'] = dict.values(astAhref[nStockIndex].attrs)[0].split('code=')[1];
                stStockNameCode['Type'] = astrStockType[nStockType];
                stStockNameCode['SISE'] = 0;
                stStockNameCode['Count'] = 0;
                astStockList.append(0);
                astStockList[len(astStockList) - 1] = copy.deepcopy(stStockNameCode);

                PrintProgress(u"[진행] " + astrStockType[nStockType] + u" List 취합 " + u"(" + str(len(astStockList)) + u")" + u" : " + u"<" + str(stStockNameCode['Code']) + u"> " + stStockNameCode['Name']);
        PrintProgress(u"[완료] " + astrStockType[nStockType] + u" List 취합");

def FILE_WriteBestStock(astStockInfor):
    stBestStockName = u'BestStock_' + gstDate + u'.txt';
    stFile = open(stBestStockName, 'w');

    nStockLen = len(astStockInfor);
    for nStockIndex in range(nStockLen):
        if (astStockInfor[nStockIndex]['BestStock'] > 0):
            stFile.write(astStockInfor[nStockIndex]['Code']);
            stFile.write(u' ');
            stFile.write(astStockInfor[nStockIndex]['Name'].encode("UTF-8"));
            stFile.write(u'\n');
    stFile.close();

############# main #############

gstDate = GetTodayString(ganYear, ganMonth, ganDay);

# 종목 코드 리스트 취합
COMPANY_GetStockCode(gastChangeStockNameCodeList);

# 재무 / 시세 정보 취합
COMPANY_GetFinanceInfor(gastChangeStockNameCodeList, gastStockInfor);

# Kospi / Kosdaq 정보 취합
SISE_GetKospiInfor(gastKospiInfor, gastKosdaqInfor);

# Best 종목 txt 출력
PrintProgress(u"[시작] 파일 출력");
FILE_WriteBestStock(gastStockInfor);
PrintProgress(u"[완료] 파일 출력");

# 종목 정보 출력
gstWorkBookName     = u'StockList_' + gstDate + u'.xlsx';
gstWorkBook         = xlsxwriter.Workbook(gstWorkBookName);

gstFnSheetName      = u'재무' + gstDate;
gstFnSheet          = gstWorkBook.add_worksheet(gstFnSheetName);
gstFnSheet.freeze_panes('E3');
gstFnSheet.set_column('A:B', None, None, {'hidden': 1});

gstSiseSheetName    = u'시세' + gstDate;
gstSiseSheet        = gstWorkBook.add_worksheet(gstSiseSheetName);
gstSiseSheet.freeze_panes('F4');
gstSiseSheet.set_row(0, None, None, {'hidden': True})

if (gbWinningSheet > 0):
    gstWinningSheetName   = u'승리율' + gstDate;
    gstWinningSheet       = gstWorkBook.add_worksheet(gstWinningSheetName);
    gstWinningSheet.freeze_panes('J4');
    gstWinningSheet.set_row(0, None, None, {'hidden': True})
    gstWinningSheet.set_column('A:A', None, None, {'hidden': True})
    gstWinningSheet.set_column('C:I', None, None, {'hidden': True})

if (gbBenefitSheet > 0):
    gstBenefitSheetName = u'수익률' + gstDate;
    gstBenefitSheet     = gstWorkBook.add_worksheet(gstBenefitSheetName);
    gstBenefitSheet.freeze_panes('E4');
    gstBenefitSheet.set_row(0, None, None, {'hidden': True})
    gstBenefitSheet.set_column('A:A', None, None, {'hidden': True})
    gstBenefitSheet.set_column('C:D', None, None, {'hidden': True})

# 엑셀 출력
gstAutoFilterEndCell = EXCEL_WriteExcelFile(gastKospiInfor, gastKosdaqInfor, gastStockInfor);
gstFnSheet.autofilter(gstAutoFilterStartCell + ':' + gstAutoFilterEndCell);

PrintProgress(u"[시작] 엑셀 출력");
gstWorkBook.close();
PrintProgress(u"[완료] 엑셀 출력");

PrintProgress(u"Complete all process");

#-*- coding: utf-8 -*-

import requests;
import pandas as pd;
import pandas.io.data as web;
from StringIO import StringIO;
import datetime;
from bs4 import BeautifulSoup;
import copy;
import matplotlib.pyplot as plt;

# Stock �̸��� �ڵ带 ��� �Լ�
gastStockList = {};
def COMPANY_GetStockCode(astStockList): # OUT (gastStockList: ���� �̸� / �ڵ�)
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

# Date & ������ ��� �Լ� (�ڽ��� / �ڽ��� / �Ϲ�����)
# �ڽ��� or �ڽ��� or �Ϲ� ���� ����
#gnStockCode             = 'KOSPI';      # '1997-07-01' ~
#gnStockCode             = 'KOSDAQ';     # '2013-03-04' ~
gnStockCode             = '014530';     # '2000-0101' ~
gastStockInfor          = [];
def SISE_GetStockInfor(nStockCode, stStockInfor):   # IN (nStock: �����ڵ�), OUT (stStockInfor: ���� ����)
    stDataInfor = {};

    if (nStockCode.isdigit()):                      # �Ϲ� ������ ���
#        stStartDate             = datetime.datetime(1900, 1, 1);
        stStartDate             = datetime.datetime(2014, 12, 30);
        stDataInfor             = web.DataReader(nStockCode + ".KS", "yahoo", stStartDate);
    else:                                           # �ڽ��� / �ڽ����� ���
        anReqCode               = {};
        anReqCode['KOSPI']      = '^KS11';
        anReqCode['KOSDAQ']     = '^KQ11';

        # nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=0&b=1&c=1900';

        # Month = a + 1 / Day = b / Year = c
        nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=11&b=30&c=2014';
        stRequest               = requests.get(nUrl);
        stDataInfor             = pd.read_csv(StringIO(stRequest.content), index_col='Date', parse_dates={'Date'});

    for nIndex in range(stDataInfor.shape[0]):
        stStock             = {};
        stStock['Code']     = nStockCode;                               # ���� �ڵ�
        stStock['Date']     = stDataInfor.index[nIndex]._date_repr;     # ��¥
        stStock['Price']    = stDataInfor.values[nIndex][3];            # ����: 'Close'
        stStockInfor.append(stStock);


gastStockNameCode = [];
def COMPANY_GetNameCode(astStockList, astStockNameCode):   # IN (nStock: �����ڵ�), OUT (stStockInfor: ���� ����)
    pFile = open("StockList.txt");
    anStockItem = pFile.readlines();
    stStockNameCode = {};

    for nStockItem in anStockItem:
        nStockName = nStockItem.split('\n')[0];
        stStockNameCode['Name'] = nStockName.decode('949');
        stStockNameCode['Code'] = astStockList[stStockNameCode['Name']];
        stStockNameCode['SISE'] = 0;

        astStockNameCode.append(0);
        astStockNameCode[len(astStockNameCode) - 1] = copy.deepcopy(stStockNameCode);


def SISE_GetCompannySise(astStockNameCode):   # IN (nStock: �����ڵ�), OUT (stStockInfor: ���� ����)
    astStockInfor          = [];

    for stStockNameCode in astStockNameCode:
        SISE_GetStockInfor(stStockNameCode['Code'], astStockInfor);
        stStockNameCode['SISE'] = copy.deepcopy(astStockInfor);
        astStockInfor = [];

def CAL_GetKospiRate(astStockInfor, astRateInfor):
    nDateCount = len(astStockInfor);

    for nDateIndex in range(1, nDateCount):
        nRate = (astStockInfor[nDateIndex]['Price'] * 100) / astStockInfor[nDateIndex - 1]['Price'];
        astRateInfor[astStockInfor[nDateIndex]['Date']] = nRate - 100;

def CAL_GetStocksRate(astStockInfor, astRateInfor):
    nStockCount = len(astStockInfor);
    nDateCount = len(astStockInfor[0]['SISE']);

    for nDateIndex in range(1, nDateCount):
        nRate = 0;

        for nStockIndex in range(nStockCount):
            nRate = nRate + ((astStockInfor[nStockIndex]['SISE'][nDateIndex]['Price'] * 100) / astStockInfor[nStockIndex]['SISE'][nDateIndex - 1]['Price']);
        nRate = nRate / nStockCount;

        astRateInfor[astStockInfor[0]['SISE'][nDateIndex]['Date']] = nRate - 100;

def CAL_CompareKospi(astKospiInfor, astKospiRateInfor, astStocksRateInfor, astCompareDate, astCompareInfor):
    nDateCount = len(astKospiInfor);
    nAccuratedRate = 0;

    for nDateIndex in range(1, nDateCount):
        stCompareInfor  = {};
        nCompareDate    = astKospiInfor[nDateIndex]['Date'];
        nCompareRate    = astStocksRateInfor[nCompareDate] - astKospiRateInfor[nCompareDate];
        nAccuratedRate  += nCompareRate;

        anSplitStr = nCompareDate.split("-");

#        astCompareDate.append(anSplitStr[0] + anSplitStr[1] + anSplitStr[2]);
        astCompareDate.append(nDateIndex);

        astCompareInfor.append(nAccuratedRate);
#        stCompareInfor['Date'] = nCompareDate;
#        stCompareInfor['Rate'] = nCompareRate;
#        astCompareInfor.append(stCompareInfor);
#        stCompareInfor = stCompareInfor;

COMPANY_GetStockCode(gastStockList);
COMPANY_GetNameCode(gastStockList, gastStockNameCode);
SISE_GetCompannySise(gastStockNameCode);
gastStocksRateInfor = {};
CAL_GetStocksRate(gastStockNameCode, gastStocksRateInfor);

gnStockCode             = 'KOSPI';
gastKospiInfor          = [];
SISE_GetStockInfor(gnStockCode, gastKospiInfor);
gastKospiInfor.sort();
gastKospiRateInfor = {};
CAL_GetKospiRate(gastKospiInfor, gastKospiRateInfor);

gastCompareDate         = [];
gastCompareInfor        = [];
CAL_CompareKospi(gastKospiInfor, gastKospiRateInfor, gastStocksRateInfor, gastCompareDate, gastCompareInfor);

plt.plot(gastCompareDate, gastCompareInfor);
plt.show();

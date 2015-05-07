#-*- coding: utf-8 -*-

import requests;
import pandas as pd;
import pandas.io.data as web;
from StringIO import StringIO;
import datetime;


# 코스피 or 코스닥 or 일반 종목 선택
#gnStockCode             = 'KOSPI';      # '1997-07-01' ~
#gnStockCode             = 'KOSDAQ';     # '2013-03-04' ~
gnStockCode             = '014530';     # '2000-0101' ~

gastStockInfor          = [];


# Date & 가격을 얻는 함수 (코스피 / 코스닥 / 일반종목)
def SISE_GetStockInfor(nStockCode, stStockInfor):   # IN (nStock: 종목코드), OUT (stStockInfor: 종목 정보)
    stDataInfor = {};

    if (nStockCode.isdigit()):                      # 일반 종목일 경우
        stStartDate             = datetime.datetime(1900, 1, 1);
        stDataInfor             = web.DataReader(nStockCode + ".KS", "yahoo", stStartDate);
    else:                                           # 코스피 / 코스닥일 경우
        anReqCode               = {};
        anReqCode['KOSPI']      = '^KS11';
        anReqCode['KOSDAQ']     = '^KQ11';

        nUrl                    = 'http://real-chart.finance.yahoo.com/table.csv?s=' + anReqCode[nStockCode] + '&a=0&b=1&c=1900';
        stRequest               = requests.get(nUrl);
        stDataInfor             = pd.read_csv(StringIO(stRequest.content), index_col='Date', parse_dates={'Date'});

    for nIndex in range(stDataInfor.shape[0]):
        stStock             = {};
        stStock['Code']     = nStockCode;                               # 종목 코드
        stStock['Date']     = stDataInfor.index[nIndex]._date_repr;     # 날짜
        stStock['Price']    = stDataInfor.values[nIndex][3];            # 종가: 'Close'

        stStockInfor.append(stStock);


SISE_GetStockInfor(gnStockCode, gastStockInfor);            # IN (gnStockCode: 종목코드), OUT (gastStockInfor: 종목 정보)

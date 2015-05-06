#-*- coding: utf-8 -*-

import requests;
import pandas as pd;
import pandas.io.data as web;
from StringIO import StringIO;
import datetime;

# 코스피 or 코스닥 선택
#gnStockCode             = 'KOSPI';      # '1997-07-01' ~
#gnStockCode             = 'KOSDAQ';     # '2013-03-04' ~
gnStockCode             = '014530';     # '2000-01-01' ~

gastStockInfor          = [];
gstDataInfor            = {};


if (gnStockCode.isdigit()):
    gstStartDate            = datetime.datetime(1900, 1, 1);
    gstDataInfor            = web.DataReader(gnStockCode + ".KS", "yahoo", gstStartDate);
else:
    ganReqCode              = {};
    ganReqCode['KOSPI']     = '^KS11';
    ganReqCode['KOSDAQ']    = '^KQ11';

    gnUrl                   = 'http://real-chart.finance.yahoo.com/table.csv?s=' + ganReqCode[gnStockCode] + '&a=0&b=1&c=1900';
    gstRequest              = requests.get(gnUrl);
    gstDataInfor            = pd.read_csv(StringIO(gstRequest.content), index_col='Date', parse_dates={'Date'});


# Date & 가격을 얻는 함수 (코스피 / 코스닥 / 일반종복 가능)
def SISE_GetStockInfor(stDataInfor, astStockInfor, nStockCode):
    for nIndex in range(stDataInfor.shape[0]):
        stStock             = {};
        stStock['Code']     = nStockCode;
        stStock['Date']     = stDataInfor.index[nIndex]._date_repr;
        stStock['Price']    = stDataInfor.values[nIndex][3]; # Close

        astStockInfor.append(stStock);


SISE_GetStockInfor(gstDataInfor, gastStockInfor, gnStockCode);

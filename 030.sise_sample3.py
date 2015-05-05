#-*- coding: utf-8 -*-

import pandas.io.data as web;
import datetime;

gnStockCode      = '014530';
gstStartDate     = datetime.datetime(1900, 1, 1);
gastDataInfor    = web.DataReader(gnStockCode + ".KS", "yahoo", gstStartDate);

gastStockInfor   = [];

for nIndex in range(gastDataInfor.shape[0]):
    nYear   = gastDataInfor.Close._stat_axis.year[nIndex];
    nMonth  = gastDataInfor.Close._stat_axis.month[nIndex];
    nDay    = gastDataInfor.Close._stat_axis.day[nIndex];

    stStock = {};
    stStock['Code'] = gnStockCode
    stStock['Date'] = str(nYear) + '-' + str(nMonth) + '-' + str(nDay);
    stStock['Price'] = gastDataInfor.values[nIndex][3]; # Close

    gastStockInfor.append(stStock);

print(gastStockInfor);

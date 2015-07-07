#-*- coding: utf-8 -*-
import json;
import urllib2;
from bs4 import BeautifulSoup;

gnOpener = urllib2.build_opener()
gnOpener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

def insert_finance(astStockName, nMaxStockCount):
    nMaxPageRange = 10;

    for nPageIndex in range(nMaxPageRange):
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

############# main #############

gastStockName = [];
gnMaxBaeDangStockCount = 200;

insert_finance(gastStockName, gnMaxBaeDangStockCount);

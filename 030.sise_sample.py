#-*- coding: utf-8 -*-
#http://finance.naver.com/item/frgn.nhn?code=014530&page=1  hello

from bs4 import BeautifulSoup
import copy
import urllib2

print('Hello World');

# page
gnStockCode             = '014530';         # 주식 코드 번호
gnStrFindAll            = 'table';          # stSoup.findAll() 수행 인자
gnSelectTableFindAll    = 2;                # HTML 코드 중 파싱할 Table Offset (12개 Table (0~11) 중, 2번째)

gastTable               = [];               # 취합할 최종 Table
gnTableType0Size        = 9;                # 필드 개수
gnTableType1LoopCount   = 15;               # Table -> Data 로 전환하기 위한 Loop 횟수
gnDataOffsetSize        = gnTableType0Size; # Data 필드 개수


gnPageLoopCount = 5;                         # 취합할 Page 개수



def MakeTable(nPageEntry, gastTable):
    stAppendEntry = [];                     # Append 구조체
    nTableAppend = 0;                       # Table 필드 명 취합을 위해 사용
    nTableOffset = 0;                       # Table 필드 명 취합을 위해 사용
    nDataOffset = 0;                        # Data 필드 취합을 위해 사용
    nType = 0;                              # 현재 취합하는 Type (Table 필드명 or Data)
    bAppendData = 0;                        # Data 취합이 완료되었는지 여부

    for nTableIndex in range(gnTableType0Size): # Append할 구조체 크기 설정
        stAppendEntry.append(0);

    nPageIndex = nPageEntry + 1;

    opener = urllib2.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]                 # header define

    nUrl = "http://finance.naver.com/item/frgn.nhn?code=" + gnStockCode + "&page=" + str(nPageIndex);
    response = opener.open(nUrl)
    page = response.read()    
    
    stSoup = BeautifulSoup(page);
    stFindAll = stSoup.findAll(gnStrFindAll);

    
    stSelectTableText = stFindAll[gnSelectTableFindAll].text;
    stSplitText = stSelectTableText.split("\n");


    if (nPageEntry == 0):
        nTableType = 0;                                                 # Table 필드 + Data 취합
    else:
        nTableType = 1;                                                 # Data 만 취합

    for stSplitEntry in stSplitText:
        stSplitEntryText = stSplitEntry.split("\t");                    # '\t' 제거
        stSplit = stSplitEntryText[len(stSplitEntryText) - 1];

        # Table List : 날짜 / 종가 / 전일비 / 등락율 / ...
        if (nType == 0):
            if (nTableType == 0):
                if (stSplit != ''):
                    if ((nTableOffset != 7) and (nTableOffset != 8)):   # "기관(7)" / "외국인(8)" 은 추가하지 않음
                        stAppendEntry[nTableAppend] = stSplit;          # stAppendEntry 로 설정
                        nTableAppend = nTableAppend + 1;

                nTableOffset = nTableOffset + 1;
                if (nTableOffset >= gnTableType1LoopCount):
                    gastTable.append(0);
                    gastTable[0] = copy.deepcopy(stAppendEntry);        # stAppendEntry -> gastTable 로 복제
                    nType = 1;                                          # Data 취합으로 전환
            else:
                nTableOffset = nTableOffset + 1;
                if (nTableOffset >= gnTableType1LoopCount):             # Table 필드 취합은 Skip 한다.
                    nType = 1;
        # Data List
        else:
            if (stSplit != ''):
                stAppendEntry[nDataOffset] = stSplit;                   # stAppendEntry 로 설정
                nDataOffset = (nDataOffset + 1) % gnDataOffsetSize;
                bAppendData = 0;

                if (nDataOffset == 0):
                    gastTable.append(0);
                    nEntryIndex = len(gastTable) - 1;
                    gastTable[nEntryIndex] = copy.deepcopy(stAppendEntry);  # stAppendEntry -> gastTable 로 복제
                    bAppendData = 1;

for nPageIndex in range(gnPageLoopCount):
    MakeTable(nPageIndex, gastTable);

print gastTable;

# 대신증권 API
# 전종목 시가총액 구하기 예제

# 전 종목의 시가총액을 구한 후, 시가총액 역순으로 나열하는 예제입니다.
# CpSysDib.MarketEye 서비스를 이용하여 복수 종목의 상장 주식수와 현재가를 가져온 후
# 시가총액  = 상장주식수 * 현재가 식으로 계산했습니다
#
# 상장 주식 수 관련해서는 20억 이상 상장주식수를 가진 종목의 경우 천단위로 제공 하고 있어 아래 함수를 통해 확인하는 부분도 참고 바랍니다.
# CpUtil.CpCodeMgr 서비스 IsBigListingStock(code) : 상장 주식수 20억 이상 여부 리턴

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False

    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0

    def Request(self, codes, dataInfo):
        # 0: 종목코드 4: 현재가 20: 상장주식수
        rqField = [0, 4, 20]  # 요청 필드

        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(2)

        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            cur = self.objRq.GetDataValue(1, i)  # 종가
            listedStock = self.objRq.GetDataValue(2, i)  # 상장주식수

            maketAmt = listedStock * cur
            if g_objCodeMgr.IsBigListingStock(code):
                maketAmt *= 1000
            #            print(code, maketAmt)

            # key(종목코드) = tuple(상장주식수, 시가총액)
            dataInfo[code] = (listedStock, maketAmt)

        return True


class CMarketTotal():
    def __init__(self):
        self.dataInfo = {}

    def GetAllMarketTotal(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        allcodelist = codeList + codeList2
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d' % (len(allcodelist), len(codeList), len(codeList2)))

        objMarket = CpMarketEye()
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo)
                rqCodeList = []
                continue
        # end of for

        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo)

    def PrintMarketTotal(self):

        # 시가총액 순으로 소팅
        data2 = sorted(self.dataInfo.items(), key=lambda x: x[1][1], reverse=True)

        print('전종목 시가총액 순 조회 (%d 종목)' % (len(data2)))
        for item in data2:
            name = g_objCodeMgr.CodeToName(item[0])
            listed = item[1][0]
            markettot = item[1][1]
            print('%s 상장주식수: %s, 시가총액 %s' % (name, format(listed, ','), format(markettot, ',')))


if __name__ == "__main__":
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()
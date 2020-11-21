# 대신증권 API
# 시세 연속 조회 제한 확인용 예제
# 예제 설명:
#     본 예제는 거래소 종목(약 1350 여 종목) 전체에 대한 현재가 통신을 반복 호출 합니다.
#     플러스 API 를 이용한 시세 요청 건수는 과다 호출을 막기 위해 기본 15초당 60건으로 제한하고 있습니다
#     만약 15초에 60건을 초과할 경우 이후 요청은 자동으로 지연 요청 하게 되어 있습니다
#     이번 예제는 시세 요청 시 현재 남아 있는 통신 요청 건수를 확인하고, 만약 요청이 지연되는 경우 지연 요청이 해제되기까지 남은 시간을 구하는 방법을 설명하기 위해 만들었습니다.
#
# 클래스 설명
#   ■ CpTimeChecker - 시세 요청 건수 체크 함수
#   ■ CpRPCurrentPrice - 현재가 시세 요청
#
# 주요 PLUS 함수 설명
#     CpUtil.CpCybos.GetLimitRemainCount - 지연 되지 않고 통신 가능한 횟수
#     CpUtil.CpCybos.LimitRequestRemainTime - GetLimitRemainCount 가 0 일 경우 남은 딜레이 시간을 구함(해당 시간이 0 이 되면 다음번 통신이 가능)

import sys
from PyQt5.QtWidgets import *
from enum import Enum
import win32com.client
import time
import pythoncom

g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")


# 감시 중인 현재가 정보 저장 구조체
class stockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "동시호가/장중 아님", ord('1'): "동시호가", ord('2'): "장중"}
        self.code = ""
        self.name = ""
        self.cur = 0  # 현재가
        self.diff = 0
        self.diffp = 0
        self.baseprice = 0
        self.offer = []  # 매도호가
        self.bid = []  # 매수호가
        self.offervol = []  # 매도호가 잔량
        self.bidvol = []  # 매수호가 잔량
        self.vievent = ""
        self.vievent2 = ""
        self.objCur = None
        self.objBid = None
        self.vistartprice = 0  # 발동 시점 현재가
        self.viendprice = 0  # 발동 해제 현재가

        # 예상체결가 정보
        self.exFlag = 0
        self.expcur = 0
        self.expdiff = 0
        self.expdiffp = 0

        # vi 관련
        self.viBase = 0
        self.viexUp = 0
        self.viexDown = 0

    def debugPrint(self):
        print("%s, 코드 %s %s, 현재가 %d 대비 %d, (%.2f), 1차매도 %d(%d) 1차매수 %d(%d) 상승VI %d 하락 VI %d "
              % (self.dicEx.get(self.exFlag), self.code,
                 self.name, self.cur, self.diff, self.diffp,
                 self.offer[0], self.offervol[0], self.bid[0], self.bidvol[0],
                 self.viexUp, self.viexDown))


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")

        return

    def Request(self, code, rtMst):
        # 현재가 통신

        rqtime = time.time()

        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        timeEllpase = time.time() - rqtime
        print("통신 시간", timeEllpase)

        # 수신 받은 현재가 정보를 rtMst 에 저장
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur = self.objStockMst.GetHeaderValue(11)  # 종가
        rtMst.diff = self.objStockMst.GetHeaderValue(12)  # 전일대비
        rtMst.baseprice = self.objStockMst.GetHeaderValue(27)  # 기준가
        rtMst.exFlag = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        if rtMst.baseprice:
            rtMst.diffp = (rtMst.diff / rtMst.baseprice) * 100

        rtMst.viBase = self.objStockMst.GetHeaderValue(80)  # 정적VI 발동 예상기준가
        rtMst.viexUp = self.objStockMst.GetHeaderValue(81)  # 정적VI 발동 예상상승가
        rtMst.viexDown = self.objStockMst.GetHeaderValue(82)  # 정적VI 발동 예상하락가

        # 10차호가
        for i in range(10):
            rtMst.offer.append(self.objStockMst.GetDataValue(0, i))  # 매도호가
            rtMst.bid.append(self.objStockMst.GetDataValue(1, i))  # 매수호가
            rtMst.offervol.append(self.objStockMst.GetDataValue(2, i))  # 매도호가 잔량
            rtMst.bidvol.append(self.objStockMst.GetDataValue(3, i))  # 매수호가 잔량
        return True


class CpTimeChecker:
    def __init__(self, checkType):
        self.chekcType = checkType  # 0: 주문 관련 1: 시세 요청 관련 2: 실시간 요청 관련

    def checkRemainTime(self):
        # 연속 요청 가능 여부 체크
        remainTime = g_objCpStatus.LimitRequestRemainTime
        remainCount = g_objCpStatus.GetLimitRemainCount(self.chekcType)  # 시세 제한
        print("남은 시간", remainTime, "남은 개수", remainCount)

        if remainCount <= 0:
            timeStart = time.time()
            while remainCount <= 0:
                #                pythoncom.PumpWaitingMessages()
                time.sleep(remainTime / 1000)
                remainCount = g_objCpStatus.GetLimitRemainCount(1)  # 시세 제한
                remainTime = g_objCpStatus.LimitRequestRemainTime  #
                print(remainCount, remainTime)
            ellapsed = time.time() - timeStart
            print("시간 지연: ", ellapsed, "남은 시세 요청 개수:", remainCount, "시간:", remainTime)


# 샘플 코드  메인 클래스
class testMain():
    def __init__(self):
        self.dicCodes = dict()  # 감시 종목 저장할 DICTIONARY
        self.obj = CpRPCurrentPrice()
        self.cptime = CpTimeChecker(1)
        return

    def ReqeustAllMst(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥

        timeStart = time.time()
        print("거래소 종목코드", len(codeList))
        for i, code in enumerate(codeList):
            # 다음번 통신 시간 확인을 위해 시간 체크
            self.cptime.checkRemainTime()

            name = g_objCodeMgr.CodeToName(code)
            newCur = stockPricedData()
            print(i + 1, time.ctime())
            self.obj.Request(code, newCur)
            self.dicCodes[code] = newCur

        ellapsed = time.time() - timeStart

        for key in self.dicCodes:
            self.dicCodes[key].debugPrint()

        print("전체 수행 시간 : ", ellapsed)


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.main = testMain()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 230)

        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)

        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)

    def btnStart_clicked(self):
        self.main.ReqeustAllMst()
        return

    def btnExit_clicked(self):
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
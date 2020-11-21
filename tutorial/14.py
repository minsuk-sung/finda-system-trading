# 대신증권 API
# VI 발동 종목에 대한 현재가 변화 추이 감시 예제

# 예제 설명:
#     VI 발동/해제 신호를 실시간으로 전송 받아, 해당 종목에 대한 현재가/예상체결/10차 호가 변화를 모니터링하는 샘플
#
#   ■ stockPricedData: 감시 중인 현재가 정보 저장 구조체
#   ■ CpEvent: 실시간 이벤트 수신
#   ■ CpPBStockCur : 실시간 현재가 요청
#   ■ CpPBStockBid : 실시간 10차 호가 요청
#   ■ CpRPCurrentPrice : 실시간 현재가 기본 조회
#   ■ testMain : 이번 샘플에 대한 메인 기능을 제공하며 실시간 vi 발동 종목에 대한 감시 수행

import sys
from PyQt5.QtWidgets import *
from enum import Enum
import win32com.client

g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")


# 감시 중인 현재가 정보 저장 구조체
class stockPricedData:
    def __init__(self):
        self.name = ""
        self.cur = 0  # 현재가
        self.diff = 0
        self.diffp = 0
        self.baseprice = 0
        self.offer = []  # 매도호가
        self.bid = []  # 매수호가
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


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, parent):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.parent = parent  # callback 을 위해 보관

        # 데이터 변환용
        self.dic2 = {ord('1'): "종목별 VI", ord('2'): "배분정보", ord('3'):
            "기준가결정", ord('4'): "임의종료", ord('5'): "종목정보공개", ord('6'): "종목조치", ord('7'): "시장조치"}
        # print(self.dic2)

    # PLUS 로 부터 실제로 시세를 수신 받는 이벤트 핸들러
    def OnReceived(self):
        if self.name == "9619s":
            # 시장조치 실시간 PB
            time = self.client.GetHeaderValue(0)
            flag = self.client.GetHeaderValue(1)
            print(self.name, self.dic2.get(flag))
            if self.dic2.get(flag) == "종목별 VI":
                code = self.client.GetHeaderValue(3)
                exptime = self.client.GetHeaderValue(7)
                event = self.client.GetHeaderValue(5)  # 조치내용(알리미 표시내용)
                event2 = self.client.GetHeaderValue(6)  # 변경사항
                print("시간:", time, code, g_objCodeMgr.CodeToName(code), exptime, event, event2)
                self.parent.monitorVI(time, code, exptime, event, event2)

            return

        if self.name == "stockcur":
            # 현재가 체결 데이터 실시간 업데이트
            code = self.client.GetHeaderValue(0)
            diff = self.client.GetHeaderValue(2)
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            # 장중이 아니면 처리 안함.

            # 현재가 업데이트
            self.parent.monitorPriceChange(code, exFlag, cprice, diff)

            return

        elif self.name == "stockbid":
            # 현재가 10차 호가 데이터 실시간 업데이트
            code = self.client.GetHeaderValue(0)
            dataindex = [3, 4, 7, 8, 11, 12, 15, 16, 19, 20, 27, 28, 31, 32, 35, 36, 39, 40, 43, 44]
            obi = 0
            offer = []
            bid = []
            for i in range(10):
                offer.append(self.client.GetHeaderValue(dataindex[obi]))
                bid.append(self.client.GetHeaderValue(dataindex[obi + 1]))
                obi += 2

            # 10차 호가 변경 call back 함수 호출
            self.parent.monitorOfferbidChange(code, offer, bid)

            return


# SB/PB 요청 ROOT 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, parent):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, parent)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__("stockcur", "DsCbo1.StockCur")


# CpPBStockBid: 실시간 10차 호가 요청 클래스
class CpPBStockBid(CpPublish):
    def __init__(self):
        super().__init__("stockbid", "Dscbo1.StockJpBid")


# CpPB9619 : 실시간 시장 공개 정보(vi 포함) 요청 클래스
class CpPB9619(CpPublish):
    def __init__(self):
        super().__init__("9619s", "CpSysDib.CpSvr9619s")


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        self.objStockjpbid = win32com.client.Dispatch("DsCbo1.StockJpBid2")

        return

    def Request(self, code, rtMst):
        # 현재가 통신
        self.objStockMst.SetInputValue(0, code)
        self.objStockMst.BlockRequest()

        # 10차 호가 통신
        self.objStockjpbid.SetInputValue(0, code)
        self.objStockjpbid.BlockRequest()

        print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
        if self.objStockMst.GetDibStatus() != 0:
            return False
        print("통신상태", self.objStockjpbid.GetDibStatus(), self.objStockjpbid.GetDibMsg1())
        if self.objStockjpbid.GetDibStatus() != 0:
            return False

        # 수신 받은 현재가 정보를 rtMst 에 저장
        rtMst.cur = self.objStockMst.GetHeaderValue(11)  # 종가
        rtMst.diff = self.objStockMst.GetHeaderValue(12)  # 전일대비
        rtMst.baseprice = self.objStockMst.GetHeaderValue(27)  # 기준가
        if rtMst.baseprice:
            rtMst.diffp = (rtMst.diff / rtMst.baseprice) * 100

        # 10차호가
        for i in range(10):
            rtMst.offer.append(self.objStockjpbid.GetDataValue(0, i))  # 매도호가
            rtMst.bid.append(self.objStockjpbid.GetDataValue(1, i))  # 매수호가

        # for debug
        #        for i in range(10):
        #            print(i+1, "차 매도/매수 호가: ", rtMst.offer[i], rtMst.bid[i])
        print("현재가 통신", g_objCodeMgr.CodeToName(code), "현재가: ", rtMst.cur, "대비:",
              rtMst.diff, "대비%:", round(rtMst.diffp, 2), "1차매도:", rtMst.offer[0], "1차매수:", rtMst.bid[0])

        return True


# 샘플 코드  메인 클래스
class testMain():
    def __init__(self):
        self.isSB = False  # 실시간 처리

        # VI 발동현황 요청
        self.pb9619 = CpPB9619()
        self.pb9619.Subscribe("", self)
        self.dicCodes = dict()  # 감시 종목 저장할 DICTIONARY
        return

    def stopSubscribe(self):
        if self.isSB:
            self.pb9619.Unsubscribe()

            for key in self.dicCodes:
                self.dicCodes[key].objCur.Unsubscribe()
                self.dicCodes[key].objBid.Unsubscribe()

        self.isSB = False
        return

    def monitorVI(self, time, code, exptime, event, event2):
        print("종목 VI", time, code, g_objCodeMgr.CodeToName(code), exptime, event, event2)

        if (event.find("발동") > 0):
            if code in self.dicCodes:
                self.dicCodes[code].vievent = event
                self.dicCodes[code].vievent2 = event2
            else:  # 신규 종목 추가
                newCur = stockPricedData()
                # 현재가 통신
                mst = CpRPCurrentPrice()
                if mst.Request(code, newCur) == False:
                    return

                newCur.name = g_objCodeMgr.CodeToName(code)
                newCur.objCur = CpPBStockCur()
                newCur.objBid = CpPBStockBid()
                newCur.vistartprice = newCur.cur  # 발동 시점 현재가
                newCur.viendprice = 0
                self.dicCodes[code] = newCur

                self.dicCodes[code].objCur.Subscribe(code, self)
                self.dicCodes[code].objBid.Subscribe(code, self)

            print("VI 발동 종목 감시 시작", code, self.dicCodes[code].name)
            print("김시 중인 종목 #", len(self.dicCodes))

        elif (event.find("해제") > 0):
            if not (code in self.dicCodes):  # 감시 중이지 않은 종목이면 skip
                return;

            self.dicCodes[code].viendprice = self.dicCodes[code].cur

            print("VI 해제 종목 삭제", code, self.dicCodes[code].name,
                  "발동시점 현재가:", self.dicCodes[code].vistartprice, "발동해제 현재가:", self.dicCodes[code].viendprice)
            self.dicCodes[code].objCur.Unsubscribe()
            self.dicCodes[code].objBid.Unsubscribe()
            del self.dicCodes[code]

        return

    def monitorPriceChange(self, code, exFlag, cur, diff):
        if not (code in self.dicCodes):
            print("error")
            return
        diffp = 0
        dicex = {ord('1'): "예상체결", ord('2'): "체결"}
        if exFlag == ord('1'):  # 예상체결
            self.dicCodes[code].expcur = cur
            self.dicCodes[code].expdiff = diff
            if self.dicCodes[code].baseprice:
                diffp = self.dicCodes[code].expdiffp = (diff / self.dicCodes[code].baseprice) * 100
        elif exFlag == ord('2'):  # 장중 체결
            self.dicCodes[code].cur = cur
            self.dicCodes[code].diff = diff
            if self.dicCodes[code].baseprice:
                diffp = self.dicCodes[code].diffp = (diff / self.dicCodes[code].baseprice) * 100

        print("현재가 PB ", dicex.get(exFlag), code, g_objCodeMgr.CodeToName(code), cur, "대비:", diff, "대비율:",
              round(diffp, 2))

        return

    def monitorOfferbidChange(self, code, offer, bid):
        if not (code in self.dicCodes):
            print("error")
            return

        print("10차 호가 PB ", code, g_objCodeMgr.CodeToName(code), self.dicCodes[code].offer[0],
              self.dicCodes[code].bid[0])
        self.dicCodes[code].offer = offer
        self.dicCodes[code].bid = bid
        return


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.main = testMain()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 230)

        btnStop = QPushButton("종료", self)
        btnStop.move(20, 70)
        btnStop.resize(200, 30)
        btnStop.clicked.connect(self.btnStop_clicked)

    def btnStop_clicked(self):
        self.main.stopSubscribe()
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
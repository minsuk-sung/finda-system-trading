# 대신증권 API
# 주식 현재가 조회/실시간
# 거래 중일때 한번 해봐야겠다

import sys
from PyQt5.QtWidgets import *
import win32com.client


class CpEvent:
    instance = None

    def OnReceived(self):
        # time = CpEvent.instance.GetHeaderValue(3)  # 시간
        timess = CpEvent.instance.GetHeaderValue(18)  # 초
        exFlag = CpEvent.instance.GetHeaderValue(19)  # 예상체결 플래그
        cprice = CpEvent.instance.GetHeaderValue(13)  # 현재가
        diff = CpEvent.instance.GetHeaderValue(2)  # 대비
        cVol = CpEvent.instance.GetHeaderValue(17)  # 순간체결수량
        vol = CpEvent.instance.GetHeaderValue(9)  # 거래량

        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)


class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        CpEvent.instance = self.objStockCur
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


class CpStockMst:
    def Request(self, code):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 현재가 객체 구하기
        objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        objStockMst.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objStockMst.GetDibStatus()
        rqRet = objStockMst.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # 현재가 정보 조회
        code = objStockMst.GetHeaderValue(0)  # 종목코드
        name = objStockMst.GetHeaderValue(1)  # 종목명
        time = objStockMst.GetHeaderValue(4)  # 시간
        cprice = objStockMst.GetHeaderValue(11)  # 종가
        diff = objStockMst.GetHeaderValue(12)  # 대비
        open = objStockMst.GetHeaderValue(13)  # 시가
        high = objStockMst.GetHeaderValue(14)  # 고가
        low = objStockMst.GetHeaderValue(15)  # 저가
        offer = objStockMst.GetHeaderValue(16)  # 매도호가
        bid = objStockMst.GetHeaderValue(17)  # 매수호가
        vol = objStockMst.GetHeaderValue(18)  # 거래량
        vol_value = objStockMst.GetHeaderValue(19)  # 거래대금

        print("코드 이름 시간 현재가 대비 시가 고가 저가 매도호가 매수호가 거래량 거래대금")
        print(code, name, time, cprice, diff, open, high, low, offer, bid, vol, vol_value)
        return True


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 1000, 1000)
        self.isRq = False
        self.objStockMst = CpStockMst()
        self.objStockCur = CpStockCur()

        btn1 = QPushButton("요청 시작", self)
        btn1.move(200, 20)
        btn1.resize(200,100)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("요청 종료", self)
        btn2.move(200, 300)
        btn2.resize(200, 100)
        btn2.clicked.connect(self.btn2_clicked)

        btn3 = QPushButton("종료", self)
        btn3.move(200, 600)
        btn3.resize(200, 100)
        btn3.clicked.connect(self.btn3_clicked)

    def StopSubscribe(self):
        if self.isRq:
            self.objStockCur.Unsubscribe()
        self.isRq = False

    def btn1_clicked(self):
        testCode = "A000660"
        if (self.objStockMst.Request(testCode) == False):
            exit()

        # 하이닉스 실시간 현재가 요청
        self.objStockCur.Subscribe(testCode)

        print("빼기빼기================-")
        print("실시간 현재가 요청 시작")
        self.isRq = True

    def btn2_clicked(self):
        self.StopSubscribe()

    def btn3_clicked(self):
        self.StopSubscribe()
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
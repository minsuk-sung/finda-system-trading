# 대신증권 API
# 주식 복수종목 실시간 등록/해지
# 요청 시작 : 3개 종목에 대해서 실시간 등록
# 요청 종료 : 3종목 실시간 해지
# 이게 도대체 뭘 의미하는거지?

import sys
from PyQt5.QtWidgets import *
import win32com.client


class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 종목코도
        name = self.client.GetHeaderValue(1)  # 종목명
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량

        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)


class CpStockCur:
    def Subscribe(self, code):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 150)
        self.isSB = False
        self.objStockCur1 = CpStockCur()
        self.objStockCur2 = CpStockCur()
        self.objStockCur3 = CpStockCur()

        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)

        btnStop = QPushButton("요청 종료", self)
        btnStop.move(20, 70)
        btnStop.clicked.connect(self.btnStop_clicked)

        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)

    def StopSubscribe(self):
        if self.isSB:
            self.objStockCur1.Unsubscribe()
            self.objStockCur2.Unsubscribe()
            self.objStockCur3.Unsubscribe()

        self.isSB = False

    def btnStart_clicked(self):
        self.objStockCur1.Subscribe("A003540")  # 대신증권
        self.objStockCur2.Subscribe("A000660")  # 하이닉스
        self.objStockCur3.Subscribe("A005930")  # 삼성전자

        print("빼기빼기================-")
        print("실시간 현재가 요청 시작")
        self.isSB = True

    def btnStop_clicked(self):
        self.StopSubscribe()

    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
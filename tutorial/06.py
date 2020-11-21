# 대신증권 API
# 주식 복수종목 조회/실시간
#

import sys
from PyQt5.QtWidgets import *
import win32com.client


# 복수 종목 실시간 조회 샘플 (조회는 없고 실시간만 있음)
class CpEvent:
    def set_params(self, client):
        self.client = client

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
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


class CpMarketEye:
    def Request(self, codes, rqField):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        # rqField = [0,17, 1,2,3,4,10]
        objRq.SetInputValue(0, rqField)  # 요청 필드
        objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        rqRet = objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = objRq.GetHeaderValue(2)

        for i in range(cnt):
            rpCode = objRq.GetDataValue(0, i)  # 코드
            rpName = objRq.GetDataValue(1, i)  # 종목명
            rpTime = objRq.GetDataValue(2, i)  # 시간
            rpDiffFlag = objRq.GetDataValue(3, i)  # 대비부호
            rpDiff = objRq.GetDataValue(4, i)  # 대비
            rpCur = objRq.GetDataValue(5, i)  # 현재가
            rpVol = objRq.GetDataValue(6, i)  # 거래량
            print(rpCode, rpName, rpTime, rpDiffFlag, rpDiff, rpCur, rpVol)

        return True


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 150)
        self.isSB = False
        self.objCur = []

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
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False

        self.objCur = []

    def btnStart_clicked(self):
        self.StopSubscribe();

        # 요청 종목 배열
        codes = ["A003540", "A000660", "A005930", "A035420", "A069500", "Q530031"]
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        rqField = [0, 1, 2, 3, 4, 10, 17]  # 요청 필드
        objMarkeyeye = CpMarketEye()
        if (objMarkeyeye.Request(codes, rqField) == False):
            exit()

        cnt = len(codes)
        for i in range(cnt):
            self.objCur.append(CpStockCur())
            self.objCur[i].Subscribe(codes[i])

        print("빼기빼기================-")
        print(cnt, "종목 실시간 현재가 요청 시작")
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
# 대신증권 API
# 해외선물 현재가/5차 호가 조회(실시간 업데이트 포함)

# 해외선물 현재가를 조회하는 기본 샘플입니다.
# 샘플 편의를 위해 종목코드를 'QMX17' 종목으로 코딩 되어 있으니 원하시는 코드로 변경 하시면 됩니다
#
# ■ 해외 선물 현재가 조회
#     CpForeDib.OvFutMst
# ■ 해외선물 실시간 시세 조회
#     CpForeDib.OvFutCur - 실시간 체결
#     CpForeDib.OvFutBid - 실사간 5차 호가

import sys
from PyQt5.QtWidgets import *
import win32com.client

# cp object
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, dicData, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관
        self.dicData = dicData

    # PLUS 로 부터 실제로 시세를 수신 받는 이벤트 핸들러
    def OnReceived(self):
        if self.name == "ovfucur":
            # 현재가 체결 데이터 실시간 업데이트
            self.dicData['cur'] = self.client.GetHeaderValue(7)
            self.dicData['open'] = self.client.GetHeaderValue(14)
            self.dicData['high'] = self.client.GetHeaderValue(15)
            self.dicData['low'] = self.client.GetHeaderValue(16)
            self.dicData['offer'] = self.client.GetHeaderValue(22)  # 매도호가
            self.dicData['bid'] = self.client.GetHeaderValue(23)  # 매수호가
            self.dicData['diff'] = self.client.GetHeaderValue(9)  # 대비
            self.dicData['vol'] = self.client.GetHeaderValue(11)  # 거래량
            print('체결실시간', self.dicData)
        elif self.name == "ovfubid":
            sindx = 5
            for i in range(1, 6):
                offerkey = 'offer' + str(i)
                bidkey = 'bid' + str(i)
                self.dicData[offerkey] = self.client.GetHeaderValue(sindx)
                self.dicData[bidkey] = self.client.GetHeaderValue(sindx + 3)
                sindx += 6

            print('5차호가실시간', self.dicData)

            return


# SB/PB 요청 ROOT 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, dicData, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, dicData, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


class CpPBOvFuCur(CpPublish):
    def __init__(self):
        super().__init__("ovfucur", "CpForeDib.OvFutCur")


class CpPBOvFuOBid(CpPublish):
    def __init__(self):
        super().__init__("ovfubid", "CpForeDib.OvFutBid")


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPOvForMst:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objMst = win32com.client.Dispatch("CpForeDib.OvFutMst")
        self.objCur = CpPBOvFuCur()
        self.objBid = CpPBOvFuOBid()
        return

    def Request(self, code, dicData):
        # 현재가 통신
        self.objCur.Unsubscribe()
        self.objBid.Unsubscribe()
        self.objMst.SetInputValue(0, code)
        ret = self.objMst.BlockRequest()
        if self.objMst.GetDibStatus() != 0:
            print("통신상태", self.objMst.GetDibStatus(), self.objMst.GetDibMsg1())
            return False
        self.objCur.Subscribe(code, dicData, None)
        self.objBid.Subscribe(code, dicData, None)

        # 수신 받은 현재가 정보를 rtMst 에 저장
        dicData['code'] = code
        dicData['cur'] = self.objMst.GetHeaderValue(29)
        dicData['open'] = self.objMst.GetHeaderValue(35)
        dicData['high'] = self.objMst.GetHeaderValue(36)
        dicData['low'] = self.objMst.GetHeaderValue(37)
        dicData['tick'] = self.objMst.GetHeaderValue(6)  # 호가 단위
        dicData['consize'] = self.objMst.GetHeaderValue(9)  # 계약크기
        dicData['float'] = self.objMst.GetHeaderValue(4)  # 가격 소수점
        dicData['jinb'] = self.objMst.GetHeaderValue(5)  # 진법
        dicData['offer'] = self.objMst.GetHeaderValue(33)  # 매도호가
        dicData['bid'] = self.objMst.GetHeaderValue(34)  # 매수호가
        dicData['diff'] = self.objMst.GetHeaderValue(31)  # 대비
        dicData['vol'] = self.objMst.GetHeaderValue(32)  # 거래량

        sindx = 62
        for i in range(1, 6):
            offerkey = 'offer' + str(i)
            bidkey = 'bid' + str(i)
            dicData[offerkey] = self.objMst.GetHeaderValue(sindx)
            dicData[bidkey] = self.objMst.GetHeaderValue(sindx + 3)
            sindx += 6

        print(dicData)


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 180)
        self.dicCurData = {}

        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)

        btnStop = QPushButton("요청 종료", self)
        btnStop.move(20, 70)
        btnStop.clicked.connect(self.btnStop_clicked)

        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)

    def btnStart_clicked(self):
        objCur = CpRPOvForMst()
        objCur.Request('QMX17', self.dicCurData)

    def btnStop_clicked(self):
        return

    def btnExit_clicked(self):
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
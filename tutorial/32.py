# 대신증권 API
# 해외선물 주문 예제

# 플러스를 이용하여 해외선물 매수/매도 주문 하는 예제입니다.

# ■ 사용된 플러스 객체
# CpForeDib.OvFutMst - 해외선물 현재가 조회
# CpForeTrade.OvFutOrder - 해외선물 주문
# CpForeDib.OvFutCur - [실시간] 해외선물 체결 시세
# CpForeDib.OvFutBid - [실시간]해외선물 5차호가 시세

# 매수 주문은 1차 매수호가로, 매도 주문은 1차 매도호가를 이용하고
# 주문 종목은 에디트에 입력된 종목코드를 기준으로 1주  주문 합니다.

import ctypes
import sys

import time
from PyQt5.QtWidgets import *
import win32com.client

# cp object
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")


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

    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False

    return True


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
            # print('체결실시간', self.dicData)
        elif self.name == "ovfubid":
            sindx = 5
            for i in range(1, 6):
                offerkey = 'offer' + str(i)
                bidkey = 'bid' + str(i)
                self.dicData[offerkey] = self.client.GetHeaderValue(sindx)
                self.dicData[bidkey] = self.client.GetHeaderValue(sindx + 3)
                sindx += 6

            # print('5차호가실시간', self.dicData)

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
            dicData['offer' + str(i)] = self.objMst.GetHeaderValue(sindx)
            dicData['bid' + str(i)] = self.objMst.GetHeaderValue(sindx + 3)
            sindx += 6

        print(dicData)


# 주식 주문 처리
class CpRPOvForOrder:
    def __init__(self):

        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 64)  # 64: 해외선물 상품
        print(self.acc, self.accFlag[0])

        self.objFOvrOrder = win32com.client.Dispatch("CpForeTrade.OvFutOrder")  # 주문
        self.orderNum = 0  # 주문 번호

    def buyOrder(self, code, price, amount, caller):
        # 주식 매수 주문
        print("신규 매수", code, price, amount)

        self.objFOvrOrder.SetInputValue(0, self.acc)  # 계좌번호
        self.objFOvrOrder.SetInputValue(1, code)
        self.objFOvrOrder.SetInputValue(2, 'N')  # 주문구분: “N”-신규, ”M”-정정, ”C”-취소
        self.objFOvrOrder.SetInputValue(3, 'B')  # 매매구분: “S”-매도, ”B”-매수
        self.objFOvrOrder.SetInputValue(4, '1')  # 가격조건: “1”-지정가, “2”-시장가, ”3”-STOP MARKET, ”4”-STOP LIMIT
        self.objFOvrOrder.SetInputValue(6, price)
        self.objFOvrOrder.SetInputValue(7, amount)

        # 매수 주문 요청
        ret = self.objFOvrOrder.BlockRequest()
        if ret == 4:
            print('연속 주문 제한으로 오류')
            QMessageBox.warning(caller, '주문오류', '연속 주문 제한으로 오류')
            return False

        rqStatus = self.objFOvrOrder.GetDibStatus()
        rqRet = self.objFOvrOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            QMessageBox.warning(caller, '주문오류', rqRet)
            return False

        return True

    def sellOrder(self, code, price, amount, caller):
        # 주식 매수 주문
        print("신규 매도", code, price, amount)

        self.objFOvrOrder.SetInputValue(0, self.acc)  # 계좌번호
        self.objFOvrOrder.SetInputValue(1, code)
        self.objFOvrOrder.SetInputValue(2, 'N')  # 주문구분: “N”-신규, ”M”-정정, ”C”-취소
        self.objFOvrOrder.SetInputValue(3, 'S')  # 매매구분: “S”-매도, ”B”-매수
        self.objFOvrOrder.SetInputValue(4, '1')  # 가격조건: “1”-지정가, “2”-시장가, ”3”-STOP MARKET, ”4”-STOP LIMIT
        self.objFOvrOrder.SetInputValue(6, price)
        self.objFOvrOrder.SetInputValue(7, amount)

        # 매수 주문 요청
        ret = self.objFOvrOrder.BlockRequest()
        if ret == 4:
            QMessageBox.warning(caller, '주문오류', '연속 주문 제한으로 오류')
            return False

        rqStatus = self.objFOvrOrder.GetDibStatus()
        rqRet = self.objFOvrOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            QMessageBox.warning(caller, '주문오류', rqRet)
            return False

        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 180)
        self.dicCurData = {}

        nH = 20
        self.codeEdit = QLineEdit("", self)
        self.codeEdit.move(20, nH)
        self.codeEdit.textChanged.connect(self.codeEditChanged)
        self.codeEdit.setText('E7H18')
        self.label = QLabel('종목코드', self)
        self.label.move(140, nH)
        nH += 50

        btnBuy = QPushButton("매수주문", self)
        btnBuy.move(20, nH)
        btnBuy.clicked.connect(self.btnBuy_clicked)
        nH += 50

        btnSell = QPushButton("매도주문", self)
        btnSell.move(20, nH)
        btnSell.clicked.connect(self.btnSell_clicked)
        nH += 50

        btnExit = QPushButton("종료", self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50
        self.setGeometry(300, 300, 300, nH)

    def codeEditChanged(self):
        code = self.codeEdit.text()
        self.setCode(code)

    def setCode(self, code):
        if len(code) < 4:
            return

        # name = g_objCodeMgr.CodeToName(code)
        # if len(name) == 0:
        #     print("종목코드 확인")
        #     return
        #
        # self.label.setText(name)
        self.code = code

    def btnBuy_clicked(self):
        objCur = CpRPOvForMst()
        if False == objCur.Request(self.code, self.dicCurData):
            return

        # 매수 1호가
        price = self.dicCurData['bid1']
        amount = 1
        objOrder = CpRPOvForOrder()
        objOrder.buyOrder(self.code, price, amount, self)

        return

    def btnSell_clicked(self):
        objCur = CpRPOvForMst()
        if False == objCur.Request(self.code, self.dicCurData):
            return

        # 매도 1호가
        price = self.dicCurData['offer1']
        amount = 1
        objOrder = CpRPOvForOrder()
        objOrder.sellOrder(self.code, price, amount, self)
        return

    def btnExit_clicked(self):
        exit()


if __name__ == "__main__":
    if InitPlusCheck() == False:
        exit()
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
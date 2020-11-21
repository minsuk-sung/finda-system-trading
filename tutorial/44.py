# 대신증권 API
# ETF NAV , ETN IIV 실시간 수신 예제

# 이번 예제는 ETF NAV 를 실시간으로 수신 또는 ETN 의 IIV 를 실시간으로 수신 하는 예제 코드입니다
# 아래는 사용된 주요 플러스 서비스입니다.
# Dscbo1.Cpsvr7244 - ETF NAV 조회 서비스
# Dscbo1.Cpsvr7718 - ETN IIV 조회 서비스
# CpSysDib.CpSvrNew7244S - NAV/IIV 실시간 수신 서비스
# DsCbo1.StockCur - 현재가 실시간 수신 서비스
# 이번 예제에서는 예제의 다양화를 위해여 조회 통신으로 BlockReqeust API 대신 Request ~ Received 이벤트 처리 방식으로 만들었습니다.
# 조회 종목은 ETF/ETN 각가 A069500, Q500032 종목으로 조회합니다.

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
import time

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

    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False

    return True


################################################
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.cpObj = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

    def OnReceived(self):
        if self.name == 'nav':
            print('실시간 nav 수신')
            self.caller.OnPublish_NAV(self.cpObj)
            return
        elif self.name == 'stockcur':
            print('실시간 현재가 수신')
            self.caller.OnPublish_Cur(self.cpObj)
            return
        elif self.name == 'etfreply':
            print('reply')
            self.caller.OnReply(self.cpObj)
            return


################################################
# plus 실시간 수신 base 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def __del__(self):
        self.Unsubscribe()

    def Subscribe(self, var, caller):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, caller)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


# 실시간 nav/iiv 수신
class CP_PB_NAV(CpPublish):
    def __init__(self):
        super().__init__('nav', 'CpSysDib.CpSvrNew7244S')


# 실시간 현재가 수신
class CP_PB_CUR(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')


# Reply 이벤트 - 1회만 수신
class CpRpETF:
    def __init__(self):
        self.name = "etfreply"

    def SetEvent(self, cpobj, caller):
        handler = win32com.client.WithEvents(cpobj, CpEvent)
        handler.set_params(cpobj, self.name, caller)


class CP_ETF_NAV:
    def __init__(self):
        self.objRq = None
        self.objReply = None
        self.objname1 = 'Dscbo1.Cpsvr7244'
        self.objname2 = 'Dscbo1.Cpsvr7718'
        self.navlist = []
        self.code = ''
        self.caller = None

    def OnReply(self, objRq):
        cnt = objRq.GetHeaderValue(0)
        print('조회 개수', cnt)

        for i in range(cnt):
            item = {}

            item['시간'] = objRq.GetDataValue(0, i)
            item['현재가'] = objRq.GetDataValue(1, i)
            item['대비'] = objRq.GetDataValue(3, i)
            item['거래량'] = objRq.GetDataValue(5, i)
            if (self.code[0] == 'A'):
                item['NAV대비'] = objRq.GetDataValue(4, i)
                item['NAV'] = objRq.GetDataValue(6, i)
            else:
                item['IIV대비'] = objRq.GetDataValue(4, i)
                item['IIV'] = objRq.GetDataValue(6, i)
            item['추적오차'] = objRq.GetDataValue(8, i)
            item['괴리율'] = objRq.GetDataValue(9, i)
            item['해당ETF지수'] = objRq.GetDataValue(10, i)
            item['지수대비'] = objRq.GetDataValue(11, i)
            print(item)
            self.navlist.append(item)

        self.caller.OnReply(self.navlist)

    def Request(self, code, caller):
        self.navlist = []
        self.code = code
        self.caller = caller

        if (self.objRq != None):
            self.objRq = None

        if (code[0] == 'A'):
            self.objRq = win32com.client.Dispatch(self.objname1)
        else:
            self.objRq = win32com.client.Dispatch(self.objname2)

        self.objReply = CpRpETF()
        self.objReply.SetEvent(self.objRq, self)

        self.objRq.SetInputValue(0, code)
        self.objRq.Request()

        return


################################################
# 테스트를 위한 메인 화면
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.navlist = []

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()
        self.rqObj = None
        self.objPB7244 = None
        self.objPBcur = None

        #######################################
        # 윈도우 처리
        self.setWindowTitle("ETF NAV TEST")
        self.setGeometry(300, 300, 300, 220)

        nH = 20
        btnETF = QPushButton('ETF NAV', self)
        btnETF.move(20, nH)
        btnETF.clicked.connect(self.btnETF_clicked)
        nH += 50

        btnETN = QPushButton('ETN NAV', self)
        btnETN.move(20, nH)
        btnETN.clicked.connect(self.btnETN_clicked)
        nH += 50

        btnPrint = QPushButton('print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

    #######################################
    # NAV 실시간 수신
    def OnPublish_NAV(self, obj):
        item = {}
        code = obj.GetHeaderValue(0)
        item['시간'] = obj.GetHeaderValue(1)
        item['현재가'] = obj.GetHeaderValue(2)
        item['대비'] = obj.GetHeaderValue(4)
        item['거래량'] = obj.GetHeaderValue(5)
        if (code[0] == 'A'):
            item['NAV'] = obj.GetHeaderValue(6)
            if item['NAV'] != 0:
                item['NAV'] /= 100
            item['NAV대비'] = item['현재가'] - item['NAV']

        else:
            item['IIV'] = obj.GetHeaderValue(6)
            if item['IIV'] != 0:
                item['IIV'] /= 100
            item['IIV대비'] = item['현재가'] - item['IIV']

        item['추적오차'] = obj.GetHeaderValue(10)
        if item['추적오차'] > 0:
            item['추적오차'] /= 100
        flag = obj.GetHeaderValue(11)
        item['괴리율'] = obj.GetHeaderValue(12)
        if item['괴리율'] != 0:
            item['괴리율'] /= 100
        if (flag == ord('-')):
            item['괴리율'] *= -1
        item['해당ETF지수'] = obj.GetHeaderValue(15)
        if item['해당ETF지수'] > 0:
            item['해당ETF지수'] != 100
        flag = obj.GetHeaderValue(13)
        item['지수대비'] = obj.GetHeaderValue(14)
        if item['지수대비'] != 0:
            item['지수대비'] /= 100
        if (flag == ord('-')):
            item['지수대비'] *= -1
        print(item)

        self.navlist.insert(0, item)

    #######################################
    # 현재가 실시간 수신
    def OnPublish_Cur(self, obj):
        item = {}
        exflag = obj.GetHeaderValue(19)  # 예상체결 플래그
        if exflag != ord('2'):
            item['동시호가여부'] = '예상'
        else:
            item['동시호가여부'] = '정규'
        item['현재가'] = obj.GetHeaderValue(13)  # 현재가
        item['대비'] = obj.GetHeaderValue(2)  # 대비
        item['거래량'] = obj.GetHeaderValue(9)  # 거래량
        print(item)

    # NAV/IIV 시간대별 데이터 조회 수신
    def OnReply(self, data):
        self.navlist = data

        #######################################
        # NAV/IIV 실시간 이벤트 요청
        if (self.objPB7244 != None):
            self.objPB7244 = None

        self.objPB7244 = CP_PB_NAV()
        self.objPB7244.Subscribe(self.code, self)

        #######################################
        # 현재가 실시간 요청
        if (self.objPBcur != None):
            self.objPBcur = None
        self.objPBcur = CP_PB_CUR()
        self.objPBcur.Subscribe(self.code, self)

    def Request(self, code):
        self.code = code
        self.navlist = []

        #######################################
        # NAV/IIV 시간대별 리스트 요청
        if (self.rqObj):
            self.rqObj = None

        self.rqObj = CP_ETF_NAV()
        self.rqObj.Request(code, self)

    # ETF 종목 요청
    def btnETF_clicked(self):
        code = 'A069500'
        self.Request(code)

    # ETN 종목 요청
    def btnETN_clicked(self):
        code = 'Q500032'
        self.Request(code)

    def btnPrint_clicked(self):
        for data in self.navlist:
            print(data)
        return

    def btnExit_clicked(self):
        self.objPB7244.Unsubscribe()
        self.objPBcur.Unsubscribe()
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
# 대신증권 API
# 주식 분할 주문 예제

# 파이썬을 이용하여 간단한 주식 분할 매수 주문 하는 예제입니다

# ■ 예제 화면 주요 설명
# - 입력 사항: 종목코드, 분할 주기(분 단위), 분할 회수
# - 주문 시작: 입력 사항에 따라 10주씩 분할 매수 시작,
#    분할 횟수가 10회, 분할주기가 1분이면 1분마다 매수 주문을 내고, 10회 반복합니다
#
# ■ 주요 클래스
# - CpRPCurrentPrice - DsCbo1.StockMst 를 이용하여 현재가 통신
# - CpRPOrder - CpTrade.CpTd0311 를 이용하여 매수 주문 통신
#
# ※ 주의 사항: 본 예제는 PLUS API 활용을 위한 참고용으로만 제공되며, 이에 대한 활용의 책임은 이용자에게 있습니다.

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QTimer
import win32com.client
import time

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        return

    def Request(self, code, caller):
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        caller.curData = {}

        caller.curData['code'] = code
        caller.curData['종목명'] = g_objCodeMgr.CodeToName(code)
        caller.curData['현재가'] = self.objStockMst.GetHeaderValue(11)  # 종가
        caller.curData['대비'] = self.objStockMst.GetHeaderValue(12)  # 전일대비
        caller.curData['기준가'] = self.objStockMst.GetHeaderValue(27)  # 기준가
        caller.curData['거래량'] = self.objStockMst.GetHeaderValue(18)  # 거래량
        caller.curData['예상플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        caller.curData['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        caller.curData['예상대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결대비

        # 10차호가
        for i in range(10):
            key1 = '매도호가%d' % (i + 1)
            key2 = '매수호가%d' % (i + 1)
            caller.curData[key1] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            caller.curData[key2] = (self.objStockMst.GetDataValue(1, i))  # 매수호가

        print(caller.curData)

        return True


# 주식 주문 처리
class CpRPOrder:
    def __init__(self):
        # 연결 여부 체크
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return

        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])

        self.objBuyOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 매수
        self.orderNum = 0  # 주문 번호

    def buyOrder(self, code, price, amount, caller):
        # 주식 매수 주문
        print("신규 매수", code, price, amount)

        self.objBuyOrder.SetInputValue(0, "2")  # 2: 매수
        self.objBuyOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objBuyOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objBuyOrder.SetInputValue(3, code)  # 종목코드
        self.objBuyOrder.SetInputValue(4, amount)  # 매수수량
        self.objBuyOrder.SetInputValue(5, price)  # 주문단가
        self.objBuyOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objBuyOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

        # 매수 주문 요청
        self.objBuyOrder.BlockRequest()

        rqStatus = self.objBuyOrder.GetDibStatus()
        rqRet = self.objBuyOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        orderdata = {}

        now = time.localtime()
        sTime = "%04d-%02d-%02d %02d:%02d:%02d" % (
        now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
        orderdata['주문시간'] = sTime
        orderdata['주문종류'] = self.objBuyOrder.GetHeaderValue(0)
        orderdata['종목코드'] = self.objBuyOrder.GetHeaderValue(3)
        orderdata['주문수량'] = self.objBuyOrder.GetHeaderValue(4)
        orderdata['주문단가'] = self.objBuyOrder.GetHeaderValue(5)
        orderdata['주문번호'] = self.objBuyOrder.GetHeaderValue(8)
        # orderdata['주문조건구분코드'] = self.objBuyOrder.GetHeaderValue(12)
        # orderdata['주문호가구분코드'] = self.objBuyOrder.GetHeaderValue(13)

        caller.orderData.append(orderdata)
        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # plus 주문 초기화
        if (g_objCpTrade.TradeInit(0) != 0):
            print("주문 초기화 실패")
            return

        # 현재가 정보
        self.curData = {}
        self.objCur = CpRPCurrentPrice()

        # 주문 데이터
        self.orderData = []
        self.objOrder = CpRPOrder()

        # 기본 값
        self.interval = 1  # 분할 주기
        self.code = 'A003540'  # 주문 종목
        self.count = 10  # 분할 횟수
        self.remaincount = 0  # 남은 주문 횟수

        # 타이머
        self.timer = None

        # 윈도우 버튼 배치
        self.setWindowTitle("분할주문 테스트")
        nH = 20

        # 종목코드
        self.label = QLabel('종목코드', self)
        self.label.move(140, nH)
        self.codeEdit = QLineEdit("", self)
        self.codeEdit.move(20, nH)
        self.codeEdit.textChanged.connect(self.codeEditChanged)
        self.codeEdit.setText('003540')
        nH += 50

        # 분할 간격
        self.editInterval = QLineEdit("", self)
        self.editInterval.move(20, nH)
        self.editInterval.textChanged.connect(self.intervalEditChanged)
        self.editInterval.setText(str(self.interval))
        self.labelInterval = QLabel('분할주기(분)', self)
        self.labelInterval.move(140, nH)
        nH += 50

        # 분할 횟수
        self.editCount = QLineEdit("", self)
        self.editCount.move(20, nH)
        self.editCount.textChanged.connect(self.countEditChanged)
        self.editCount.setText(str(self.count))
        self.labelCount = QLabel('분할회수', self)
        self.labelCount.move(140, nH)
        nH += 50

        # 분할 주문 시작
        self.bntStartOrder = QPushButton("주문시작", self)
        self.bntStartOrder.move(20, nH)
        self.bntStartOrder.clicked.connect(self.bntStartOrder_clicked)
        nH += 50

        # 분할 주문 중지
        self.bntStopOrder = QPushButton("주문중지", self)
        self.bntStopOrder.move(20, nH)
        self.bntStopOrder.clicked.connect(self.bntStopOrder_clicked)
        nH += 50

        btnExit = QPushButton("종료", self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        self.setGeometry(300, 300, 300, nH)
        self.setCode('003540')

    # 분할 주문 시작
    def bntStartOrder_clicked(self):
        # 종목 코드 체크
        name = g_objCodeMgr.CodeToName(self.code)
        if len(name) == 0:
            w = QWidget()
            QMessageBox.warning(w, '종목코드 확인', '종목코드를 확인해 주세요')
            return

        self.bntStartOrder.setEnabled(False)

        # 주문 시작
        self.remaincount = self.count
        self.setCDOrder()

    # 분할 주문 중지
    def bntStopOrder_clicked(self):
        if (self.timer):
            self.timer.stop()
            self.timer.deleteLater()
        self.remaincount = 0
        self.bntStartOrder.setEnabled(True)

    # 종목코드 입력기_ 변경 이벤트
    def codeEditChanged(self):
        code = self.codeEdit.text()
        self.setCode(code)

    # 분할 주기 입력기 _변경 이벤트
    def intervalEditChanged(self):
        self.interval = int(self.editInterval.text())

    # 분할 횟수 입력기 _변경 이벤트
    def countEditChanged(self):
        self.count = int(self.editCount.text())

    def setCode(self, code):
        if len(code) < 6:
            return

        print(code)
        if not (code[0] == "A"):
            code = "A" + code

        name = g_objCodeMgr.CodeToName(code)
        if len(name) == 0:
            print("종목코드 확인")
            return

        self.label.setText(name)
        self.code = code

    def setCDOrder(self):
        if (self.timer):
            self.timer.stop()
            self.timer.deleteLater()
        # 현재가 통신
        if (self.objCur.Request(self.code, self) == False):
            w = QWidget()
            QMessageBox.warning(w, '오류', '현재가 통신 오류 발생/주문 중단')
            self.bntStartOrder.setEnabled(True)
            return

        # 주문 전송
        if (self.objOrder.buyOrder(self.code, self.curData['현재가'], 10, self) == False):
            w = QWidget()
            QMessageBox.warning(w, '오류', '매수주문 오류/주문중단')
            self.bntStartOrder.setEnabled(True)
            return

        self.remaincount -= 1
        print('남은 주문 횟수: ', self.remaincount)
        for data in self.orderData:
            print(data)

        if self.remaincount <= 0:
            w = QWidget()
            QMessageBox.information(w, '완료', '분할 주문 완료')
            self.bntStartOrder.setEnabled(True)
            return

        # 다음 timer 작동
        self.timer = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.setCDOrder)
        self.timer.start(self.interval * 60 * 1000)

    def btnExit_clicked(self):
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
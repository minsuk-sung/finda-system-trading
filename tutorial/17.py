# 대신증권 API
# 주식 예약 매수/매도/취소/조회 예제

# 주식 예약 주문 예제 입니다
# 제공된 샘플은 단순 참고용으로 제공되는 예제임으로
# 자세한 기능은 충분히 모의투자에서 테스트 해 보신 이후 검증하시고 운영에 적용하시기 바랍니다

# ■ 화면 설명
#   - 예약매수 전송 - 예약 매수(오늘 종가 기준)
#   - 예약매도 전송 - 예약 매수(오늘 종가 기준)
#   - 예약주문 취소 - 미체결 예약주문 리스트에서 첫번째 예약 주문 취소
#   - 예약내역 가져오기 - 예약주문 리스트 가져와 미체결 건만 리스트에 저장(취소 주문 할 수 있도록)

# 예약매수/매도/예약주문 내역 조회 예제
# 화면 설명
#   예약매수 전송 - 예약 매수(오늘 종가 기준)
#   예약매도 전송 - 예약 매수(오늘 종가 기준)
#   예약주문 취소 - 미체결 예약주문 리스트에서 첫번째 예약 주문 취소
#   예약내역 가져오기 - 예약주문 리스트 가져와 미체결 건만 리스트에 저장(취소 주문 할 수 있도록)
import sys
from PyQt5.QtWidgets import *
from enum import Enum
import win32com.client
import time
import pythoncom

g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")


# enum 주문 상태 세팅용
class EorderBS(Enum):
    buy = 1  # 매수
    sell = 2  # 매도
    none = 3


# 현재가 정보 저장 구조체
class stockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "동시호가/장중 아님", ord('1'): "동시호가", ord('2'): "장중"}
        self.code = ""
        self.name = ""
        self.cur = 0  # 현재가
        self.open = self.high = self.low = 0  # 시/고/저
        self.diff = 0
        self.diffp = 0
        self.objCur = None
        self.objBid = None
        self.vol = 0  # 거래량
        self.offer = [0 for _ in range(10)]  # 매도호가
        self.bid = [0 for _ in range(10)]  # 매수호가
        self.offervol = [0 for _ in range(10)]  # 매도호가 잔량
        self.bidvol = [0 for _ in range(10)]  # 매수호가 잔량

    # 전일 대비 계산
    def makediffp(self, baseprice):
        lastday = 0
        if baseprice:
            lastday = baseprice
        else:
            lastday = self.cur - self.diff
        if lastday:
            self.diffp = (self.diff / lastday) * 100
        else:
            self.diffp = 0

    def debugPrint(self, type):
        if type == 0:
            print("%s, %s %s, 현재가 %d 대비 %d, (%.2f), 1차매도 %d(%d) 1차매수 %d(%d)"
                  % (self.dicEx.get(self.exFlag), self.code,
                     self.name, self.cur, self.diff, self.diffp,
                     self.offer[0], self.offervol[0], self.bid[0], self.bidvol[0]))
        else:
            print("%s %s, 현재가 %.2f 대비 %.2f, (%.2f), 1차매도 %.2f(%d) 1차매수 %.2f(%d)"
                  % (self.code,
                     self.name, self.cur, self.diff, self.diffp,
                     self.offer[0], self.offervol[0], self.bid[0], self.bidvol[0]))


# 주문 데이터
class orderData:
    def __init__(self):
        self.dicEx = {EorderBS.buy: "매수", EorderBS.sell: "매도", EorderBS.none: "없음"}
        self.orderNum = 0
        self.bs = EorderBS.none  # 0 : buy 1: sell
        self.code = ""
        self.amount = 0
        self.price = 0

    def debugPrint(self):
        print(self.dicEx.get(self.bs), self.code, self.orderNum, self.amount, self.price)


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
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

        # 수신 받은 현재가 정보를 rtMst 에 저장
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur = self.objStockMst.GetHeaderValue(11)  # 종가
        rtMst.diff = self.objStockMst.GetHeaderValue(12)  # 전일대비
        rtMst.baseprice = self.objStockMst.GetHeaderValue(27)  # 기준가
        rtMst.exFlag = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        if rtMst.baseprice:
            rtMst.diffp = (rtMst.diff / rtMst.baseprice) * 100

        # 10차호가
        for i in range(10):
            rtMst.offer[i] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            rtMst.bid[i] = (self.objStockMst.GetDataValue(1, i))  # 매수호가
            rtMst.offervol[i] = (self.objStockMst.GetDataValue(2, i))  # 매도호가 잔량
            rtMst.bidvol[i] = (self.objStockMst.GetDataValue(3, i))  # 매수호가 잔량
        return True


# CpRPPreOrder - 예약 주문 매수/매도/취소/내역 조회 클래스
class CpRPPreOrder:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        if (g_objCpTrade.TradeInit(0) != 0):
            print("주문 초기화 실패")
            exit
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTdNew9061")
        self.objCancel = win32com.client.Dispatch("CpTrade.CpTdNew9064")
        self.objResult = win32com.client.Dispatch("CpTrade.CpTd9065")

        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분

        return

    # 예약 매수 또는 매도
    def RequestOrder(self, bs, code, price, amount, data):
        data.code = code
        data.amount = amount
        data.price = price;
        rqBS = "2"
        if bs == EorderBS.buy:  # 매수
            rqBS = "2"
            data.bs = EorderBS.buy
        elif bs == EorderBS.sell:  # 매도
            data.bs = EorderBS.sell
            rqBS = "1"

        # 예약 주문
        print(self.acc, self.accFlag[0])
        self.objOrder.SetInputValue(0, self.acc)
        self.objOrder.SetInputValue(1, self.accFlag[0])
        self.objOrder.SetInputValue(2, rqBS)
        self.objOrder.SetInputValue(3, code)
        self.objOrder.SetInputValue(4, amount)
        self.objOrder.SetInputValue(5, "01")  # 주문호가 구분 01: 보통
        self.objOrder.SetInputValue(6, price)

        # 예약 주문 요청
        self.objOrder.BlockRequest()

        if self.objOrder.GetDibStatus() != 0:
            print("통신상태", self.objOrder.GetDibStatus(), self.objOrder.GetDibMsg1())
            return False

        data.orderNum = self.objOrder.GetHeaderValue(0)  # 예약번호
        print("예약주문 성공, 예약번호 #", data.orderNum)

    # 예약 취소 주문
    def RequestCancel(self, ordernum, code):
        # 예약주문 취소
        print(self.acc, self.accFlag[0])
        self.objCancel.SetInputValue(0, ordernum)
        self.objCancel.SetInputValue(1, self.acc)
        self.objCancel.SetInputValue(2, self.accFlag[0])
        self.objCancel.SetInputValue(3, code)

        # 예약 취소  주문 요청
        self.objCancel.BlockRequest()

        if self.objCancel.GetDibStatus() != 0:
            print("통신상태", self.objCancel.GetDibStatus(), self.objCancel.GetDibMsg1())
            return False

        print("예약주문 취소 ", ordernum, self.objCancel.GetDibMsg1())

    # 예약 주문 내역 조회 및 미체결 리스트 구하기
    def RequestOrderList(self, orderList):
        print(self.acc, self.accFlag[0])
        self.objResult.SetInputValue(0, self.acc)
        self.objResult.SetInputValue(1, self.accFlag[0])
        self.objResult.SetInputValue(2, 20)

        while True:  # 연속 조회로 전체 예약 주문 가져온다.
            self.objResult.BlockRequest()
            if self.objResult.GetDibStatus() != 0:
                print("통신상태", self.objResult.GetDibStatus(), self.objResult.GetDibMsg1())
                return False

            cnt = self.objResult.GetHeaderValue(4)
            if cnt == 0:
                break

            for i in range(cnt):
                i1 = self.objResult.GetDataValue(1, i)  # 주문구분(매수 또는 매도)
                i2 = self.objResult.GetDataValue(2, i)  # 코드
                i3 = self.objResult.GetDataValue(3, i)  # 주문 수량
                i4 = self.objResult.GetDataValue(4, i)  # 주문호가구분
                i5 = self.objResult.GetDataValue(6, i)  # 예약번호
                i6 = self.objResult.GetDataValue(12, i)  # 처리구분내용 - 주문취소 또는 주문예정
                i7 = self.objResult.GetDataValue(9, i)  # 주문단가
                i8 = self.objResult.GetDataValue(11, i)  # 주문번호
                i9 = self.objResult.GetDataValue(12, i)  # 처리구분코드
                i10 = self.objResult.GetDataValue(13, i)  # 거부코드
                i11 = self.objResult.GetDataValue(14, i)  # 거부내용
                print(i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11)

                # 미체결
                if (i6 == "주문예정"):
                    item = orderData()
                    item.orderNum = i5
                    if (i1 == "매수"):
                        item.bs = EorderBS.buy
                    else:
                        item.bs = EorderBS.sell
                    item.code = i2
                    item.amount = i3
                    item.price = i7

                    orderList.append(item)

            # 연속 처리 체크 - 다음 데이터가 없으면 중지
            if self.objResult.Continue == False:
                break


# 샘플 코드  메인 클래스
class testMain():
    def __init__(self):
        self.orderList = []
        self.objMst = CpRPCurrentPrice()
        self.obj9061 = CpRPPreOrder()

        # 미체결 된 예약 주문 처음부터 받아옴.
        self.resultOrder()
        return

    def newBuyOrder(self, code, amount):
        mstData = stockPricedData()
        if self.objMst.Request(code, mstData) == False:
            print("현재가 요청 실패")
            return

        item = orderData()
        ret = self.obj9061.RequestOrder(EorderBS.buy, code, mstData.cur, amount, item)
        if (ret == False):
            return False
        self.orderList.append(item)
        item.debugPrint()

    def newSellOrder(self, code, amount):
        mstData = stockPricedData()
        if self.objMst.Request(code, mstData) == False:
            print("현재가 요청 실패")
            return

        item = orderData()
        ret = self.obj9061.RequestOrder(EorderBS.sell, code, mstData.cur, amount, item)
        if (ret == False):
            return False
        self.orderList.append(item)
        item.debugPrint()

    def cancelOrder(self):
        if len(self.orderList) == 0:
            print("취소 할 주문 확인하세요")
            return
        item = self.orderList[0]
        ret = self.obj9061.RequestCancel(item.orderNum, item.code)
        if (ret == False):
            return False

        del (self.orderList[0])

    def resultOrder(self):
        self.orderList = []
        # 미체결 주문만 받아 보자
        ret = self.obj9061.RequestOrderList(self.orderList)
        if (ret == False):
            return False


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.main = testMain()
        self.setWindowTitle("PLUS API TEST")
        w = 200
        h = 30

        nHeight = 20
        btnBuy = QPushButton("예약매수 전송", self)
        btnBuy.move(20, nHeight)
        btnBuy.resize(w, h)
        btnBuy.clicked.connect(self.btnBuy_clicked)

        nHeight += 50
        btnSell = QPushButton("예약 매도 전송", self)
        btnSell.move(20, nHeight)
        btnSell.resize(w, h)
        btnSell.clicked.connect(self.btnSell_clicked)

        nHeight += 50
        btnCancel = QPushButton("예약 주문 취소", self)
        btnCancel.move(20, nHeight)
        btnCancel.resize(w, h)
        btnCancel.clicked.connect(self.btnCancel_clicked)

        nHeight += 50
        btnResult = QPushButton("예약 내역 가져오기", self)
        btnResult.move(20, nHeight)
        btnResult.resize(w, h)
        btnResult.clicked.connect(self.btnResult_clicked)

        nHeight += 50
        btnExit = QPushButton("종료", self)
        btnExit.move(20, nHeight)
        btnExit.resize(w, h)
        btnExit.clicked.connect(self.btnExit_clicked)

        nHeight += 50
        self.setGeometry(300, 500, 300, nHeight)

    def btnBuy_clicked(self):
        self.main.newBuyOrder("A003540", 10)
        return

    def btnSell_clicked(self):
        self.main.newSellOrder("A003540", 10)
        return

    def btnCancel_clicked(self):
        self.main.cancelOrder()
        return

    def btnResult_clicked(self):
        self.main.resultOrder()
        return

    def btnExit_clicked(self):
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
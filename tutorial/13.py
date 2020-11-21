# 대신증권 API
# 주식 주문 체결 실시간 처리 예제
# 주문 체결 데이터를 받아 처리 하는 샘플 코드입니다.
#
#  설명:
#    주식 한 종목의 매수/정정/취소 주문 처리 및 실시간 시세와 주문 체결 처리 예제
#    > 매수주문 - 현재가/10차 호가를 구해 10차 호가로 매수 주문 냄
#    > 정정주문 - 누를 때 마다 호가를 9차 > 8차 > 7차 매수호가식으로 올려 정정주문 (가격은 실시간으로 업데이트 된 가격임)
#    > 취소주문 - 취소 주문
#
#  CpEvent: 실시간 현재가 수신 클래스 - 아래 3가지 실시간 시세 수신
#        실시간 체결 현재가
#        실시간 10차 호가
#        실시간 주문 체결
#
#  CpPBStockCur : (실시간)현재가 체결 요청 클래스
#  CpPBStockBid : (실시간)현재가 10차 호가 요청 클래스
#  CpPBConclusion : (실시간)주문체결 데이터 요청 클래스
#  CpRPOrder : (RQ/RP)주식 매수/매도/정정 통신 클래스
#  CpRPCurrentPrice : (RQ/RP)주식 현재가 통신 클래스
#  OrderMain : 주문/체결에 대한 핵심 처리 클래스
#        매수/정정/취소 주문 버튼 클릭에 대한 이벤트 처리
#        실시간 주문 체결 업데이트에 따른 주문 상태 업데이트

import sys
from PyQt5.QtWidgets import *
from enum import Enum
import win32com.client


# 설명: 주식 한 종목의 매수/정정/취소 주문 처리 및 실시간 시세와 주문 체결 처리 예제
#   매수주문 - 현재가/10차 호가를 구해 10차 호가로 매수 주문 냄
#   정정주문 - 누를 때 마다 호가를 9차 > 8차 > 7차 식으로 올려 정정주문 (가격은 실시간으로 업데이트 된 가격임)
#   취소주문 - 취소 주문

# CpEvent: 실시간 현재가 수신 클래스 - 아래 3가지 실시간 시세 수신
#       실시간 체결 현재가
#       실시간 10차 호가
#       실시간 주문 체결
# CpPBStockCur : (실시간)현재가 체결 요청 클래스
# CpPBStockBid : (실시간)현재가 10차 호가 요청 클래스
# CpPBConclusion : (실시간)주문체결 데이터 요청 클래스
# CpRPOrder : (RQ/RP)주식 매수/매도/정정 통신 클래스
# CpRPCurrentPrice : (RQ/RP)주식 현재가 통신 클래스
# OrderMain : 주문/체결에 대한 핵심 처리 클래스
#       매수/정정/취소 주문 버튼 클릭에 대한 이벤트 처리
#       실시간 주문 체결 업데이트에 따른 주문 상태 업데이트


# enum 주문 상태 세팅용
class orderStatus(Enum):
    nothing = 1  # 별 일 없는 상태
    newOrder = 2  # 신규 주문 낸 상태
    orderConfirm = 3  # 신규 주문 처리 확인
    modifyOrder = 4  # 정정 주문 낸 상태
    cancelOrder = 5  # 취소 주문 낸 상태


# 현재가와 10차 호가를 저장하기 위한 단순 저장소
class stockPricedData:
    def __init__(self):
        self.cur = 0  # 현재가
        self.offer = []  # 매도호가
        self.bid = []  # 매수호가


# 주문 체결 pb 기록용(종료 시 받은 데이터 print)
class orderHistoryData:
    def __init__(self):
        self.flag = ""
        self.code = ""
        self.price = 0
        self.orderamount = 0
        self.contamount = 0
        self.etc = ""

    def sethistory(self, flag, code, price, amount, contamount, ordernum, etc):
        self.flag = flag
        self.code = code
        self.price = price
        self.orderamount = amount
        self.contamount = contamount
        self.ordernum = ordernum
        self.etc = etc

    def printhistory(self):
        print(self.flag, self.code, "가격:", self.price, "수량:", self.orderamount, "체결수량:", self.contamount, "주문번호:",
              self.ordernum, self.etc)


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, parent):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.parent = parent  # callback 을 위해 보관

        # 데이터 변환용
        self.concdic = {"1": "체결", "2": "확인", "3": "거부", "4": "접수"}
        self.buyselldic = {"1": "매도", "2": "매수"}
        print(self.concdic)
        print(self.buyselldic)

    # PLUS 로 부터 실제로 시세를 수신 받는 이벤트 핸들러
    def OnReceived(self):
        print(self.name)
        if self.name == "stockcur":
            # 현재가 체결 데이터 실시간 업데이트
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            # 장중이 아니면 처리 안함.
            if exFlag != ord('2'):
                return

            # 현재가 업데이트
            self.parent.sprice.cur = cprice
            print("PB > 현재가 업데이트 : ", cprice)

            # 현재가 변경  call back 함수 호출
            self.parent.monitorPriceChange()

            return

        elif self.name == "stockbid":
            # 현재가 10차 호가 데이터 실시간 업데이트
            dataindex = [3, 4, 7, 8, 11, 12, 15, 16, 19, 20, 27, 28, 31, 32, 35, 36, 39, 40, 43, 44]
            obi = 0
            for i in range(10):
                self.parent.sprice.offer[i] = self.client.GetHeaderValue(dataindex[obi])
                self.parent.sprice.bid[i] = self.client.GetHeaderValue(dataindex[obi + 1])
                obi += 2

            # for debug
            for i in range(10):
                print("PB > 10차 호가 : ", i + 1, "차 매도/매수 호가: ", self.parent.sprice.offer[i], self.parent.sprice.bid[i])
            return True

            # 10차 호가 변경 call back 함수 호출
            self.parent.monitorPriceChange()

            return

        elif self.name == "conclution":
            # 주문 체결 실시간 업데이트
            conflag = self.client.GetHeaderValue(14)  # 체결 플래그
            ordernum = self.client.GetHeaderValue(5)  # 주문번호
            amount = self.client.GetHeaderValue(3)  # 체결 수량
            price = self.client.GetHeaderValue(4)  # 가격
            code = self.client.GetHeaderValue(9)  # 종목코드
            bs = self.client.GetHeaderValue(12)  # 매수/매도 구분
            balace = self.client.GetHeaderValue(23)  # 체결 후 잔고 수량

            conflags = ""
            if conflag in self.concdic:
                conflags = self.concdic.get(conflag)
                print(conflags)

            bss = ""
            if (bs in self.buyselldic):
                bss = self.buyselldic.get(bs)

            print(conflags, bss, code, "주문번호:", ordernum)
            # call back 함수 호출해서 orderMain 에서 후속 처리 하게 한다.
            self.parent.monitorOrderStatus(code, ordernum, conflags, price, amount, balace)
            return


# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur:
    def __init__(self):
        self.name = "stockcur"
        self.obj = win32com.client.Dispatch("DsCbo1.StockCur")

    def Subscribe(self, code, sprice, parent):
        self.obj.SetInputValue(0, code)
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, parent)
        self.obj.Subscribe()
        self.sprice = sprice

    def Unsubscribe(self):
        self.obj.Unsubscribe()


# CpPBStockBid: 실시간 10차 호가 요청 클래스
class CpPBStockBid:
    def __init__(self):
        self.name = "stockbid"
        self.obj = win32com.client.Dispatch("Dscbo1.StockJpBid")

    def Subscribe(self, code, sprice, parent):
        self.obj.SetInputValue(0, code)
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, parent)
        self.obj.Subscribe()
        self.sprice = sprice

    def Unsubscribe(self):
        self.obj.Unsubscribe()


# CpPBConclusion: 실시간 주문 체결 수신 클래그
class CpPBConclusion:
    def __init__(self):
        self.name = "conclution"
        self.obj = win32com.client.Dispatch("DsCbo1.CpConclusion")

    def Subscribe(self, parent):
        self.parent = parent
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, parent)
        self.obj.Subscribe()

    def Unsubscribe(self):
        self.obj.Unsubscribe()


class CpRPOrder:
    def __init__(self):
        # 연결 여부 체크
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return

        # 주문 초기화
        self.objTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
        initCheck = self.objTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            return

        self.acc = self.objTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = self.objTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])

        # 매수/정정/취소 주문 object 생성
        self.objBuyOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 매수
        self.objModifyOrder = win32com.client.Dispatch("CpTrade.CpTd0313")  # 정정
        self.objCancelOrder = win32com.client.Dispatch("CpTrade.CpTd0314")  # 취소
        self.orderNum = 0  # 주문 번호

    def buyOrder(self, code, price, amount):
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

        # 주의: 매수 주문에  대한 구체적인 처리는 cpconclution 으로 파악해야 한다.
        return True

    def modifyOrder(self, ordernum, code, price):
        # 주식 정정 주문
        print("정정주문", ordernum, code, price)
        self.objModifyOrder.SetInputValue(1, ordernum)  # 원주문 번호 - 정정을 하려는 주문 번호
        self.objModifyOrder.SetInputValue(2, self.acc)  # 상품구분 - 주식 상품 중 첫번째
        self.objModifyOrder.SetInputValue(3, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objModifyOrder.SetInputValue(4, code)  # 종목코드
        self.objModifyOrder.SetInputValue(5, 0)  # 정정 수량, 0 이면 잔량 정정임
        self.objModifyOrder.SetInputValue(6, price)  # 정정주문단가

        # 정정주문 요청
        self.objModifyOrder.BlockRequest()

        rqStatus = self.objModifyOrder.GetDibStatus()
        rqRet = self.objModifyOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        # 새로운 주문 번호 구한다.
        self.orderNum = self.objModifyOrder.GetHeaderValue(7)

    def cancelOrder(self, ordernum, code):
        # 주식 취소 주문
        print("취소주문", ordernum, code)
        self.objCancelOrder.SetInputValue(1, ordernum)  # 원주문 번호 - 정정을 하려는 주문 번호
        self.objCancelOrder.SetInputValue(2, self.acc)  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(3, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(4, code)  # 종목코드
        self.objCancelOrder.SetInputValue(5, 0)  # 정정 수량, 0 이면 잔량 취소임

        # 취소주문 요청
        self.objCancelOrder.BlockRequest()

        rqStatus = self.objCancelOrder.GetDibStatus()
        rqRet = self.objCancelOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False


# CpRPCurrentPrice : 주식 현재가 및 10차 호가 조회
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
        # 10차호가
        for i in range(10):
            rtMst.offer.append(self.objStockjpbid.GetDataValue(0, i))  # 매도호가
            rtMst.bid.append(self.objStockjpbid.GetDataValue(1, i))  # 매수호가

        # for debug
        for i in range(10):
            print(i + 1, "차 매도/매수 호가: ", rtMst.offer[i], rtMst.bid[i])
        return True


# 주문 테스트용 클래스
class OrderMain():
    def __init__(self):
        self.isSB = False  # 실시간 처리
        self.initOrder()  # 주문 상태 - 초기화

        self.sprice = stockPricedData()  # 주문 현재가/10차 호가 저장 (실시간 업데이트)
        self.cporder = CpRPOrder()  # 주문 통신 object

        # 실시간 통신 object
        self.cur = CpPBStockCur()
        self.bid = CpPBStockBid()

        # 주문체결은 미리 실시간 요청
        self.conclution = CpPBConclusion()
        self.conclution.Subscribe(self)

        self.history = []

    def stopSubscribe(self):
        if self.isSB:
            self.cur.Unsubscribe()
            self.bid.Unsubscribe()

        self.isSB = False

    # 매수 주문
    def BuyOrder(self):
        self.stopSubscribe()
        self.code = "A003540"  # 테스트용 종목코드
        self.buyamount = 1  # 주문 수량

        # 1. 현재가 구하기
        price = CpRPCurrentPrice()
        if price.Request(self.code, self.sprice) == False:
            print("현재가 통신 실패")
            self.initOrder()
            return

        print("신규 매수주문 - ", self.orderNonce + 1, "차 매수호가 ", + self.sprice.bid[self.orderNonce])
        bResult = self.cporder.buyOrder(self.code, self.sprice.bid[self.orderNonce], self.buyamount)
        if bResult == False:
            print("주문 실패")
            self.initOrder()
            return
        self.orderStatus = orderStatus.newOrder  # 주문상태 업데이트

        # 실시간 통신 요청
        self.cur.Subscribe(self.code, self.sprice, self)
        self.bid.Subscribe(self.code, self.sprice, self)
        self.isSB = True

    # 정정주문
    def ModifyOrder(self):
        if not (self.orderStatus == orderStatus.orderConfirm):
            print("정정주문 확인 불가 상태 ")
            return

        if self.ordernum == 0:
            print("주문 번호가 없습니다")
            return

        # 정정주문 할 때 마다 1 호가씩 올린다.
        self.orderNonce -= 1
        if self.orderNonce <= 0:
            self.orderNonce = 0
        print("정정 주문 - ", self.orderNonce + 1, "차 매수호가 ", + self.sprice.bid[self.orderNonce])
        bResult = self.cporder.modifyOrder(self.ordernum, self.code, self.sprice.bid[self.orderNonce])
        if bResult == False:
            print("정정 주문 실패")
            return

        # 주문상태 업데이트
        self.orderStatus = orderStatus.modifyOrder

        # 정정주문은 거래소에서 거부 당할 수 있어 확인/거부 여부를 반드시 확인 해야 함.

        return

    # 취소주문
    def CancelOrder(self):
        if not (self.orderStatus == orderStatus.orderConfirm):
            print("취소주문 확인 불가 상태 ")
            return

        if self.ordernum == 0:
            print("주문 번호가 없습니다")
            return
        # 취소 주문
        bResult = self.cporder.cancelOrder(self.ordernum, self.code)
        if bResult == False:
            print("취소 주문 실패")
            return

        self.orderStatus = orderStatus.cancelOrder  # 주문상태 업데이트
        # 취소주문은 거래소에서 거부 당할 수 있어 확인/거부 여부를 반드시 확인 해야 함.

        return

    # 전체 클리어
    def clearAll(self):
        self.initOrder()
        self.stopSubscribe()
        self.conclution.Unsubscribe()

        # debug
        if (len(self.history)):
            print("주문 내역 정리 ============================")
            for i in range(0, len(self.history)):
                self.history[i].printhistory()

        self.history = []

        return

    def initOrder(self):
        # 주문 정보 초기화
        self.orderStatus = orderStatus.nothing
        self.ordernum = 0  # 주문번호
        self.remainAmount = 0  # 주문 후 미체결 수량
        self.orderNonce = 9  # 매수 주문 호가 조정 변수 ( 9 > 8 > 7 .. 순으로 호가 조정)

    def monitorPriceChange(self):
        # 이곳에서 시세 변경에 대한 감시 등의 로직 추가고려

        return

    # 실시간 주문 체결 업데이트
    def monitorOrderStatus(self, code, ordernum, conflags, price, amount, balance):
        print("주문체결: ", code, ordernum, conflags, price, amount, balance)
        if self.orderStatus == orderStatus.nothing:
            return
        # 체결: 체결 시 체결 수량/미체결 수량 계산
        if conflags == "체결":
            self.remainAmount -= amount  # 미체결 수량 계산
            if self.orderStatus == orderStatus.orderConfirm:
                print("주문 체결 됨 ", "수량", amount, "잔고수량:", balance, "미체결수량:", self.remainAmount)

            if self.remainAmount <= 0:  # 전량 체결 됨
                self.initOrder()

            # for debug
            history = orderHistoryData()
            history.sethistory(conflags, code, price, self.remainAmount, amount, ordernum, "")
            self.history.append(history)


        #  접수: 신규 주문 > 접수 ;--> 주문번호, 주문 정상 처리
        elif conflags == "접수":
            if self.orderStatus == orderStatus.newOrder:
                self.ordernum = ordernum  # 주문번호 업데이트
                self.remainAmount = amount  # 주문 후 미체결 수량
                self.orderStatus = orderStatus.orderConfirm

                # for debug
                history = orderHistoryData()
                history.sethistory(conflags, code, price, amount, 0, ordernum, "신규 매수")
                self.history.append(history)
                history.printhistory()

        #  확인: 정정/취소 주문 > 확인 ;--> 정정/취소 주문 정상 처리 확인
        elif conflags == "확인":
            etc = ""
            if self.orderStatus == orderStatus.modifyOrder:  # 정정 확인
                self.ordernum = ordernum  # 주문번호 업데이트
                self.orderStatus = orderStatus.orderConfirm
                etc = "정정확인"
            elif self.orderStatus == orderStatus.cancelOrder:  # 취소 확인
                self.initOrder()
                etc = "취소확인"

            # for debug
            history = orderHistoryData()
            print(code, price)
            print(self.remainAmount, ordernum)
            history.sethistory(conflags, code, price, self.remainAmount, 0, ordernum, etc)
            self.history.append(history)
            history.printhistory()


        # 거부: 정정/취소 주문 > 거부 ;--> 정정/취소 주문 거부, 정정/취소 불가
        elif conflags == "거부":
            if self.orderStatus == orderStatus.modifyOrder or self.orderStatus == orderStatus.cancelOrder:
                print("주문거부 발생, 반드시 확인 필요")
                self.orderStatus = orderStatus.newOrder  # 주문 상태를 이전으로 돌림

            # for debug
            history = orderHistoryData()
            history.sethistory(conflags, code, price, amount, 0, ordernum, "")
            self.history.append(history)
            history.printhistory()

        return


class MyWindow(QMainWindow):

    def __init__(self):
        self.orerMain = OrderMain()
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 230)

        btnBuy = QPushButton("매수주문", self)
        btnBuy.move(20, 20)
        btnBuy.resize(200, 30)
        btnBuy.clicked.connect(self.btnBuy_clicked)

        btnModify = QPushButton("정정주문(1호가씩 올림)", self)
        btnModify.move(20, 70)
        btnModify.resize(200, 30)
        btnModify.clicked.connect(self.btnModify_clicked)

        btnCancel = QPushButton("취소주문", self)
        btnCancel.move(20, 120)
        btnCancel.resize(200, 30)
        btnCancel.clicked.connect(self.btnCancel_clicked)

        btnExit = QPushButton("종료", self)
        btnExit.move(20, 170)
        btnExit.resize(200, 30)
        btnExit.clicked.connect(self.btnExit_clicked)

    # 매수 주문
    def btnBuy_clicked(self):
        self.orerMain.BuyOrder()

    # 정정주문
    def btnModify_clicked(self):
        self.orerMain.ModifyOrder()
        return

    # 취소주문
    def btnCancel_clicked(self):
        self.orerMain.CancelOrder()

    # 종료
    def btnExit_clicked(self):
        self.orerMain.clearAll()
        exit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
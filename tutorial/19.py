# 대신증권 API
# 주식 미체결 조회 및 실시간 미체결 업데이트/취소주문/일괄 취소 예제

# 주식 미체결을 실시간으로 처리하는 파이썬 예제 입니다
# ※ 제공된 예제는 PLUS API 학습을 위해 제공되는 예제로, 모든 예외 케이스와 상세 처리가 포함되어 있지 않습니다. 참고로만 이용하시기 바랍니다.
#
# 1. 미체결 (CpTrade.CpTd5339)
#     예제에서는 Cp5339 클래스에서 미체결을 조회 합니다
#     연속 조회를 통해 당일 발생한 모든 미체결을 조회 합니다.
#
# 2. 실시간 주문 체결 (DsCbo1.CpConclusion)
#   미체결을 실시간으로 감시하기 위해서는 "DsCbo1.CpConclusion" 에 이벤트를 등록하고 이벤트를 수신 받아 처리 해야 합니다
#   DsCbo1.CpConclusion 이벤트는 CpEvent::OnReceived 에서 수신 받아 처리 합니다.
#
#   기본적으로 실시간 주문 체결의 경우 아래 4가지 type 이 있습니다
#       ▶ 접수 - 주문에 대한 1차 수신
#       ▶ 확인(정정 또는 취소) - 정정이나 취소 주문에 대해 거래소로 부터 확인 응답을 받음
#       ▶ 체결 - 주문 내역 중 일부 또는 전체가 체결됨을 통보 받음
#       ▶ 거부 - 정정 또는 취소 주문 등이 거래소로 부터 거부 됨(이미 체결 된 경우)
#
# 3. 취소 주문 (CpTrade.CpTd0314)
#   클래스 CpRPOrder 에서는 CpTrade.CpTd0314 를 이용하여 취소주문을 냅니다
#   CpRPOrder 에서는 2가지 취소 방식을 제공하는데
#   RequestCancel 는 Request API 를 BlockRequestCancel 는 BlockRequest API 를 각각 이용하여 취소 주문을 냅니다
#   Request API 의 경우 수신이벤트를 등록하고 처리해야 함으로 이에 대한 예제로 참고 바랍니다.
#
# 4. 일괄 취소
#   일괄 취소는 미체결 된 전체 주문 리스트에 대해 BlockRequest 를 이용하여 취소 주문을 냅니다
#
# 5. 연속 주문에 대한 처리
#   PLUS 는 연속적으로 주문/계좌 조회가 발생 할 경우 서비스 보호를 위해 오류 코드 4를 리턴합니다
#   이 경우에는 CpUtil.CpCybos > LimitRequestRemainTime 를 이용하여 남은 대기 시간동안 대기  후 재 시도 또는 다른 방법을 찾아야 합니다 (신규 주문인 경우 대기 시간 동안 주문 가격이 달라질 수 있으니 이런 점에 유의가 필요 합니다)
#   예제에 취소 주문과 미체결 통신에 연속 조회 오류 코드를 참고 하시기 바랍니다.

import sys
from PyQt5.QtWidgets import *
import win32com.client
import time

g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")


# 미체결 주문 정보 저장 구조체
class orderData:
    def __init__(self):
        self.code = ""  # 종목코드
        self.name = ""  # 종목명
        self.orderNum = 0  # 주문번호
        self.orderPrev = 0  # 원주문번호
        self.orderDesc = ""  # 주문구분내용
        self.amount = 0  # 주문수량
        self.price = 0  # 주문 단가
        self.ContAmount = 0  # 체결수량
        self.credit = ""  # 신용 구분 "현금" "유통융자" "자기융자" "유통대주" "자기대주"
        self.modAvali = 0  # 정정/취소 가능 수량
        self.buysell = ""  # 매매구분 코드  1 매도 2 매수
        self.creditdate = ""  # 대출일
        self.orderFlag = ""  # 주문호가 구분코드
        self.orderFlagDesc = ""  # 주문호가 구분 코드 내용

        # 데이터 변환용
        self.concdic = {"1": "체결", "2": "확인", "3": "거부", "4": "접수"}
        self.buyselldic = {"1": "매도", "2": "매수"}

    def debugPrint(self):
        print("%s, %s, 주문번호 %d, 원주문 %d, %s, 주문수량 %d, 주문단가 %d, 체결수량 %d, %s, "
              "정정가능수량 %d, 매수매도: %s, 대출일 %s, 주문호가구분 %s %s"
              % (self.code, self.name, self.orderNum, self.orderPrev, self.orderDesc, self.amount, self.price,
                 self.ContAmount, self.credit, self.modAvali, self.buyselldic.get(self.buysell),
                 self.creditdate, self.orderFlag, self.orderFlagDesc))


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, parent):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.parent = parent  # callback 을 위해 보관

        self.concdic = {"1": "체결", "2": "확인", "3": "거부", "4": "접수"}

    # PLUS 로 부터 실제로 이벤트(체결/주문 응답/시세 이벤트 등)를 수신 받아 처리하는 함수.
    # 여러가지 이벤트가 이 클래스로 들어 오기 때문에 구분은  이벤트 등록 시 사용한 self.name 을 통해 구분한다.
    def OnReceived(self):
        # 주문 Request 에 대한 응답 처리
        if self.name == "td0314":
            print("[CpEvent]주문응답")
            self.parent.OrderReply()
            return

        # 주문 체결 PB 에 대한 처리
        elif self.name == "conclusion":
            # 주문 체결 실시간 업데이트
            i3 = self.client.GetHeaderValue(3)  # 체결 수량
            i4 = self.client.GetHeaderValue(4)  # 가격
            i5 = self.client.GetHeaderValue(5)  # 주문번호
            i6 = self.client.GetHeaderValue(6)  # 원주문번호
            i9 = self.client.GetHeaderValue(9)  # 종목코드
            i12 = self.client.GetHeaderValue(12)  # 매수/매도 구분 1 매도 2매수
            i14 = self.client.GetHeaderValue(14)  # 체결 플래그 1 체결 2 확인...
            i15 = self.client.GetHeaderValue(15)  # 신용대출구분
            i16 = self.client.GetHeaderValue(16)  # 정정/취소 구분코드 (1 정상, 2 정정 3 취소)
            i17 = self.client.GetHeaderValue(17)  # 현금신용대용 구분
            i18 = self.client.GetHeaderValue(18)  # 주문호가구분코드
            i19 = self.client.GetHeaderValue(19)  # 주문조건구분코드
            i20 = self.client.GetHeaderValue(20)  # 대출일
            i21 = self.client.GetHeaderValue(21)  # 장부가
            i22 = self.client.GetHeaderValue(22)  # 매도가능수량
            i23 = self.client.GetHeaderValue(23)  # 체결기준잔고수량

            # for debug
            print("[CpEvent]%s, 수량 %d, 가격 %d, 주문번호 %d, 원주문 %d, 코드 %s, 매도매수 %s, 신용대출 %s 정정취소 %s,"
                  "현금신용대용 %s, 주문호가구분 %s, 주문조건구분 %s, 대출일 %s, 장부가 %d, 매도가능 %d, 체결기준잔고%d"
                  % (self.concdic.get(i14), i3, i4, i5, i6, i9, i12, i15, i16, i17, i18, i19, i20, i21, i22, i23))

            # 체결 에 대한 처리
            #   미체결에서 체결이 발생한 주문번호를 찾아 주문 수량과 체결 수량을 비교한다
            #   전부 체결이면 미체결을 지우고, 부분 체결일 경우 체결 된 수량만큼 주문 수량에서 제한다.
            if (i14 == "1"):  # 체결
                if not (i5 in self.parent.diOrderList):
                    print("[CpEvent]주문번호 찾기 실패", i5)
                    return
                item = self.parent.diOrderList[i5]
                if (item.amount - i3 > 0):  # 일부 체결인경우
                    # 기존 데이터 업데이트
                    item.amount -= i3
                    item.modAvali = item.amount
                    item.ContAmount += i3
                else:  # 전체 체결인 경우
                    self.parent.deleteOrderNum(i5)

                # for debug
                # for i in range(len(self.parent.orderList)):
                #    self.parent.orderList[i].debugPrint()
                print("[CpEvent]미체결 개수 ", len(self.parent.orderList))



            # 확인 에 대한 처리
            #   정정확인 - 정정주문이 발생한 원주문을 찾아
            #       부분 정정인 경우 - 기존 주문은 수량을 업데이트, 새로운 정정에 의한 미체결 주문번호는 신규 추가
            #       전체 정정인 경우 - 주문 리스트의 원주문/주문번호만 업데이트
            #   취소 확인 - 취소주문이 발생한 원주문을 찾아 미체결 리스트에서 제거 한다.
            elif (i14 == "2"):  # 확인
                # 원주문 번호로 찾는다.
                if not (i6 in self.parent.diOrderList):
                    print("[CpEvent]원주문번호 찾기 실패", i6)
                    # IOC/FOK 의 경우 취소 주문을 낸적이 없어도 자동으로 취소 확인이 들어 온다.
                    if i5 in self.parent.diOrderList and (i16 == "3"):
                        self.parent.deleteOrderNum(i5)
                        self.parent.ForwardPB("cancelpb", i5)

                    return
                item = self.parent.diOrderList[i6]
                if (i16 == "2"):  # 정정 확인 ==> 미체결 업데이트 해야 함.
                    print("[CpEvent]정정확인", item.amount, i3)
                    if (item.amount - i3 > 0):  # 일부 정정인 경우
                        # 기존 데이터 업데이트
                        item.amount -= i3
                        item.modAvali = item.amount
                        # 새로운  미체결 추가
                        item2 = orderData()
                        item2.code = i9
                        item2.name = g_objCodeMgr.CodeToName(i9)
                        item2.orderNum = i5
                        item2.orderPrev = i6
                        item2.buysell = i12
                        item2.modAvali = item2.amount = i3
                        item2.price = i4
                        item2.orderFlag = i18
                        item2.debugPrint()
                        self.parent.diOrderList[i5] = item2
                        self.parent.orderList.append(item2)

                    else:  # 잔량 정정 인 경우 ==> 업데이트
                        item.orderNum = i5  # 주문번호 변경
                        item.orderPrev = i6  # 원주문번호 변경
                        item.modAvali = item.amount = i3
                        item.price = i4
                        item.orderFlag = i18
                        item.debugPrint()
                        # 주문번호가  변경 되어 기존 key 는 제거
                        self.parent.diOrderList[i5] = item
                        del self.parent.diOrderList[i6]


                elif (i16 == "3"):  # 취소 확인 ==> 미체결 찾아 지운다.
                    self.parent.deleteOrderNum(i6)
                    self.parent.ForwardPB("cancelpb", i6)
                # for debug
                # for i in range(len(self.parent.orderList)):
                #    self.parent.orderList[i].debugPrint()
                print("[CpEvent]미체결 개수 ", len(self.parent.orderList))

            elif (i14 == "3"):  # 거부
                print("[CpEvent]거부")

            # 접수 - 신규 접수만 처리. 새로운 주문에 대한 접수는 미체결 리스트에 추가한다.
            elif (i14 == "4"):  # 접수
                if not (i16 == "1"):
                    print("[CpEvent]정정이나 취소 접수는 일단 무시한다.")
                    return
                item = orderData()
                item.code = i9
                item.name = g_objCodeMgr.CodeToName(i9)
                item.orderNum = i5
                item.buysell = i12
                item.modAvali = item.amount = i3
                item.price = i4
                item.orderFlag = i18
                item.debugPrint()
                self.parent.diOrderList[i5] = item
                self.parent.orderList.append(item)

                print("[CpEvent]미체결 개수 ", len(self.parent.orderList))

            return


# SB/PB 요청 ROOT 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def __del__(self):
        self.Unsubscribe()

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
class CpConclution(CpPublish):
    def __init__(self):
        super().__init__("conclusion", "DsCbo1.CpConclusion")


# 취소 주문 요청에 대한 응답 이벤트 처리 클래스
class CpPB0314:
    def __init__(self, obj):
        self.name = "td0314"
        self.obj = obj

    def Subscribe(self, parent):
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, parent)


# 주식 주문 취소 클래스
class CpRPOrder:
    def __init__(self):
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        self.objCancelOrder = win32com.client.Dispatch("CpTrade.CpTd0314")  # 취소
        self.callback = None
        self.bIsRq = False
        self.RqOrderNum = 0  # 취소 주문 중인 주문 번호

    # 주문 취소 통신 - Request 를 이용하여 취소 주문
    # callback 은 취소 주문의 reply 이벤트를 전달하기 위해 필요
    def RequestCancel(self, ordernum, code, amount, callback):
        # 주식 취소 주문
        if self.bIsRq:
            print("RequestCancel - 통신 중이라 주문 불가 ")
            return False
        self.callback = callback
        print("[CpRPOrder/RequestCancel]취소주문", ordernum, code, amount)
        self.objCancelOrder.SetInputValue(1, ordernum)  # 원주문 번호 - 정정을 하려는 주문 번호
        self.objCancelOrder.SetInputValue(2, self.acc)  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(3, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(4, code)  # 종목코드
        self.objCancelOrder.SetInputValue(5, amount)  # 정정 수량, 0 이면 잔량 취소임

        # 취소주문 요청
        ret = 0
        while True:
            ret = self.objCancelOrder.Request()
            if ret == 0:
                break

            print("[CpRPOrder/RequestCancel] 주문 요청 실패 ret : ", ret)
            if ret == 4:
                remainTime = g_objCpStatus.LimitRequestRemainTime
                print("연속 통신 초과에 의해 재 통신처리 : ", remainTime / 1000, "초 대기")
                time.sleep(remainTime / 1000)
                continue
            else:  # 1 통신 요청 실패 3 그 외의 오류 4: 주문요청제한 개수 초과
                return False;

        self.bIsRq = True
        self.RqOrderNum = ordernum

        # 주문 응답(이벤트로 수신
        self.objReply = CpPB0314(self.objCancelOrder)
        self.objReply.Subscribe(self)
        return True

    # 취소 주문 - BloockReqeust 를 이용해서 취소 주문
    def BlockRequestCancel(self, ordernum, code, amount, callback):
        # 주식 취소 주문
        self.callback = callback
        print("[CpRPOrder/BlockRequestCancel]취소주문2", ordernum, code, amount)
        self.objCancelOrder.SetInputValue(1, ordernum)  # 원주문 번호 - 정정을 하려는 주문 번호
        self.objCancelOrder.SetInputValue(2, self.acc)  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(3, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objCancelOrder.SetInputValue(4, code)  # 종목코드
        self.objCancelOrder.SetInputValue(5, amount)  # 정정 수량, 0 이면 잔량 취소임

        # 취소주문 요청
        ret = 0
        while True:
            ret = self.objCancelOrder.BlockRequest()
            if ret == 0:
                break;
            print("[CpRPOrder/RequestCancel] 주문 요청 실패 ret : ", ret)
            if ret == 4:
                remainTime = g_objCpStatus.LimitRequestRemainTime
                print("연속 통신 초과에 의해 재 통신처리 : ", remainTime / 1000, "초 대기")
                time.sleep(remainTime / 1000)
                continue
            else:  # 1 통신 요청 실패 3 그 외의 오류 4: 주문요청제한 개수 초과
                return False;

        print("[CpRPOrder/BlockRequestCancel] 주문결과", self.objCancelOrder.GetDibStatus(),
              self.objCancelOrder.GetDibMsg1())
        if self.objCancelOrder.GetDibStatus() != 0:
            return False
        return True

    # 주문 취소 Request 에 대한 응답 처리
    def OrderReply(self):
        self.bIsRq = False

        if self.objCancelOrder.GetDibStatus() != 0:
            print("[CpRPOrder/OrderReply]통신상태",
                  self.objCancelOrder.GetDibStatus(), self.objCancelOrder.GetDibMsg1())
            self.callback.ForwardReply(-1, 0)
            return False

        orderPrev = self.objCancelOrder.GetHeaderValue(1)
        code = self.objCancelOrder.GetHeaderValue(4)
        orderNum = self.objCancelOrder.GetHeaderValue(6)
        amount = self.objCancelOrder.GetHeaderValue(5)

        print("[CpRPOrder/OrderReply] 주문 취소 reply, 취소한 주문:", orderPrev, code, orderNum, amount)

        # 주문 취소를 요청한 클래스로 포워딩 한다.
        if (self.callback != None):
            self.callback.ForwardReply(0, orderPrev)


# 미체결 조회 서비스
class Cp5339:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd5339")
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분

    def Request5339(self, dicOrderList, orderList):
        self.objRq.SetInputValue(0, self.acc)
        self.objRq.SetInputValue(1, self.accFlag[0])
        self.objRq.SetInputValue(4, "0")  # 전체
        self.objRq.SetInputValue(5, "1")  # 정렬 기준 - 역순
        self.objRq.SetInputValue(6, "0")  # 전체
        self.objRq.SetInputValue(7, 20)  # 요청 개수 - 최대 20개

        print("[Cp5339] 미체결 데이터 조회 시작")
        # 미체결 연속 조회를 위해 while 문 사용
        while True:
            ret = self.objRq.BlockRequest()
            if self.objRq.GetDibStatus() != 0:
                print("통신상태", self.objRq.GetDibStatus(), self.objRq.GetDibMsg1())
                return False

            if (ret == 2 or ret == 3):
                print("통신 오류", ret)
                return False;

            # 통신 초과 요청 방지에 의한 요류 인 경우
            while (ret == 4):  # 연속 주문 오류 임. 이 경우는 남은 시간동안 반드시 대기해야 함.
                remainTime = g_objCpStatus.LimitRequestRemainTime
                print("연속 통신 초과에 의해 재 통신처리 : ", remainTime / 1000, "초 대기")
                time.sleep(remainTime / 1000)
                ret = self.objRq.BlockRequest()

            # 수신 개수
            cnt = self.objRq.GetHeaderValue(5)
            print("[Cp5339] 수신 개수 ", cnt)
            if cnt == 0:
                break

            for i in range(cnt):
                item = orderData()
                item.orderNum = self.objRq.GetDataValue(1, i)
                item.orderPrev = self.objRq.GetDataValue(2, i)
                item.code = self.objRq.GetDataValue(3, i)  # 종목코드
                item.name = self.objRq.GetDataValue(4, i)  # 종목명
                item.orderDesc = self.objRq.GetDataValue(5, i)  # 주문구분내용
                item.amount = self.objRq.GetDataValue(6, i)  # 주문수량
                item.price = self.objRq.GetDataValue(7, i)  # 주문단가
                item.ContAmount = self.objRq.GetDataValue(8, i)  # 체결수량
                item.credit = self.objRq.GetDataValue(9, i)  # 신용구분
                item.modAvali = self.objRq.GetDataValue(11, i)  # 정정취소 가능수량
                item.buysell = self.objRq.GetDataValue(13, i)  # 매매구분코드
                item.creditdate = self.objRq.GetDataValue(17, i)  # 대출일
                item.orderFlagDesc = self.objRq.GetDataValue(19, i)  # 주문호가구분코드내용
                item.orderFlag = self.objRq.GetDataValue(21, i)  # 주문호가구분코드

                # 사전과 배열에 미체결 item 을 추가
                dicOrderList[item.orderNum] = item
                orderList.append(item)

            # 연속 처리 체크 - 다음 데이터가 없으면 중지
            if self.objRq.Continue == False:
                print("[Cp5339] 연속 조회 여부: 다음 데이터가 없음")
                break

        return True


# 샘플 코드  메인 클래스
class testMain():
    def __init__(self):
        self.bTradeInit = False
        # 연결 여부 체크
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
        if (g_objCpTrade.TradeInit(0) != 0):
            print("주문 초기화 실패")
            return False
        self.bTradeInit = True

        # 미체결 리스트를 보관한 자료 구조체
        self.diOrderList = dict()  # 미체결 내역 딕셔너리 - key: 주문번호, value - 미체결 레코드
        self.orderList = []  # 미체결 내역 리스트 - 순차 조회 등을 위한 미체결 리스트

        # 미체결 통신 object
        self.obj = Cp5339()
        # 주문 취소 통신 object
        self.objOrder = CpRPOrder()

        # 실시간 주문 체결
        self.contsb = CpConclution()
        self.contsb.Subscribe("", self)

        return

    # 더 이상 미체결이 아닌 주문번호를 찾아 지운다.
    def deleteOrderNum(self, orderNum):
        print("미체결 주문 번호 삭제: ", orderNum)
        del self.diOrderList[orderNum]
        for i in range(len(self.orderList)):
            if (self.orderList[i].orderNum == orderNum):
                del self.orderList[i]
                break

    # 미체결 주문 조회
    def Reqeust5339(self):
        if self.bTradeInit == False:
            print("TradeInit 실패")
            return False

        self.diOrderList = {}
        self.orderList = []
        self.obj.Request5339(self.diOrderList, self.orderList)

        for item in self.orderList:
            item.debugPrint()
        print("[Reqeust5339]미체결 개수 ", len(self.orderList))

    # 첫번째 미체결을 취소 한다.
    # Request 함수 이용 - OnRecieved 이벤트를 통해 응답을 받는다.
    def RequestCancel(self):
        if len(self.orderList) > 0:
            item = self.orderList[0]
            self.objOrder.RequestCancel(item.orderNum, item.code, item.amount, self)

    # 첫번째 미체결을 취소 한다. -  BlockReqest 이용
    def BlockRequestCancel(self):
        print(2)
        if len(self.orderList) > 0:
            item = self.orderList[0]
            self.objOrder.BlockRequestCancel(item.orderNum, item.code, item.amount, self)

    # 일괄 취소
    def RequestCancelAll(self):
        onums = []
        codes = []
        amounts = []
        for item in self.orderList:
            onums.append(item.orderNum)
            codes.append(item.code)
            amounts.append(item.amount)

        for i in range(len(onums)):
            self.objOrder.BlockRequestCancel(onums[i], codes[i], amounts[i], self)

    # 주문 응답 받음.
    def ForwardReply(self, ret, orderNum):
        print("[testMain/ForwardReply] reply ret %d, 주문번호 %d" % (ret, orderNum))

    # 주문 체결에 대한 실시간 업데이트
    def ForwardPB(self, name, orderNum):
        # 취소 확인을 받은 후 , 다음 취소 할 게 있음 취소 주문 전송
        if (name == "cancelpb"):
            print("[testMain/ForwardPB] 취소 확인 받음, 주문번호", orderNum)


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.main = testMain()
        self.setWindowTitle("PLUS API TEST")

        nHeight = 20
        btnNoContract = QPushButton("미체결 조회", self)
        btnNoContract.move(20, nHeight)
        btnNoContract.clicked.connect(self.btnNoContract_clicked)

        nHeight += 50
        btnCancel = QPushButton("취소 주문(Request)", self)
        btnCancel.move(20, nHeight)
        btnCancel.resize(200, 30)
        btnCancel.clicked.connect(self.btnCancel_clicked)

        nHeight += 50
        btnCancel2 = QPushButton("취소 주문(BlockRequest)", self)
        btnCancel2.move(20, nHeight)
        btnCancel2.resize(200, 30)
        btnCancel2.clicked.connect(self.btnCancel2_clicked)

        nHeight += 50
        btnAllCancel = QPushButton("일괄 취소", self)
        btnAllCancel.move(20, nHeight)
        btnAllCancel.clicked.connect(self.btnAllCancel_clicked)

        nHeight += 50
        btnExit = QPushButton("종료", self)
        btnExit.move(20, nHeight)
        btnExit.clicked.connect(self.btnExit_clicked)

        nHeight += 50
        self.setGeometry(300, 500, 300, nHeight)

        # 시작 부터 미체결 미리 조회 한다.
        self.main.Reqeust5339()

    # 미체결 조회
    def btnNoContract_clicked(self):
        self.main.Reqeust5339()
        return

    # 취소 주문- 주문 리스트에 최근 거 부터
    def btnCancel_clicked(self):
        self.main.RequestCancel()
        return

    # 취소 주문- 주문 리스트에 최근 거 부터
    def btnCancel2_clicked(self):
        self.main.BlockRequestCancel()
        return

    # 일괄 취소 - 미체결 전체 취소
    def btnAllCancel_clicked(self):
        self.main.RequestCancelAll()
        return

    def btnExit_clicked(self):
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
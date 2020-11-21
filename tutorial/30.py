# 대신증권 API
# 보유 주식 잔고 일괄 매도 예제

# [예제 목적]
# 주식 잔고를 조회 하여 특정 호가(매수 5차 호가)로 일괄 매도 하는 예제
#
# [예제에 사용된 플러스 객체]
# CpTrade.CpTd6033 - 주식 잔고 조회 Object
# DsCbo1.StockMst - 현재가 조회 Object
# CpTrade.CpTd0311 - 주식 매수/매도 주문 Object
# DsCbo1.CpConclusion - [실시간]주문 체결 처리 Object
#
# [주의 사항]
# - 매수 5차로 매도 하는 예제로 실제 매도가 체결 될 수 있으니 주의가 필요
# - 해당 예제의 테스트는 모의투자에서 먼저 검증해 보시길 권장합니다.
# - 예제는 현금 잔고에 대해서만 처리 했음 (신용, 담보 등의 잔고는 처리 안됨)
# - PLUS 연속 조회 오류(15초간 20건) 발생 시에는 매도가 안될 수 있음 : 해당 오류에 대한 별도 처리가 필요
# - 주문 수량은 매도 가능 수량으로 했음(주문 중인 다른 미체결이 있을 경우 남은 매도 가능 수량으로만 매도)

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


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
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

        # 구분값 : 텍스트로 변경하기 위해 딕셔너리 이용
        self.dicflag12 = {'1': '매도', '2': '매수'}
        self.dicflag14 = {'1': '체결', '2': '확인', '3': '거부', '4': '접수'}
        self.dicflag15 = {'00': '현금', '01': '유통융자', '02': '자기융자', '03': '유통대주',
                          '04': '자기대주', '05': '주식담보대출', '07': '채권담보대출',
                          '06': '매입담보대출', '08': '플러스론',
                          '13': '자기대용융자', '15': '유통대용융자'}
        self.dicflag16 = {'1': '정상주문', '2': '정정주문', '3': '취소주문'}
        self.dicflag17 = {'1': '현금', '2': '신용', '3': '선물대용', '4': '공매도'}
        self.dicflag18 = {'01': '보통', '02': '임의', '03': '시장가', '05': '조건부지정가'}
        self.dicflag19 = {'0': '없음', '1': 'IOC', '2': 'FOK'}

    def OnReceived(self):
        #        print(self.name)
        # 실시간 처리 - 주문체결
        if self.name == 'conclution':
            # 주문 체결 실시간 업데이트
            conc = {}

            # 체결 플래그
            conc['체결플래그'] = self.dicflag14[self.client.GetHeaderValue(14)]

            conc['주문번호'] = self.client.GetHeaderValue(5)  # 주문번호
            conc['주문수량'] = self.client.GetHeaderValue(3)  # 주문/체결 수량
            conc['주문가격'] = self.client.GetHeaderValue(4)  # 주문/체결 가격
            conc['원주문'] = self.client.GetHeaderValue(6)
            conc['종목코드'] = self.client.GetHeaderValue(9)  # 종목코드
            conc['종목명'] = g_objCodeMgr.CodeToName(conc['종목코드'])

            conc['매수매도'] = self.dicflag12[self.client.GetHeaderValue(12)]

            flag15 = self.client.GetHeaderValue(15)  # 신용대출구분코드
            if (flag15 in self.dicflag15):
                conc['신용대출'] = self.dicflag15[flag15]
            else:
                conc['신용대출'] = '기타'

            conc['정정취소'] = self.dicflag16[self.client.GetHeaderValue(16)]
            conc['현금신용'] = self.dicflag17[self.client.GetHeaderValue(17)]
            conc['주문조건'] = self.dicflag19[self.client.GetHeaderValue(19)]

            conc['체결기준잔고수량'] = self.client.GetHeaderValue(23)
            loandate = self.client.GetHeaderValue(20)
            if (loandate == 0):
                conc['대출일'] = ''
            else:
                conc['대출일'] = str(loandate)
            flag18 = self.client.GetHeaderValue(18)
            if (flag18 in self.dicflag18):
                conc['주문호가구분'] = self.dicflag18[flag18]
            else:
                conc['주문호가구분'] = '기타'

            conc['장부가'] = self.client.GetHeaderValue(21)
            conc['매도가능'] = self.client.GetHeaderValue(22)

            print(conc)
            self.caller.updateJangoCont(conc)

            return


class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

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


# CpPBConclusion: 실시간 주문 체결 수신 클래그
class CpPBConclusion(CpPublish):
    def __init__(self):
        super().__init__('conclution', 'DsCbo1.CpConclusion')


# Cp6033 : 주식 잔고 조회
class Cp6033:
    def __init__(self):
        acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = g_objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)

    # 실제적인 6033 통신 처리
    def requestJango(self, caller):
        while True:
            ret = self.objRq.BlockRequest()
            if ret == 4:
                remainTime = g_objCpStatus.LimitRequestRemainTime
                print('연속조회 제한 오류, 남은 시간', remainTime)
                return False
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(7)
            print(cnt)

            for i in range(cnt):
                item = {}
                code = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목코드'] = code
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['현금신용'] = self.objRq.GetDataValue(1, i)  # 신용구분
                print(code, '현금신용', item['현금신용'])
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                # 잔고 추가
                caller.jangoData[code] = item

                if len(caller.jangoData) >= 200:  # 최대 200 종목만,
                    break

            if len(caller.jangoData) >= 200:
                break
            if (self.objRq.Continue == False):
                break
        return True


# 현재가 - 한종목 통신
class CpRPCurrentPrice:
    def __init__(self, caller):
        self.caller = caller
        self.objStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        return

    def Request(self, code):
        self.caller.curData[code] = {}
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        item = {}
        item['code'] = code
        # caller.curData['종목명'] = g_objCodeMgr.CodeToName(code)
        item['cur'] = self.objStockMst.GetHeaderValue(11)  # 종가
        item['diff'] = self.objStockMst.GetHeaderValue(12)  # 전일대비
        item['vol'] = self.objStockMst.GetHeaderValue(18)  # 거래량

        # 10차호가
        for i in range(10):
            key1 = 'offer%d' % (i + 1)
            key2 = 'bid%d' % (i + 1)
            item[key1] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            item[key2] = (self.objStockMst.GetDataValue(1, i))  # 매수호가

        self.caller.curData[code] = item
        return True


# 주식 주문 처리
class CpRPOrder:
    def __init__(self, caller):
        self.caller = caller
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])
        self.objOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 매수

    def buyOrder(self, code, price, amount):
        # 주식 매수 주문
        print("신규 매수", code, price, amount)

        self.objOrder.SetInputValue(0, "2")  # 2: 매수
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

        # 매수 주문 요청
        ret = self.objOrder.BlockRequest()
        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('주의: 주문 연속 통신 제한에 걸렸음. 대기해서 주문할 지 여부 판단이 필요 남은 시간', remainTime)
            return False

        rqStatus = self.objOrder.GetDibStatus()
        rqRet = self.objOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        return True

    def sellOrder(self, code, price, amount):
        # 주식 매도 주문
        print("신규 매도", code, price, amount)

        self.objOrder.SetInputValue(0, "1")  # 1: 매도
        self.objOrder.SetInputValue(1, self.acc)  # 계좌번호
        self.objOrder.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objOrder.SetInputValue(3, code)  # 종목코드
        self.objOrder.SetInputValue(4, amount)  # 매수수량
        self.objOrder.SetInputValue(5, price)  # 주문단가
        self.objOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통

        # 매도 주문 요청
        ret = self.objOrder.BlockRequest()
        if ret == 4:
            remainTime = g_objCpStatus.LimitRequestRemainTime
            print('주의: 주문 연속 통신 제한에 걸렸음. 대기해서 주문할 지 여부 판단이 필요 남은 시간', remainTime)
            return False

        rqStatus = self.objOrder.GetDibStatus()
        rqRet = self.objOrder.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        self.setWindowTitle("주식 잔고 일괄매도 예제")

        # 6033 잔고 object
        self.obj6033 = Cp6033()
        self.jangoData = {}
        self.objConclusion = CpPBConclusion()

        self.objRpCur = CpRPCurrentPrice(self)
        self.curData = {}

        self.objRpOrder = CpRPOrder(self)

        nH = 20
        btnSellAll = QPushButton('전체매도', self)
        btnSellAll.move(20, nH)
        btnSellAll.clicked.connect(self.btnSellAll_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        self.setGeometry(300, 300, 300, nH)

    def btnSellAll_clicked(self):
        # 1. 잔고 요청
        self.objConclusion.Unsubscribe()
        if self.obj6033.requestJango(self) == False:
            return
        self.objConclusion.Subscribe('', self)

        for code, value in self.jangoData.items():
            print(code, value)
            # 2. 현재가 통신
            self.objRpCur.Request(code)

            # 3. 매수 5호가로 매도 주문
            if (self.curData[code]['bid5'] > 0):
                jangoNum = value['잔고수량']
                amount = value['매도가능']
                if (jangoNum != amount):
                    print("경고: 미체결 수량이 있어 잔고와 매도 가능 수량이 다름", code, jangoNum, amount)

                price = self.curData[code]['bid5']
                self.objRpOrder.sellOrder(code, price, amount)

        return

    def btnExit_clicked(self):
        exit()
        return

    # 매도 후 실시간 주문 체결 받는 로직
    def updateJangoCont(self, pbCont):
        # 주문 체결에서 들어온 신용 구분 값 ==> 잔고 구분값으로 치환
        dicBorrow = {
            '현금': ord(' '),
            '유통융자': ord('Y'),
            '자기융자': ord('Y'),
            '주식담보대출': ord('B'),
            '채권담보대출': ord('B'),
            '매입담보대출': ord('M'),
            '플러스론': ord('P'),
            '자기대용융자': ord('I'),
            '유통대용융자': ord('I'),
            '기타': ord('Z')
        }

        # 잔고 리스트 map 의 key 값
        code = pbCont['종목코드']

        # 접수, 거부, 확인 등은 매도 가능 수량만 업데이트 한다.
        if pbCont['체결플래그'] == '접수' or pbCont['체결플래그'] == '거부' or pbCont['체결플래그'] == '확인':
            if (code not in self.jangoData):
                return
            self.jangoData[code]['매도가능'] = pbCont['매도가능']
            return

        if (pbCont['체결플래그'] == '체결'):
            if (code not in self.jangoData):  # 신규 잔고 추가
                print('신규 잔고 추가', code)
                # 신규 잔고 추가
                item = {}
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                self.jangoData[code] = item

            else:
                # 기존 잔고 업데이트
                item = self.jangoData[code]
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                # 잔고 수량이 0 이면 잔고 제거
                if item['잔고수량'] == 0:
                    del self.jangoData[code]
                    print('매도 전부 체결', item['종목명'])

        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
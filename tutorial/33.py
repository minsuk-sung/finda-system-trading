# 대신증권 API
# 주식 잔고 실시간 조회(현재가 및 주문 체결 실시간 반영)

# 이번 예제는 PLUS API 를 이용하여 주식 잔고를 조회 하고, 실시간으로 현재가와 주문 체결 변경 상황을 관리하는 예제 입니다.
# ■ 사용된 클래스
#     ▥ CpEvent - 실시간 이벤트 수신 (현재가와 주문 체결 실시간 처리)
#     ▥ Cp6033 - 주식 잔고 조회
#     ▥ CpRPCurrentPrice - 현재가 한 종목 조회
#     ▥ CpMarketEye - 복수 현재가 종목 조회
# ■ 활용
#     잔고데이터에 대해 조회 및 매수 주문이나 매도 주문 후 잔고에 반영 여부를 실시간으로 확인가능합니다.
#
# ※ 주의사항: 본 예제는 PLUS 활용을 돕기 위해 예제로만 제공됩니다.
# 또한 현금 잔고에 대해서만 처리 하고 있어 신용 잔고 등은 별도 코드 처리가 필요 하니 이 점 유의하시기 바랍니다.

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

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
        # 실시간 처리 - 현재가 주문 체결
        if self.name == 'stockcur':
            code = self.client.GetHeaderValue(0)  # 초
            name = self.client.GetHeaderValue(1)  # 초
            timess = self.client.GetHeaderValue(18)  # 초
            exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            cprice = self.client.GetHeaderValue(13)  # 현재가
            diff = self.client.GetHeaderValue(2)  # 대비
            cVol = self.client.GetHeaderValue(17)  # 순간체결수량
            vol = self.client.GetHeaderValue(9)  # 거래량

            item = {}
            item['code'] = code
            # rpName = self.objRq.GetDataValue(1, i)  # 종목명
            # rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = vol

            # 현재가 업데이트
            self.caller.updateJangoCurPBData(item)

        # 실시간 처리 - 주문체결
        elif self.name == 'conclution':
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
            conc['매도가능수량'] = self.client.GetHeaderValue(22)

            print(conc)
            self.caller.updateJangoCont(conc)

            return


################################################
# plus 실시간 수신 base 클래스
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


################################################
# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('stockcur', 'DsCbo1.StockCur')


################################################
# CpPBConclusion: 실시간 주문 체결 수신 클래그
class CpPBConclusion(CpPublish):
    def __init__(self):
        super().__init__('conclution', 'DsCbo1.CpConclusion')


################################################
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
        self.dicflag1 = {ord(' '): '현금',
                         ord('Y'): '융자',
                         ord('D'): '대주',
                         ord('B'): '담보',
                         ord('M'): '매입담보',
                         ord('P'): '플러스론',
                         ord('I'): '자기융자',
                         }

    # 실제적인 6033 통신 처리
    def requestJango(self, caller):
        while True:
            self.objRq.BlockRequest()
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
                item['현금신용'] = self.dicflag1[self.objRq.GetDataValue(1, i)]  # 신용구분
                print(code, '현금신용', item['현금신용'])
                item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
                item['잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
                item['매도가능'] = self.objRq.GetDataValue(15, i)
                item['장부가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
                # item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
                # item['평가손익'] = self.objRq.GetDataValue(11, i)  # 평가손익(천원미만은 절사 됨)
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']
                item['현재가'] = 0
                item['대비'] = 0
                item['거래량'] = 0

                # 잔고 추가
                #                key = (code, item['현금신용'],item['대출일'] )
                key = code
                caller.jangoData[key] = item

                if len(caller.jangoData) >= 200:  # 최대 200 종목만,
                    break

            if len(caller.jangoData) >= 200:
                break
            if (self.objRq.Continue == False):
                break
        return True


################################################
# 현재가 - 한종목 통신
class CpRPCurrentPrice:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        return

    def Request(self, code, caller):
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
        caller.curDatas[code] = item
        '''
        caller.curData['기준가'] = self.objStockMst.GetHeaderValue(27)  # 기준가
        caller.curData['예상플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        caller.curData['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        caller.curData['예상대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결대비
        # 10차호가
        for i in range(10):
            key1 = '매도호가%d' % (i + 1)
            key2 = '매수호가%d' % (i + 1)
            caller.curData[key1] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            caller.curData[key2] = (self.objStockMst.GetDataValue(1, i))  # 매수호가
        '''

        return True


################################################
# CpMarketEye : 복수종목 현재가 통신 서비스
class CpMarketEye:
    def __init__(self):
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        self.rqField = [0, 1, 2, 3, 4, 10, 17]  # 요청 필드

        # 관심종목 객체 구하기
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")

    def Request(self, codes, caller):
        # 요청 필드 세팅 - 종목코드, 종목명, 시간, 대비부호, 대비, 현재가, 거래량
        self.objRq.SetInputValue(0, self.rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(2)

        for i in range(cnt):
            item = {}
            item['code'] = self.objRq.GetDataValue(0, i)  # 코드
            # rpName = self.objRq.GetDataValue(1, i)  # 종목명
            # rpDiffFlag = self.objRq.GetDataValue(3, i)  # 대비부호
            item['diff'] = self.objRq.GetDataValue(3, i)  # 대비
            item['cur'] = self.objRq.GetDataValue(4, i)  # 현재가
            item['vol'] = self.objRq.GetDataValue(5, i)  # 거래량

            caller.curDatas[item['code']] = item

        return True


################################################
# 테스트를 위한 메인 화면
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        self.setWindowTitle("주식 잔고(실시간) 처리 예제")
        self.setGeometry(300, 300, 300, 180)

        # 6033 잔고 object
        self.obj6033 = Cp6033()
        self.jangoData = {}

        self.isSB = False
        self.objCur = {}

        # 현재가 정보
        self.curDatas = {}
        self.objRPCur = CpRPCurrentPrice()

        # 실시간 주문 체결
        self.objConclusion = CpPBConclusion()

        nH = 20
        btnExcel = QPushButton('Excel 내보내기', self)
        btnExcel.move(20, nH)
        btnExcel.clicked.connect(self.btnExcel_clicked)
        nH += 50

        btnPrint = QPushButton('잔고 Print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        # 잔고 요청
        self.requestJango()

    def StopSubscribe(self):
        if self.isSB:
            for key, obj in self.objCur.items():
                obj.Unsubscribe()
            self.objCur = {}

        self.isSB = False
        self.objConclusion.Unsubscribe()

    def requestJango(self):
        self.StopSubscribe();

        # 주식 잔고 통신
        if self.obj6033.requestJango(self) == False:
            return

        # 잔고 현재가 통신
        codes = set()
        for code, value in self.jangoData.items():
            codes.add(code)

        objMarkeyeye = CpMarketEye()
        codelist = list(codes)
        if (objMarkeyeye.Request(codelist, self) == False):
            exit()

        # 실시간 현재가  요청
        cnt = len(codelist)
        for i in range(cnt):
            code = codelist[i]
            self.objCur[code] = CpPBStockCur()
            self.objCur[code].Subscribe(code, self)
        self.isSB = True

        # 실시간 주문 체결 요청
        self.objConclusion.Subscribe('', self)

    def btnExcel_clicked(self):
        return

    def btnPrint_clicked(self):
        print('잔고')
        for code, value in self.jangoData.items():
            print(code, value)

        print('실시간 현재가 수신 중인 종목')
        for key, obj in self.objCur.items():
            print(key)

        return

    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()
        return

    # 실시간 주문 체결 처리 로직
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
        # key = (pbCont['종목코드'], dicBorrow[pbCont['현금신용']], pbCont['대출일'])
        # key = pbCont['종목코드']
        code = pbCont['종목코드']

        # 접수, 거부, 확인 등은 매도 가능 수량만 업데이트 한다.
        if pbCont['체결플래그'] == '접수' or pbCont['체결플래그'] == '거부' or pbCont['체결플래그'] == '확인':
            if (code not in self.jangoData):
                return
            self.jangoData[code]['매도가능'] = pbCont['매도가능수량']
            return

        if (pbCont['체결플래그'] == '체결'):
            if (code not in self.jangoData):  # 신규 잔고 추가
                if (pbCont['체결기준잔고수량'] == 0):
                    return
                print('신규 잔고 추가', code)
                # 신규 잔고 추가
                item = {}
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                print('신규 현재가 요청', code)
                self.objRPCur.Request(code, self)
                self.objCur[code] = CpPBStockCur()
                self.objCur[code].Subscribe(code, self)

                item['현재가'] = self.curDatas[code]['cur']
                item['대비'] = self.curDatas[code]['diff']
                item['거래량'] = self.curDatas[code]['vol']

                self.jangoData[code] = item

            else:
                # 기존 잔고 업데이트
                item = self.jangoData[code]
                item['종목코드'] = pbCont['종목코드']
                item['종목명'] = pbCont['종목명']
                item['현금신용'] = dicBorrow[pbCont['현금신용']]
                item['대출일'] = pbCont['대출일']
                item['잔고수량'] = pbCont['체결기준잔고수량']
                item['매도가능'] = pbCont['매도가능수량']
                item['장부가'] = pbCont['장부가']
                # 매입금액 = 장부가 * 잔고수량
                item['매입금액'] = item['장부가'] * item['잔고수량']

                # 잔고 수량이 0 이면 잔고 제거
                if item['잔고수량'] == 0:
                    del self.jangoData[code]
                    self.objCur[code].Unsubscribe()
                    del self.objCur[code]

        return

    # 실시간 현재가 처리 로직
    def updateJangoCurPBData(self, curData):
        code = curData['code']
        self.curDatas[code] = curData
        self.upjangoCurData(code)

    def upjangoCurData(self, code):
        # 잔고에 동일 종목을 찾아 업데이트 하자 - 현재가/대비/거래량/평가금액/평가손익
        curData = self.curDatas[code]
        item = self.jangoData[code]
        item['현재가'] = curData['cur']
        item['대비'] = curData['diff']
        item['거래량'] = curData['vol']


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
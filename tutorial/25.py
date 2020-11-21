# 대신증권 API
# 해외선물 잔고(미결제약정)과 실시간 주문체결 처리 예제

# 해외선물 미결제 내역을 조회 하고 실시간으로 주문 체결 처리를 하는 참고용 예제입니다.
#
# 사용된 PLUS OBJECT
# ■ CpForeTrade.OvfNotPaymentInq - 해외선물 미결제(잔고) 조회
# ■ CpForeDib.OvFutBalance : 해외선물 미결제(잔고) 실시간 업데이트
#
# 제공되는 기능
# - 해외선물 미결제 잔고 조회
# - 실사간 주문 체결 업데이트
#
# 미제공 기능
# - 현재가 조회 및 실시간 현재가 업데이트 미제공 예제임
# - 평가금액 실시간 업데이트 안됨.
#
# ※ 주의사항: 본 예제는 단순 참고용으로만 제공되는 예제임

import sys
from PyQt5.QtWidgets import *
import win32com.client
from pandas import Series, DataFrame
import pandas as pd
import locale
import os

locale.setlocale(locale.LC_ALL, '')
# cp object
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objOptionMgr = win32com.client.Dispatch("CpUtil.CpOptionCode")

gExcelFile = 'ovfuturejango.xlsx'


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

    def OnReceived(self):
        if self.name == "ovfjango":
            pbdata = {}
            pbdata["처리구분"] = self.client.GetHeaderValue(2)  # “00”-주문, “01”-정정, “02”-취소, “03”-체결
            pbdata["code"] = self.client.GetHeaderValue(6)  # 종목코드
            pbdata["종목명"] = self.client.GetHeaderValue(7)  # 종목코드
            pbdata["매매구분"] = self.client.GetHeaderValue(8)  # 매매구분
            pbdata["잔고수량"] = self.client.GetHeaderValue(9)  # 잔고수량
            pbdata["단가"] = self.client.GetHeaderValue(10)  # 단가
            pbdata["청산가능"] = self.client.GetHeaderValue(11)  # 청산가능수량
            pbdata["미체결수량"] = self.client.GetHeaderValue(12)  # 미체결수량
            pbdata["현재가"] = self.client.GetHeaderValue(13)  # 현재가
            pbdata["대비부호"] = self.client.GetHeaderValue(14)  # 전일대비부호
            pbdata["전일대비"] = self.client.GetHeaderValue(15)  # 전일대비
            pbdata["전일대비율"] = self.client.GetHeaderValue(16)  # 전일대비율
            pbdata["평가금액"] = self.client.GetHeaderValue(17)  # 평가금액
            pbdata["평가손익"] = self.client.GetHeaderValue(18)  # 평가손익
            pbdata["손익률"] = self.client.GetHeaderValue(19)  # 손익률
            pbdata["매입금액"] = self.client.GetHeaderValue(20)  # 매입금액
            pbdata["승수"] = self.client.GetHeaderValue(21)  # 승수
            pbdata["통화코드"] = self.client.GetHeaderValue(23)  # 통화코드

            print(pbdata)
            self.caller.updateContract(pbdata)


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


class CpPBOvfJango(CpPublish):
    def __init__(self):
        super().__init__('ovfjango', 'CpForeDib.OvFutBalance')


# 해외선물 잔고(미결제) 통신
class CpOvfJango:
    def __init__(self):
        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호

    def Request(self, caller):
        if (g_objCpStatus.IsConnect == 0):
            print('PLUS가 정상적으로 연결되지 않음. ')
            return False

        # 해외선물 미결제 잔고
        objRq = win32com.client.Dispatch('CpForeTrade.OvfNotPaymentInq')
        objRq.SetInputValue(1, self.acc)  # 계좌번호

        while True:
            objRq.BlockRequest()
            # 현재가 통신 및 통신 에러 처리
            rqStatus = objRq.GetDibStatus()
            rqRet = objRq.GetDibMsg1()
            print('통신상태', rqStatus, rqRet)
            if rqStatus != 0:
                return False

            # 조회 건수
            cnt = objRq.GetHeaderValue(0)
            print(cnt)
            if cnt == 0:
                break

            for i in range(cnt):
                item = {}
                item['code'] = objRq.GetDataValue(3, i)  # 코드
                item['종목명'] = objRq.GetDataValue(4, i)
                item['매매구분'] = objRq.GetDataValue(5, i)
                item['잔고수량'] = objRq.GetDataValue(6, i)
                item['단가'] = objRq.GetDataValue(7, i)
                item['청산가능'] = objRq.GetDataValue(8, i)
                item['미체결수량'] = objRq.GetDataValue(9, i)
                item['현재가'] = objRq.GetDataValue(10, i)
                item['전일대비'] = objRq.GetDataValue(11, i)
                item['전일대비율'] = objRq.GetDataValue(12, i)
                item['평가금액'] = objRq.GetDataValue(13, i)
                item['평가손익'] = objRq.GetDataValue(14, i)
                item['손익률'] = objRq.GetDataValue(15, i)
                item['매입금액'] = objRq.GetDataValue(16, i)
                item['승수'] = objRq.GetDataValue(17, i)
                item['통화코드'] = objRq.GetDataValue(18, i)

                key = item['code'] + item['매매구분']

                caller.ovfJangadata[key] = item
                print(item)

            if objRq.Continue == False:
                print("연속 조회 여부: 다음 데이터가 없음")
                break


#        print(self.data)


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.bTradeInit = False
        # 연결 여부 체크
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
        if (g_objCpTrade.TradeInit(0) != 0):
            print("주문 초기화 실패")
            return False
        self.bTradeInit = True

        self.pbContract = CpPBOvfJango()

        self.setWindowTitle('PLUS API TEST')
        self.setGeometry(300, 300, 300, 240)

        # 해외선물 잔고
        self.ovfJangadata = {}

        nH = 20

        btnPrint = QPushButton('DF Print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        btnExcel = QPushButton('엑셀 내보내기', self)
        btnExcel.move(20, nH)
        btnExcel.clicked.connect(self.btnExcel_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)

        self.btnStart_clicked()

    def btnStart_clicked(self):
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        obj = CpOvfJango()
        obj.Request(self);

        self.pbContract.Unsubscribe()
        self.pbContract.Subscribe("", self)

    def btnPrint_clicked(self):
        for key, value in self.ovfJangadata.items():
            print(key, value)

    def btnExit_clicked(self):
        self.pbContract.Unsubscribe()
        exit()

    def btnExcel_clicked(self):
        if (len(self.ovfJangadata) == 0):
            print('잔고 없음')
            return

        #        df= pd.DataFrame(columns=self.ovfJangadata.keys())

        isFirst = True
        for k, v in self.ovfJangadata.items():
            # 데이터 프레임의 컬럼은 데이터의 key 값으로 생성
            if isFirst:
                df = pd.DataFrame(columns=v.keys())
                isFirst = False
            df.loc[len(df)] = v

        # create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(gExcelFile, engine='xlsxwriter')
        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name='Sheet1')
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        os.startfile(gExcelFile)
        return

    # 실시간 주문 체결 업데이트
    def updateContract(self, pbdata):
        key = pbdata['code'] + pbdata['매매구분']
        print(key)

        # 새로운 잔고 추가
        if key not in self.ovfJangadata.keys():
            print(key, '찾기 실패')
            if pbdata["잔고수량"] == 0:
                return

            item = {}
            item['code'] = pbdata['code']
            item['종목명'] = pbdata['종목명']
            item['매매구분'] = pbdata['매매구분']
            item['잔고수량'] = pbdata['잔고수량']
            item['단가'] = pbdata['단가']
            item['청산가능'] = pbdata['청산가능']
            item['미체결수량'] = pbdata['미체결수량']
            item['현재가'] = pbdata['현재가']
            item['전일대비'] = pbdata['전일대비']
            item['전일대비율'] = pbdata['전일대비율']
            item['평가금액'] = pbdata['평가금액']
            item['평가손익'] = pbdata['평가손익']
            item['손익률'] = pbdata['손익률']
            item['매입금액'] = pbdata['매입금액']
            item['승수'] = pbdata['승수']
            item['통화코드'] = pbdata['통화코드']
            self.ovfJangadata[key] = item
            print('새로운 잔고 추가 -', key)
            return

        # 기존 잔고에 대한 처리
        item = self.ovfJangadata[key]
        item['잔고수량'] = pbdata['잔고수량']
        item['청산가능'] = pbdata['청산가능']

        if (pbdata["처리구분"] == '00'):  # 주문 접수
            print('주문 접수 -', key)
            self.ovfJangadata[key] = item
            return

        if item['잔고수량'] == 0:  # 잔고 삭제
            del self.ovfJangadata[key]
            print('잔고 삭제 -', key)
            return

        print('체결 업데이트-', key)
        item['단가'] = pbdata['단가']
        item['미체결수량'] = pbdata['미체결수량']
        item['현재가'] = pbdata['현재가']
        item['전일대비'] = pbdata['전일대비']
        item['전일대비율'] = pbdata['전일대비율']
        item['평가금액'] = pbdata['평가금액']
        item['평가손익'] = pbdata['평가손익']
        item['손익률'] = pbdata['손익률']
        item['매입금액'] = pbdata['매입금액']
        self.ovfJangadata[key] = item

        return


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
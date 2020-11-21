# 대신증권 API
# 지수옵션 최근월물 시세 조회(실시간 포함)

# 지수 옵션 최근월물을 가져와 현재가를 실시간으로 조회 하는 예제입니다
# 엑셀 내보내기를 통해 수신 받은 데이터를 확인할 수 있습니다

# 사용된 PLUS OBJET 는 다음과 같습니다
# ■ CpUtil.CpOptionCode : 옵션 종목코드 정보 구하기
# ■ CpSysDib.MarketEye : 복수 종목 동시 조회
# ■ CpSysDib.OptionCurOnly : [실시간] 옵션 종목 실시간 시세 수신

import sys
from PyQt5.QtWidgets import *
import win32com.client
from pandas import Series, DataFrame
import pandas as pd
import locale
import os
import time
import datetime

locale.setlocale(locale.LC_ALL, '')
# cp object
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objOptionMgr = win32com.client.Dispatch("CpUtil.CpOptionCode")

gExcelFile = 'market_data.xlsx'


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

    def OnReceived(self):
        if self.name == 'optioncur':
            pbTime = time.time()
            # curData = {}
            code = self.client.GetHeaderValue(0)
            # curData = self.caller.marketDF.loc[code]

            self.caller.marketDF.set_value(code, 'time', self.client.GetHeaderValue(1))  # 초
            self.caller.marketDF.set_value(code, '시가', self.client.GetHeaderValue(4))
            self.caller.marketDF.set_value(code, '고가', self.client.GetHeaderValue(5))
            self.caller.marketDF.set_value(code, '저가', self.client.GetHeaderValue(6))
            self.caller.marketDF.set_value(code, '매도호가', self.client.GetHeaderValue(17))
            self.caller.marketDF.set_value(code, '매수호가', self.client.GetHeaderValue(18))
            self.caller.marketDF.set_value(code, '현재가', self.client.GetHeaderValue(2))  # 현재가
            self.caller.marketDF.set_value(code, '대비', self.client.GetHeaderValue(3))  # 대비
            self.caller.marketDF.set_value(code, '거래량', self.client.GetHeaderValue(7))  # 거래량
            self.caller.marketDF.set_value(code, '미결제', self.client.GetHeaderValue(16))
            lastday = self.caller.marketDF.get_value(code, '전일종가')
            diff = self.caller.marketDF.get_value(code, '대비')
            if lastday:
                diffp = (diff / lastday) * 100
                self.caller.marketDF.set_value(code, '대비율', diffp)

            # for debug
            # curData = self.caller.marketDF.loc[code]
            # print('실시간', curData)


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


# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__('optioncur', 'CpSysDib.OptionCurOnly')


# CpMarketEye : 복수종목 현재가 통신 서비스
class CpMarketEye:
    def Request(self, codes, caller):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print('PLUS가 정상적으로 연결되지 않음. ')
            return False

        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch('CpSysDib.MarketEye')
        # 필드
        # 0 코드, 1 시간 2:대비부호(char) 3:전일대비 - 주의) 반드시대비부호(2)와같이요청을하여야함
        # 4:현재가 5:시가 6:고가  7:저가 8:매도호가 9:매수호가 10:거래량 23:전일종가
        # 27: 미결제약정
        rqField = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 23, 27]  # 요청 필드
        objRq.SetInputValue(0, rqField)  # 요청 필드
        objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        objRq.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = objRq.GetDibStatus()
        rqRet = objRq.GetDibMsg1()
        print('통신상태', rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = objRq.GetHeaderValue(2)

        caller.marketDF = None
        caller.marketDF = pd.DataFrame(columns=('code', '종목명', 'time', '현재가', '대비',
                                                '대비율', '행사가', '거래량', '매도호가', '매수호가',
                                                '시가', '고가', '저가', '전일종가', '미결제'))

        for i in range(cnt):
            item = {}
            item['code'] = objRq.GetDataValue(0, i)  # 코드
            item['종목명'] = g_objCodeMgr.CodeToName(item['code'])
            item['time'] = objRq.GetDataValue(1, i)  # 시간
            item['대비'] = objRq.GetDataValue(3, i)  # 전일대비
            item['현재가'] = objRq.GetDataValue(4, i)  # 현재가
            item['시가'] = objRq.GetDataValue(5, i)  # 시가
            item['고가'] = objRq.GetDataValue(6, i)  # 고가
            item['저가'] = objRq.GetDataValue(7, i)  # 저가
            item['매도호가'] = objRq.GetDataValue(8, i)  # 매도호가
            item['매수호가'] = objRq.GetDataValue(9, i)  # 매수호가
            item['거래량'] = objRq.GetDataValue(10, i)  # 거래량
            item['전일종가'] = objRq.GetDataValue(11, i)  # 전일종가
            item['미결제'] = objRq.GetDataValue(12, i)

            if item['전일종가'] != 0:
                item['대비율'] = (item['대비'] / item['전일종가']) * 100
            else:
                item['대비율'] = 0
            item['행사가'] = caller.codeTooPrice[item['code']]

            caller.marketDF.loc[len(caller.marketDF)] = item

        # data frmae 의  기본 인덱스(0,1,2,3~ ) ==> 'code' 로 변경
        caller.marketDF = caller.marketDF.set_index('code')
        # 인덱스 이름 제거
        caller.marketDF.index.name = None
        print(caller.marketDF)
        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PLUS API TEST')
        self.setGeometry(300, 300, 300, 240)
        self.isSB = False
        self.objCur = []

        self.marketDF = DataFrame()
        self.codeTooPrice = {}

        btnStart = QPushButton('요청 시작', self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)

        btnExcel = QPushButton('Excel 내보내기', self)
        btnExcel.move(20, 70)
        btnExcel.clicked.connect(self.btnExcel_clicked)

        btnPrint = QPushButton('DF Print', self)
        btnPrint.move(20, 120)
        btnPrint.clicked.connect(self.btnPrint_clicked)

        btnExit = QPushButton('종료', self)
        btnExit.move(20, 190)
        btnExit.clicked.connect(self.btnExit_clicked)

    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, '종목 실시간 해지되었음')
        self.isSB = False

        self.objCur = []

    def btnStart_clicked(self):
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        codes = []
        months = {}
        count = g_objOptionMgr.GetCount()
        # 전체 코드를 읽어 월물별로 종목 정보를 저장
        for i in range(0, count):
            code = g_objOptionMgr.GetData(0, i)
            name = g_objOptionMgr.GetData(1, i)
            mon = g_objOptionMgr.GetData(3, i)
            opprice = g_objOptionMgr.GetData(4, i)  # 행사가
            if not (mon in months.keys()):
                months[mon] = []
            # 튜플()을 이용해서 데이터를 저장.
            months[mon].append((code, name, mon, opprice))
            # print(mon, code)

        # 최근월물만 꺼낸다.
        firstkey = ''
        for key, value in months.items():
            print(key)
            if (firstkey == ''):
                firstkey = key
                break;

        for data in months[firstkey]:
            codes.append(data[0])
            self.codeTooPrice[data[0]] = data[3]

        objMarkeyeye = CpMarketEye()
        if (objMarkeyeye.Request(codes, self) == False):
            exit()

        cnt = len(codes)
        for i in range(cnt):
            self.objCur.append(CpPBStockCur())
            self.objCur[i].Subscribe(codes[i], self)

        print('================-')
        print(cnt, '종목 실시간 현재가 요청 시작')
        self.isSB = True

    def btnExcel_clicked(self):
        print(len(self.marketDF.index))
        # create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(gExcelFile, engine='xlsxwriter')
        # Convert the dataframe to an XlsxWriter Excel object.
        self.marketDF.to_excel(writer, sheet_name='Sheet1')
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        os.startfile(gExcelFile)
        return

    def btnPrint_clicked(self):
        print(self.marketDF)

    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
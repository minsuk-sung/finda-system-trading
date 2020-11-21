# 대신증권 API
# 특징주 포착(뉴스/공시/차트/외국인 신호) 예제

# HTS #8092 특징주 포착(뉴스/차트 신호) 화면 예제 입니다 .
# 특징주 포착 서비스를 통해  뉴스/공시/차트/외국인 관련 주요 신호를 조회 및 실시간으로 받아 볼 수 있습니다
# 자세한 내용은 HTS 화면과 함께 참고 하시기 바랍니다
#
# ■ 사용된 PLUS 객체
#  - CpSysDib.CpMarketWatch - #8092 특징주 포착 조회 서비스
#  - Dscbo1.CpSvr8092S - 실시간 뉴스/공시 수신
#  - CpSysDib.CpMarketWatchS - 실시간 차트/외국인 신호 수신
#
#  제공된 예제는 UI 부분은 없이 수신 된 신호를 기본 print() 로 표출하고 있습니다
#  엑셀 내보내기를 제공하고 있어 수신 받은 내용을 #8092 와 비교 할 수 있습니다.

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes
import pandas as pd
import os

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

gExcelFile = '8092.xlsx'


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

    '''
    # 주문 관련 초기화
    if (g_objCpTrade.TradeInit(0) != 0):
        print("주문 초기화 실패")
        return False
    '''
    return True


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관
        self.diccode = {
            10: '외국계증권사창구첫매수',
            11: '외국계증권사창구첫매도',
            12: '외국인순매수',
            13: '외국인순매도',
            21: '전일거래량갱신',
            22: '최근5일거래량최고갱신',
            23: '최근5일매물대돌파',
            24: '최근60일매물대돌파',
            28: '최근5일첫상한가',
            29: '최근5일신고가갱신',
            30: '최근5일신저가갱신',
            31: '상한가직전',
            32: '하한가직전',
            41: '주가 5MA 상향돌파',
            42: '주가 5MA 하향돌파',
            43: '거래량 5MA 상향돌파',
            44: '주가데드크로스(5MA < 20MA)',
            45: '주가골든크로스(5MA > 20MA)',
            46: 'MACD 매수-Signal(9) 상향돌파',
            47: 'MACD 매도-Signal(9) 하향돌파',
            48: 'CCI 매수-기준선(-100) 상향돌파',
            49: 'CCI 매도-기준선(100) 하향돌파',
            50: 'Stochastic(10,5,5)매수- 기준선상향돌파',
            51: 'Stochastic(10,5,5)매도- 기준선하향돌파',
            52: 'Stochastic(10,5,5)매수- %K%D 교차',
            53: 'Stochastic(10,5,5)매도- %K%D 교차',
            54: 'Sonar 매수-Signal(9) 상향돌파',
            55: 'Sonar 매도-Signal(9) 하향돌파',
            56: 'Momentum 매수-기준선(100) 상향돌파',
            57: 'Momentum 매도-기준선(100) 하향돌파',
            58: 'RSI(14) 매수-Signal(9) 상향돌파',
            59: 'RSI(14) 매도-Signal(9) 하향돌파',
            60: 'Volume Oscillator 매수-Signal(9) 상향돌파',
            61: 'Volume Oscillator 매도-Signal(9) 하향돌파',
            62: 'Price roc 매수-Signal(9) 상향돌파',
            63: 'Price roc 매도-Signal(9) 하향돌파',
            64: '일목균형표매수-전환선 > 기준선상향교차',
            65: '일목균형표매도-전환선 < 기준선하향교차',
            66: '일목균형표매수-주가가선행스팬상향돌파',
            67: '일목균형표매도-주가가선행스팬하향돌파',
            68: '삼선전환도-양전환',
            69: '삼선전환도-음전환',
            70: '캔들패턴-상승반전형',
            71: '캔들패턴-하락반전형',
            81: '단기급락후 5MA 상향돌파',
            82: '주가이동평균밀집-5%이내',
            83: '눌림목재상승-20MA 지지'
        }

    def OnReceived(self):
        print(self.name)
        # 실시간 처리 - marketwatch : 특이 신호(차트, 외국인 순매수 등)
        if self.name == 'marketwatch':
            code = self.client.GetHeaderValue(0)
            name = g_objCodeMgr.CodeToName(code)
            cnt = self.client.GetHeaderValue(2)

            for i in range(cnt):
                item = {}
                newcancel = ''
                time = self.client.GetDataValue(0, i)
                h, m = divmod(time, 100)
                item['시간'] = '%02d:%02d' % (h, m)
                update = self.client.GetDataValue(1, i)
                item['코드'] = code
                item['종목명'] = name
                cate = self.client.GetDataValue(2, i)
                if (update == ord('c')):
                    newcancel = '[취소]'
                if cate in self.diccode:
                    item['특이사항'] = newcancel + self.diccode[cate]
                else:
                    item['특이사항'] = newcancel + ''

                self.caller.listWatchData.insert(0, item)
                print(item)

        # 실시간 처리 - marketnews : 뉴스 및 공시 정보
        elif self.name == 'marketnews':
            item = {}
            update = self.client.GetHeaderValue(0)
            cont = ''
            if update == ord('D'):
                cont = '[삭제]'
            code = item['코드'] = self.client.GetHeaderValue(1)
            time = self.client.GetHeaderValue(2)
            h, m = divmod(time, 100)
            item['시간'] = '%02d:%02d' % (h, m)
            item['종목명'] = name = g_objCodeMgr.CodeToName(code)
            cate = self.client.GetHeaderValue(4)
            item['특이사항'] = cont + self.client.GetHeaderValue(5)
            print(item)
            self.caller.listWatchData.insert(0, item)


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


# CpPBMarkeWatch:
class CpPBMarkeWatch(CpPublish):
    def __init__(self):
        super().__init__('marketwatch', 'CpSysDib.CpMarketWatchS')


# CpPBMarkeWatch:
class CpPB8092news(CpPublish):
    def __init__(self):
        super().__init__('marketnews', 'Dscbo1.CpSvr8092S')


# CpRpMarketWatch : 특징주 포착 통신
class CpRpMarketWatch:
    def __init__(self):
        self.objStockMst = win32com.client.Dispatch('CpSysDib.CpMarketWatch')
        self.objpbMarket = CpPBMarkeWatch()
        self.objpbNews = CpPB8092news()
        return

    def Request(self, code, caller):
        self.objpbMarket.Unsubscribe()
        self.objpbNews.Unsubscribe()

        self.objStockMst.SetInputValue(0, code)
        # 1: 종목 뉴스 2: 공시정보 10: 외국계 창구첫매수, 11:첫매도 12 외국인 순매수 13 순매도
        rqField = '1,2,10,11,12,13'
        self.objStockMst.SetInputValue(1, rqField)
        self.objStockMst.SetInputValue(2, 0)  # 시작 시간: 0 처음부터

        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        cnt = self.objStockMst.GetHeaderValue(2)  # 수신 개수
        for i in range(cnt):
            item = {}

            time = self.objStockMst.GetDataValue(0, i)
            h, m = divmod(time, 100)
            item['시간'] = '%02d:%02d' % (h, m)
            item['코드'] = self.objStockMst.GetDataValue(1, i)
            item['종목명'] = g_objCodeMgr.CodeToName(item['코드'])
            cate = self.objStockMst.GetDataValue(3, i)
            item['특이사항'] = self.objStockMst.GetDataValue(4, i)
            print(item)
            caller.listWatchData.append(item)

        self.objpbMarket.Subscribe(code, caller)
        self.objpbNews.Subscribe(code, caller)

        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        self.listWatchData = []
        self.objMarketWatch = CpRpMarketWatch()

        self.setWindowTitle("특징주 포착(#8092 Market-Watch")
        self.setGeometry(300, 300, 300, 180)

        nH = 20

        btnPrint = QPushButton('Print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        btnExcel = QPushButton('Excel 내보내기', self)
        btnExcel.move(20, nH)
        btnExcel.clicked.connect(self.btnExcel_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        self.objMarketWatch.Request('*', self)

    def btnPrint_clicked(self):
        for item in self.listWatchData:
            print(item)
        return

    def btnExcel_clicked(self):

        if (len(self.listWatchData) == 0):
            print('데이터 없음')
            return

        df = pd.DataFrame(columns=['시간', '코드', '종목명', '특이사항'])

        for item in self.listWatchData:
            df.loc[len(df)] = item

        writer = pd.ExcelWriter(gExcelFile, engine='xlsxwriter')
        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name='Sheet1')
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        os.startfile(gExcelFile)
        return

    def btnExit_clicked(self):
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
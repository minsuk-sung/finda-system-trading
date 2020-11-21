# 대신증권 API
# 매매입체분석(투자주체별현황) 예제
# 매매입체분석(CpSysDib.CpSvr7254) 서비스를 이용하여 데이터를 조회하는 파이썬 예제입니다
# 조회 후 엑셀 내보내기를 통해 엑셀 파일로 데이터 확인 가능합니다.

import sys
from PyQt5.QtWidgets import *
import win32com.client
from pandas import Series, DataFrame
import pandas as pd
import locale
import os
import time

locale.setlocale(locale.LC_ALL, '')
# cp object
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

gExcelFile = '7254.xlsx'


class CpRp7354:
    def Request(self, code, caller):
        # 연결 여부 체크
        objCpCybos = win32com.client.Dispatch('CpUtil.CpCybos')
        bConnect = objCpCybos.IsConnect
        if (bConnect == 0):
            print('PLUS가 정상적으로 연결되지 않음. ')
            return False

        # 관심종목 객체 구하기
        objRq = win32com.client.Dispatch('CpSysDib.CpSvr7254')
        objRq.SetInputValue(0, code)
        objRq.SetInputValue(1, 6)  # 일자별
        objRq.SetInputValue(4, ord('0'))  # '0' 순매수 '1' 매매비중
        objRq.SetInputValue(5, 0)  # '전체
        objRq.SetInputValue(6, ord('1'))  # '1' 순매수량 '2' 추정금액(백만)

        sumcnt = 0
        caller.data7254 = None
        caller.data7254 = pd.DataFrame(columns=('date', 'close', '개인', '외국인', '기관계',
                                                '금융투자', '보험', '투신', '은행', '기타금융', '연기금', '국가,지자체',
                                                '기타법인', '기타외인'))

        while True:
            remainCount = g_objCpStatus.GetLimitRemainCount(1)  # 1 시세 제한
            if remainCount <= 0:
                print('시세 연속 조회 제한 회피를 위해 sleep', g_objCpStatus.LimitRequestRemainTime)
                time.sleep(g_objCpStatus.LimitRequestRemainTime / 1000)

            objRq.BlockRequest()

            # 현재가 통신 및 통신 에러 처리
            rqStatus = objRq.GetDibStatus()
            print('통신상태', rqStatus, objRq.GetDibMsg1())
            if rqStatus != 0:
                return False

            cnt = objRq.GetHeaderValue(1)
            sumcnt += cnt

            for i in range(cnt):
                item = {}
                item['date'] = objRq.GetDataValue(0, i)
                item['close'] = objRq.GetDataValue(14, i)
                item['개인'] = objRq.GetDataValue(1, i)
                item['외국인'] = objRq.GetDataValue(2, i)
                item['기관계'] = objRq.GetDataValue(3, i)
                item['금융투자'] = objRq.GetDataValue(4, i)
                item['보험'] = objRq.GetDataValue(5, i)
                item['투신'] = objRq.GetDataValue(6, i)
                item['은행'] = objRq.GetDataValue(7, i)
                item['기타금융'] = objRq.GetDataValue(8, i)
                item['연기금'] = objRq.GetDataValue(9, i)
                item['국가,지자체'] = objRq.GetDataValue(13, i)
                item['기타법인'] = objRq.GetDataValue(10, i)
                item['기타외인'] = objRq.GetDataValue(11, i)

                caller.data7254.loc[len(caller.data7254)] = item

            # 1000 개 정도만 처리
            if sumcnt > 1000:
                break;
            # 연속 처리
            if objRq.Continue != True:
                break

        caller.data7254 = caller.data7254.set_index('date')
        # 인덱스 이름 제거
        caller.data7254.index.name = None
        print(caller.data7254)
        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('PLUS API TEST')
        self.setGeometry(300, 300, 300, 240)
        self.isSB = False
        self.objCur = []

        self.data7254 = DataFrame()

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

    def btnStart_clicked(self):
        # 요청 필드 배열 - 종목코드, 시간, 대비부호 대비, 현재가, 거래량, 종목명
        obj7254 = CpRp7354()
        obj7254.Request('A000660', self)

    def btnExcel_clicked(self):
        print(len(self.data7254.index))
        # create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(gExcelFile, engine='xlsxwriter')
        # Convert the dataframe to an XlsxWriter Excel object.
        self.data7254.to_excel(writer, sheet_name='Sheet1')
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        os.startfile(gExcelFile)
        return

    def btnPrint_clicked(self):
        print(self.data7254)

    def btnExit_clicked(self):
        exit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
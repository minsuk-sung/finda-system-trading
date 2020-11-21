# 대신증권 API
# 종목별 투자자 매매동향 (잠정)데이터

# CpSysDib.CpSvr7210d 를 통해 "종목별 투자자 매매동향(잠정)데이터" 를 조회 합니다.
#
# ■ 사용된 PLUS OBJECT
#    - CpSysDib.CpSvr7210d - 투자자 매매 동향(잠정) 데이터 조회
#
# ■ 화면 설명
#   - 기관계 상위 :  기간계 상위 순으로 조회
#   - 외국인상위 :  외국인상위 순으로 조회
#   - print : 조회 내용 print
#   - 엑셀 내보내기 : 조회 내용 엑셀로 내보내기

import sys
from PyQt5.QtWidgets import *
import win32com.client
import pandas as pd
import os

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


# Cp7210 : 종목별 투자자 매매동향(잠정)데이터
class Cp7210:
    def __init__(self):
        self.objRq = None
        return

    def request(self, investFlag, caller):
        maxRqCont = 1
        rqCnt = 0
        caller.data7210 = []
        self.objRq = None
        self.objRq = win32com.client.Dispatch("CpSysDib.CpSvr7210d")

        while True:
            self.objRq.SetInputValue(0, '0')  # 0 전체 1 거래소 2 코스닥 3 업종 4 관심종목
            self.objRq.SetInputValue(1, ord('0'))  # 0 수량 1 금액
            self.objRq.SetInputValue(2, investFlag)  # 0 종목 1 외국인 2 기관계 3 보험기타 4 투신..
            self.objRq.SetInputValue(3, ord('0'))  # 0 상위순 1 하위순

            self.objRq.BlockRequest()
            rqCnt += 1

            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(0)
            date = self.objRq.GetHeaderValue(1)  # 집계날짜
            time = self.objRq.GetHeaderValue(2)  # 집계시간
            print(cnt)

            for i in range(cnt):
                item = {}
                item['code'] = self.objRq.GetDataValue(0, i)
                item['종목명'] = self.objRq.GetDataValue(1, i)
                item['현재가'] = self.objRq.GetDataValue(2, i)
                item['대비'] = self.objRq.GetDataValue(3, i)
                item['대비율'] = self.objRq.GetDataValue(4, i)
                item['거래량'] = self.objRq.GetDataValue(5, i)
                item['외국인'] = self.objRq.GetDataValue(6, i)
                item['기관계'] = self.objRq.GetDataValue(7, i)
                item['보험기타금융'] = self.objRq.GetDataValue(8, i)
                item['투신'] = self.objRq.GetDataValue(9, i)
                item['은행'] = self.objRq.GetDataValue(10, i)
                item['연기금'] = self.objRq.GetDataValue(11, i)
                item['국가지자체'] = self.objRq.GetDataValue(12, i)
                item['기타법인'] = self.objRq.GetDataValue(13, i)

                caller.data7210.append(item)

            if rqCnt >= maxRqCont:
                break

            if self.objRq.Continue == False:
                break
        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("종목별 투자자 매매동향(잠정)")
        self.setGeometry(300, 300, 300, 180)

        self.obj7210 = Cp7210()
        self.data7210 = []

        nH = 20
        btnOpt1 = QPushButton('기관계 상위', self)
        btnOpt1.move(20, nH)
        btnOpt1.clicked.connect(self.btnOpt1_clicked)
        nH += 50

        btnOpt2 = QPushButton('외국인 상위', self)
        btnOpt2.move(20, nH)
        btnOpt2.clicked.connect(self.btnOpt2_clicked)
        nH += 50

        btnPrint = QPushButton('print', self)
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
        nH += 50
        self.setGeometry(300, 300, 300, nH)

    # 기관계 상위
    def btnOpt1_clicked(self):
        self.obj7210.request(2, self)
        return

    # 외국인 상위
    def btnOpt2_clicked(self):
        self.obj7210.request(1, self)
        return

    def btnPrint_clicked(self):
        for item in self.data7210:
            print(item)
        return

    # 엑셀 내보내기
    def btnExcel_clicked(self):
        excelfile = '7210.xlsx'
        df = pd.DataFrame(columns=['code', '종목명', '현재가', '대비', '대비율', '거래량', '외국인', '기관계',
                                   '보험기타금융', '투신', '은행', '연기금', '국가지자체', '기타법인'])

        for item in self.data7210:
            df.loc[(len(df))] = item

        # create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter(excelfile, engine='xlsxwriter')
        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name='Sheet1')
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()
        os.startfile(excelfile)
        return

    def btnExit_clicked(self):
        exit()
        return

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
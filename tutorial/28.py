# 대신증권 API
# 종목검색 서비스를 이용하여 예제 전략 조건 조회하기 예제

# HTS 의 종목검색(#8537) 서비스를 이용하면 원하는 조건에 해당하는 종목을 바로 바로 찾아 볼 수 있습니다.
# 이번 예제는 PLUS 에서 #8537 의 예제 전략을 가져와 조회 하는 예제입니다
#
# ■ 사용된 PLUS OBJECT:
#   - CpSysDib.CssStgList : 전략 리스트 조회 (예제 또는 사용자 전략 선택 가능)
#   - CpSysDib.CssStgFind : 특정 전략조건에 해당하는 종목 리스트 조회

import sys
from PyQt5.QtWidgets import *
import win32com.client
import pandas as pd
import os

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


# Cp8537 : 종목검색 전략 조회
class Cp8537:
    def __init__(self):
        self.objRq = None
        return

    def requestList(self, caller):
        caller.data8537 = {}
        self.objRq = None
        self.objRq = win32com.client.Dispatch("CpSysDib.CssStgList")

        # 예제 전략에서 전략 리스트를 가져옵니다.
        self.objRq.SetInputValue(0, ord('0'))  # '0' : 예제전략, '1': 나의전략
        self.objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False

        cnt = self.objRq.GetHeaderValue(0)  # 0 - (long) 전략 목록 수
        flag = self.objRq.GetHeaderValue(1)  # 1 - (char) 요청구분
        print('종목검색 전략수:', cnt)

        for i in range(cnt):
            item = {}
            item['전략명'] = self.objRq.GetDataValue(0, i)
            item['ID'] = self.objRq.GetDataValue(1, i)
            item['전략등록일시'] = self.objRq.GetDataValue(2, i)
            item['작성자필명'] = self.objRq.GetDataValue(3, i)
            item['평균종목수'] = self.objRq.GetDataValue(4, i)
            item['평균승률'] = self.objRq.GetDataValue(5, i)
            item['평균수익'] = self.objRq.GetDataValue(6, i)
            caller.data8537[item['전략명']] = item
            print(item)

        return True

    def requestStgID(self, id, caller):
        caller.dataStg = []
        self.objRq = None
        self.objRq = win32com.client.Dispatch("CpSysDib.CssStgFind")
        self.objRq.SetInputValue(0, id)  # 전략 id 요청
        self.objRq.BlockRequest()
        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        if rqStatus != 0:
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            return False

        cnt = self.objRq.GetHeaderValue(0)  # 0 - (long) 검색된 결과 종목 수
        totcnt = self.objRq.GetHeaderValue(1)  # 1 - (long) 총 검색 종목 수
        stime = self.objRq.GetHeaderValue(2)  # 2 - (string) 검색시간
        print('검색된 종목수:', cnt, '전체종목수:', totcnt, '검색시간:', stime)

        for i in range(cnt):
            item = {}
            item['code'] = self.objRq.GetDataValue(0, i)
            item['종목명'] = g_objCodeMgr.CodeToName(item['code'])
            caller.dataStg.append(item)

        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("종목검색 예제")
        self.setGeometry(300, 300, 500, 180)

        self.obj8537 = Cp8537()
        self.data8537 = {}
        self.dataStg = []

        nH = 20
        btnOpt1 = QPushButton('전략리스트 조회', self)
        btnOpt1.move(20, nH)
        btnOpt1.clicked.connect(self.btnOpt1_clicked)
        nH += 50

        self.comboStg = QComboBox(self)
        self.comboStg.move(20, nH)
        self.comboStg.currentIndexChanged.connect(self.comboChanged)
        self.comboStg.resize(400, 30)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50
        self.setGeometry(300, 300, 500, nH)

        self.btnOpt1_clicked()

    # 전략리스트 조회
    def btnOpt1_clicked(self):
        self.obj8537.requestList(self)

        for k, v in self.data8537.items():
            self.comboStg.addItem(k)
        return

    def comboChanged(self):
        cur = self.comboStg.currentText()
        print(cur)
        self.requestStgID(cur)

    def requestStgID(self, stgName):
        item = self.data8537[stgName]
        id = item['ID']
        name = item['전략명']

        self.obj8537.requestStgID(id, self)

        print('검색전략:', id, '전략명:', name, '검색종목수:', len(self.dataStg))
        for item in self.dataStg:
            print(item)
        return

    def btnExit_clicked(self):
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
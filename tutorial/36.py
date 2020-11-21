# 대신증권 API
# 선물 분 차트 그리기(matplotlib 이용)
# 선물 분 차트를 조회하고 이를 차트로 그리는 간단한 예제 입니다.
# ※ 주의사항: 본 예제는 PLUS 활용을 돕기 위해 예제로만 제공됩니다.

import datetime
import sys
import ctypes
import time

import numpy as np
from PyQt5.QtWidgets import *
import win32com.client
import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.finance as matfin
import matplotlib.ticker as ticker

C_DT = 0  # 일자
C_TM = 1  # 시간
C_OP = 2  # 시가
C_HP = 3  # 고가
C_LP = 4  # 저가
C_CP = 5  # 종가
C_VL = 6  # 거래량
C_MA5 = 7  # 5일 이동평균
C_MA10 = 8  # 10일 이동평균
C_MA20 = 9  # 20일 이동평균

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
g_objFutureMgr = win32com.client.Dispatch("CpUtil.CpFutureCode")


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


class CpFutureChart:
    def __init__(self):
        self.objFutureChart = win32com.client.Dispatch("CpSysDib.FutOptChart")

    # 차트 요청 - 분간, 틱 차트
    def RequestMT(self, code, dwm, count, caller):
        # 연결 여부 체크
        bConnect = g_objCpStatus.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        self.objFutureChart.SetInputValue(0, code)  # 종목코드
        self.objFutureChart.SetInputValue(1, ord('2'))  # 개수로 받기
        self.objFutureChart.SetInputValue(4, count)  # 조회 개수
        self.objFutureChart.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])  # 요청항목 - 날짜, 시간,시가,고가,저가,종가,거래량
        self.objFutureChart.SetInputValue(6, dwm)  # '차트 주기 - 분/틱
        self.objFutureChart.SetInputValue(7, 1)  # 분틱차트 주기
        self.objFutureChart.SetInputValue(8, ord('0'))  # 갭보정
        self.objFutureChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objFutureChart.BlockRequest()

        rqStatus = self.objFutureChart.GetDibStatus()
        rqRet = self.objFutureChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        len = self.objFutureChart.GetHeaderValue(3)

        caller.chartData = {}
        caller.chartData[C_DT] = []
        caller.chartData[C_TM] = []
        caller.chartData[C_OP] = []
        caller.chartData[C_HP] = []
        caller.chartData[C_LP] = []
        caller.chartData[C_CP] = []
        caller.chartData[C_VL] = []
        caller.chartData[C_MA5] = []
        caller.chartData[C_MA10] = []
        caller.chartData[C_MA20] = []

        for i in range(len):
            caller.chartData[C_DT].insert(0, self.objFutureChart.GetDataValue(0, i))
            caller.chartData[C_TM].insert(0, self.objFutureChart.GetDataValue(1, i))
            caller.chartData[C_OP].insert(0, self.objFutureChart.GetDataValue(2, i))
            caller.chartData[C_HP].insert(0, self.objFutureChart.GetDataValue(3, i))
            caller.chartData[C_LP].insert(0, self.objFutureChart.GetDataValue(4, i))
            caller.chartData[C_CP].insert(0, self.objFutureChart.GetDataValue(5, i))
            caller.chartData[C_VL].insert(0, self.objFutureChart.GetDataValue(6, i))

        # print(self.objFutureChart.Continue)

        return


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        # 기본 변수들
        self.chartData = {}
        self.code = ''
        self.isRq = False

        self.objChart = CpFutureChart()

        self.sizeControl()
        # 선물 종목 코드 추가
        for i in range(g_objFutureMgr.GetCount()):
            code = g_objFutureMgr.GetData(0, i)
            self.comboStg.addItem(code)
        self.comboStg.setCurrentIndex(0)

    def sizeControl(self):
        # 윈도우 버튼 배치
        self.setWindowTitle("PLUS API TEST")

        self.setGeometry(50, 50, 1200, 600)
        self.comboStg = QComboBox()
        # self.comboStg.move(20, nH)
        self.comboStg.currentIndexChanged.connect(self.comboChanged)
        # self.comboStg.resize(100, 30)
        self.label = QLabel('종목코드')
        # self.label.move(140, nH)

        # Figure 를 먼저 만들고 레이아웃에 들어 갈 sub axes 를 생성 한다.
        self.fig = plt.Figure()
        self.canvas = FigureCanvas(self.fig)

        # top layout
        topLayout = QHBoxLayout()
        topLayout.addWidget(self.comboStg)
        topLayout.addWidget(self.label)
        topLayout.addStretch(1)
        # topLayout.addSpacing(20)

        chartlayout = QVBoxLayout()
        chartlayout.addWidget(self.canvas)

        layout = QVBoxLayout()
        layout.addLayout(topLayout)
        layout.addLayout(chartlayout)
        layout.setStretchFactor(topLayout, 0)
        layout.setStretchFactor(chartlayout, 1)

        self.setLayout(layout)

    # 분차트 받기
    def RequestMinchart(self):
        if self.objChart.RequestMT(self.code, ord('m'), 100, self) == False:
            exit()

    def makeMovingAverage(self, maData, interval):
        # maData = []
        for i in range(0, len(self.chartData[C_DT])):
            if (i < interval):
                maData.append(float('nan'))
                continue
            sum = 0
            for j in range(0, interval):
                sum += self.chartData[C_CP][i - j]
            ma = sum / interval
            maData.append(ma)
        # print(maData)

    def drawMinchart(self):
        # 기존 거 지운다.
        self.fig.clf()

        # 211 - 2(행) * 1(열) 배치 1번째
        self.ax1 = self.fig.add_subplot(2, 1, 1)
        # 212 - 2(행) * 1(열) 배치 2번째
        self.ax2 = self.fig.add_subplot(2, 1, 2)

        ###############################################
        # 봉차트 그리기
        # self.ax1.xaxis.set_major_formatter(ticker.FixedFormatter(schartData[C_TM]))
        matfin.candlestick2_ohlc(self.ax1, self.chartData[C_OP], self.chartData[C_HP], self.chartData[C_LP],
                                 self.chartData[C_CP],
                                 width=0.8, colorup='r', colordown='b')

        ###############################################
        # x 축 인덱스 만들기 - 기본 순차 배열 추가
        x_tick_raw = [i for i in range(len(self.chartData[C_DT]))]
        # x 축 인덱스 만들기 - 실제 화면에 표시될 텍스트 만들기
        x_tick_labels = []

        startDate = 0
        dateChanged = True
        for i in range(len(self.chartData[C_DT])):
            # 날짜 변경 된 경우 날짜 정보 저장
            date = self.chartData[C_DT][i]
            if (date != startDate):
                yy, mm = divmod(date, 10000)
                mm, dd = divmod(mm, 100)
                sDate = '%2d/%d ' % (mm, dd)
                print(sDate)
                startDate = date
                dateChanged = True

            # 0 분 또는 30분 단위로 시간 표시
            hhh, mmm = divmod(self.chartData[C_TM][i], 100)
            stime = '%02d:%02d' % (hhh, mmm)
            if (mmm == 0 or mmm == 30):
                if dateChanged == True:
                    sDate += stime
                    x_tick_labels.append(sDate)
                    dateChanged = False
                else:
                    x_tick_labels.append(stime)
            else:
                x_tick_labels.append('')

        ###############################################
        # 이동 평균 그리기
        self.ax1.plot(x_tick_raw, self.chartData[C_MA5], label='ma5')
        self.ax1.plot(x_tick_raw, self.chartData[C_MA10], label='ma10')
        self.ax1.plot(x_tick_raw, self.chartData[C_MA20], label='ma20')

        ###############################################
        # 거래량 그리기
        self.ax2.bar(x_tick_raw, self.chartData[C_VL])

        ###############################################
        # x 축 가로 인덱스 지정
        self.ax1.set(xticks=x_tick_raw, xticklabels=x_tick_labels)
        self.ax2.set(xticks=x_tick_raw, xticklabels=x_tick_labels)

        self.ax1.grid()
        self.ax2.grid()
        plt.tight_layout()
        self.ax1.legend(loc='upper left')

        self.canvas.draw()

    def comboChanged(self):
        if self.isRq == True:
            return
        self.isRq = True
        self.code = self.comboStg.currentText()
        self.name = g_objFutureMgr.CodetoName(self.code)
        self.label.setText(self.name)
        self.RequestMinchart()
        self.makeMovingAverage(self.chartData[C_MA5], 5)
        self.makeMovingAverage(self.chartData[C_MA10], 10)
        self.makeMovingAverage(self.chartData[C_MA20], 20)
        self.drawMinchart()
        self.isRq = False

        # self.requestStgID(cur)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
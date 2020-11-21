# 대신증권 API
# MACD 차트지표 계산(실시간)
#
# 클래스 설명
#   - class CpStockChart: 차트 데이터 수신
#   - class CpEvent: 실시간 시세 수신
#   - class MyWindow : 기본 UI 클래스
#
# 예제에서 사용한 PLUS OBJECT
#   - CpSysDib.StockChart : 차트 기본 데이터 수신(일주월분틱 별  시가, 고가, 저가, 고가, 거래량 등)
#   - DsCbo1.StockCur : 현재가 실시간 데이터 수신
#   - CpIndexes.CpSeries : 차트 지표 계산을 위한 중간 OBJECT
#   - CpIndexes.CpIndex : 실질적인 지표 계산 OBJECT 로 MACD, 이동평균, CCI, STOCHASTIC 등 다양한 지표 계산 및 시그널 발생도 가능
#
# PLUS API 를 사용하여 MACD 지표 구하는 기본 적인 방법
#   1. 차트 데이터 통신
#   2. 수신 받은 차트 데이터를 CpIndexes.CpSeries 에 역순으로 입력
#   3. CpIndexes.CpIndex 를 이용하여 MACD 계산
#
# PLUS API 를 사용하여 MACD 실시간 update 하기
#   1. DsCbo1.StockCur 를 이용하여 실시간 현재가 수신
#   2. 실시간 현재가 정보를 CpIndexes.CpSeries 에 update (update 함수)
#   3. CpIndexes.CpIndex 에서 재 계산(update 함수)

import sys
from PyQt5.QtWidgets import *
import win32com.client


# 요약: MACD 지표 데이터 실시간 구하기
#     : 차트 OBJECT 를 통해 차트 데이터를 받은 후
#     : 지표 실시간 계산 OBJECT 를 통해 지표 데이터를 계산

class CpEvent:
    def set_params(self, client, objCaller):
        self.client = client
        self.caller = objCaller

    def OnReceived(self):
        code = self.client.GetHeaderValue(0)  # 초
        name = self.client.GetHeaderValue(1)  # 초
        timess = self.client.GetHeaderValue(18)  # 초
        exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
        cprice = self.client.GetHeaderValue(13)  # 현재가
        diff = self.client.GetHeaderValue(2)  # 대비
        cVol = self.client.GetHeaderValue(17)  # 순간체결수량
        vol = self.client.GetHeaderValue(9)  # 거래량
        open = self.client.GetHeaderValue(4)  # 고가
        high = self.client.GetHeaderValue(5)  # 고가
        low = self.client.GetHeaderValue(6)  # 저가

        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", name, timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
            return  # 차트는 예상 체결 시간 update 없음.
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", name, timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)

        # MACD 지표 update 함수 호출
        self.caller.updateMACD(cprice, open, high, low, vol)


class CpStockChart:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

    def Request(self, code, objCaller):
        # 연결 여부 체크
        bConnect = self.objCpCybos.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False

        print("111")
        # 현재가 객체 구하기
        self.objStockChart.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
        self.objStockChart.SetInputValue(4, 500)  # 최근 500일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        # MACD 지표 계산 함수 호출
        objCaller.makeChartSeries(self.objStockChart)


class CpStockCur:
    def Subscribe(self, code, objIndex):
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        handler = win32com.client.WithEvents(self.objStockCur, CpEvent)
        self.objStockCur.SetInputValue(0, code)
        handler.set_params(self.objStockCur, objIndex)
        self.objStockCur.Subscribe()

    def Unsubscribe(self):
        self.objStockCur.Unsubscribe()


class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 300, 150)
        self.isSB = False
        self.objCur = []

        btnStart = QPushButton("요청 시작", self)
        btnStart.move(20, 20)
        btnStart.clicked.connect(self.btnStart_clicked)

        btnStop = QPushButton("요청 종료", self)
        btnStop.move(20, 70)
        btnStop.clicked.connect(self.btnStop_clicked)

        btnExit = QPushButton("종료", self)
        btnExit.move(20, 120)
        btnExit.clicked.connect(self.btnExit_clicked)

        # obj 미리 선언

    def StopSubscribe(self):
        if self.isSB:
            cnt = len(self.objCur)
            for i in range(cnt):
                self.objCur[i].Unsubscribe()
            print(cnt, "종목 실시간 해지되었음")
        self.isSB = False

        self.objCur = []

    def btnStart_clicked(self):
        self.StopSubscribe();

        # 요청 종목
        code = "A000660"

        # 지표 계산을 위한 시리즈 선언 - 차트 데이터 수신 받아 데이터를 넣어야 함.
        self.objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")

        # 1. 차트 데이터 통신 요청
        self.objChart = CpStockChart()
        if self.objChart.Request(code, self) == False:
            exit()

        # 2. macd 지표 만들기
        self.makeMACD()

        # 3. 현재가 실시간 요청하기
        self.objCur.append(CpStockCur())
        self.objCur[0].Subscribe(code, self)

        print("빼기빼기================-")
        print("종목 실시간 현재가 요청 시작")
        self.isSB = True

    def btnStop_clicked(self):
        self.StopSubscribe()

    def btnExit_clicked(self):
        self.StopSubscribe()
        exit()

    # 차트 수신 데이터 --> 시리즈 생성
    # 차트 수신 데이터의 경우 최근 데이터가 맨 앞에 있으나
    # 시리즈는 반대로 넣어야 함.
    # 차트 데이터를 가져와 역순으로 시리즈에 넣는 작업 필요
    def makeChartSeries(self, objStockChart):
        len = objStockChart.GetHeaderValue(3)

        print("날짜", "시가", "고가", "저가", "종가", "거래량")
        print("빼기빼기==============================================-")

        for i in range(len):
            day = objStockChart.GetDataValue(0, len - i - 1)
            open = objStockChart.GetDataValue(1, len - i - 1)
            high = objStockChart.GetDataValue(2, len - i - 1)
            low = objStockChart.GetDataValue(3, len - i - 1)
            close = objStockChart.GetDataValue(4, len - i - 1)
            vol = objStockChart.GetDataValue(5, len - i - 1)
            print(day, open, high, low, close, vol)
            # objSeries.Add 종가, 시가, 고가, 저가, 거래량, 코멘트
            self.objSeries.Add(close, open, high, low, vol)
        return

    # CpIndex 를 이용하여 MACD 지표 계산
    # MACD 는 총 3가지 지표가 들어 있음(MACD, SIGNAL, OSCILLATOR)
    # 최근 데이터는 지표의 맨 마지막 데이터에 들어 있음.
    def makeMACD(self):
        # 지표 계산 object
        self.objIndex = win32com.client.Dispatch("CpIndexes.CpIndex")
        self.objIndex.series = self.objSeries
        self.objIndex.put_IndexKind("MACD")  # 계산할 지표: MACD
        self.objIndex.put_IndexDefault("MACD")  # MACD 지표 기본 변수 자동 세팅

        print("MACD 변수", self.objIndex.get_Term1(), self.objIndex.get_Term2(), self.objIndex.get_Signal())

        # 지표 데이터 계산 하기
        self.objIndex.Calculate()

        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex)
        indexName = ["MACD", "SIGNAL", "OSCILLATOR"]
        for index in range(cntofIndex):
            cnt = self.objIndex.GetCount(index)
            # for j in range(cnt) :
            #    value = self.objIndex.GetResult(index,j)
            value = self.objIndex.GetResult(index, cnt - 1)  # 지표의 가장 최근 값 - 맨 뒤 데이터
            print(indexName[index], value)  # 지표의 최근 값 표시

    # 실시간 시세 수신 받아 MACD 계산
    def updateMACD(self, cprice, open, high, low, vol):
        # 지표 데이터 update
        self.objSeries.update(cprice, open, high, low, vol)
        self.objIndex.update()
        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex)

        indexName = ["MACD", "SIGNAL", "OSCILLATOR"]

        for index in range(cntofIndex):
            cnt = self.objIndex.GetCount(index)
            # print(index , "번째 지표의 데이터 개수", cnt)
            value = self.objIndex.GetResult(index, cnt - 1)  # 지표의 가장 최근 값 - 맨 뒤 데이터
            print(indexName[index], value)  # 지표의 최근 값 표시

        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
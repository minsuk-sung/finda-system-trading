# 대신증권 API
# 차트 지표 계산(MACD/RSI 등)

# 차트 지표 계산하는 예제 입니다
#
# 예제를 단순화 하기 위해서
# 기본적으로 '하이닉스' 종목 차트 데이터를 가져온 후
# 지정한 지표를 계산하도록 했습니다
#
# ■ 지표 리스트 - PLUS 에서 제공하는 모든 지표들과 기본 변수 값을 리스트 합니다
# ■ 계산할 지표를 선택 - 예제로 이동평균/Stochastic Slow/MACD/RSI 등의 지표를 선택해서 결과를 확인할 수 있습니다
# ■ 지표 조건 변경 방법 - Stochastic Slow의 변수를 12,5,5 식으로 변경 하는 코드 예제입니다.
#
#
# ※ 주의사항: 지표의 계산 값은 지표마다 상이합니다.
# 지표 값을 HTS 차트와 비교 시에는 기간, 주기(일/주/월/분/틱), 수정주가 여부 등 모든 조건이 동일 한 지 반드시 확인 해야 합니다
# 특히 기간에 따라 값이 달라지는 지표가 있어 기간이 동일 한지 확인해야 합니다 지표에 대한 자세한 사항은 본 게시판에서 지원하지 않습니다

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

    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


# 차트 기본 데이터 통신
class CpStockChart:
    def __init__(self):
        self.objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

    def Request(self, code, cnt, objSeries):
        #######################################################
        # 1. 일간 차트 데이터 요청
        self.objStockChart.SetInputValue(0, code)  # 종목 코드 -
        self.objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
        self.objStockChart.SetInputValue(4, cnt)  # 최근 100일치
        self.objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
        self.objStockChart.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.objStockChart.BlockRequest()

        rqStatus = self.objStockChart.GetDibStatus()
        rqRet = self.objStockChart.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        #######################################################
        # 2. 일간 차트 데이터 ==> CpIndexes.CpSeries 로 변환
        len = self.objStockChart.GetHeaderValue(3)

        print("날짜", "시가", "고가", "저가", "종가", "거래량")
        print("==============================================-")
        for i in range(len):
            day = self.objStockChart.GetDataValue(0, len - i - 1)
            open = self.objStockChart.GetDataValue(1, len - i - 1)
            high = self.objStockChart.GetDataValue(2, len - i - 1)
            low = self.objStockChart.GetDataValue(3, len - i - 1)
            close = self.objStockChart.GetDataValue(4, len - i - 1)
            vol = self.objStockChart.GetDataValue(5, len - i - 1)
            print(day, open, high, low, close, vol)
            # objSeries.Add 종가, 시가, 고가, 저가, 거래량, 코멘트
            objSeries.Add(close, open, high, low, vol)
        print("==============================================-")

        return

# 지표 계산 관리 클래스
class CpIndexTest:
    def __init__(self):
        # 1. 차트 데이터 통신 요청
        self.objChart = CpStockChart()
        # CpIndexes.CpSeries : 차트 기본 데이터 관리 PLUS 객체
        self.objSeries = win32com.client.Dispatch("CpIndexes.CpSeries")
        # CpIndexes.CpIndex : 지표 계산을 담당하는 PLUS 객체
        self.objIndex = win32com.client.Dispatch("CpIndexes.CpIndex")

        # 테스트를 위해 하이닉스 종목으로 미리 차트 데이터 구성
        self.objChart.Request('A000660', 100, self.objSeries)

    # 플러스에서 제공하는 모든 지표와 지표의 기본 조건 나열
    def listAllIndex(self):
        allIndexList = []
        for i in range(0,6) :
            indexlist = self.objIndex.GetChartIndexCodeListByIndex(i)
            #print(indexlist)
            allIndexList += indexlist


        for index in allIndexList:
            self.objIndex.put_IndexKind(index)
            self.objIndex.put_IndexDefault(index)

            print('%s 변수1 %d 변수2 %d 변수3 %d 변수4 %d Signal %d'
                  %(index, self.objIndex.get_Term1(), self.objIndex.get_Term2(), self.objIndex.get_Term3(),
                  self.objIndex.get_Term4(), self.objIndex.get_Signal()))


    # 주어진 지표의 이름(indexName) 으로 지표 계산 및 데이터 리턴
    def makeIndex(self, indexName, lineName, chartvalue):
        self.objIndex.series = self.objSeries
        self.objIndex.put_IndexKind(indexName)  # 계산할 지표:
        self.objIndex.put_IndexDefault(indexName)  # 지표 기본 변수 자동 세팅

        condList = {}
        condList['조건1'] = self.objIndex.get_Term1()
        condList['조건2'] = self.objIndex.get_Term2()
        condList['조건3'] = self.objIndex.get_Term3()
        condList['조건4'] = self.objIndex.get_Term4()
        condList['Signal'] = self.objIndex.get_Signal()

        condNameList = ['조건1', '조건2', '조건3', '조건4', 'Signal']

        indexcond = indexName
        for conName in condNameList:
            if (condList[conName] != 0) :
                indexcond += ' ' + conName + ' ' + str(condList[conName])

        print(indexcond)


        # 지표 데이터 계산 하기
        self.objIndex.Calculate()

        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex)
        # 지표의 각 라인 이름은 HTS 차트의 각 지표 조건 참고
        for index in range(cntofIndex):
            name = lineName[index]
            chartvalue[name] = []
            cnt = self.objIndex.GetCount(index)
            for j in range(cnt) :
                value = self.objIndex.GetResult(index,j)
                chartvalue[name].append(value)



    # 예제를 위해 StochasticSlow 지표의 조건 값을 12,5,5 로 변경
    def makeStochasticSlow(self, chart_value):
        self.objIndex.series = self.objSeries
        self.objIndex.put_IndexKind("Stochastic Slow")  # 계산할 지표: Stochastic Slow
        self.objIndex.put_IndexDefault("Stochastic Slow")  # Stochastic Slow 지표 기본 변수 자동 세팅

        print("Stochastic Slow 변수", self.objIndex.get_Term1(), self.objIndex.get_Term2(), self.objIndex.get_Signal())

        print('지표 조건값 변경')
        self.objIndex.Term1 = 12
        self.objIndex.Term2 = 5
        self.objIndex.Signal = 5
        print("Stochastic Slow 변수", self.objIndex.get_Term1(), self.objIndex.get_Term2(), self.objIndex.get_Signal())
        # 지표 데이터 계산 하기
        self.objIndex.Calculate()

        cntofIndex = self.objIndex.ItemCount
        print("지표 개수:  ", cntofIndex)
        # 지표의 각 라인 이름은 HTS 차트의 각 지표 조건 참고
        indexName = ["SLOW K", "SLOW D"]
        for index in range(cntofIndex):
            name = indexName[index]
            chart_value[name] = []
            cnt = self.objIndex.GetCount(index)
            for j in range(cnt) :
                value = self.objIndex.GetResult(index,j)
                chart_value[name].append(value)


################################################
# 테스트를 위한 메인 화면
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.objCpIndex = CpIndexTest()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        # 지표의 기본 정보 세팅
        self.initIndexInfo()

        self.setWindowTitle("지표테스트")
        self.setGeometry(300, 300, 300, 180)

        nH = 20
        btnPrint = QPushButton('지표리스트', self)
        btnPrint.move(20, nH)
        btnPrint.resize(200, 30)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        labelDesc1 = QLabel('계산할 지표를 선택', self)
        labelDesc1.move(20, nH)
        labelDesc1.resize(200, 30)
        nH += 30


        self.comboStg = QComboBox(self)
        # self.comboStg.move(20, nH)
        for ikey, ivalue in self.indexList.items():
            self.comboStg.addItem(ikey)

        self.comboStg.currentIndexChanged.connect(self.OnComboChanged)
        self.comboStg.move(20, nH)
        self.comboStg.resize(200, 30)
        nH += 50


        labelDesc2 = QLabel('지표 조건 변경 방법 5,3,3 ==> 12,5,5', self)
        labelDesc2.move(20, nH)
        labelDesc2.resize(200, 30)
        nH += 30

        btnStochSlow = QPushButton('STOCHASTIC SLOW 조건변경', self)
        btnStochSlow.move(20, nH)
        btnStochSlow.resize(200, 30)
        btnStochSlow.clicked.connect(self.btnStochSlow_clicked)
        nH += 50


        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        self.setGeometry(300, 300, 300, nH)


    # 지표의 기본 정보 세팅
    def initIndexInfo(self):
        # 지표 기본 정보, 지표 이름 = 지표 라인 리스트
        self.indexList = {}
        self.indexList['지표선택 없음'] =['없음']
        self.indexList['이동평균(라인1개)'] =['이동평균']
        self.indexList['Stochastic Slow'] = ['SLOW K', 'SLOW D']
        self.indexList['MACD'] = ['MACD', 'SIGNAL', 'OSCILLATOR']
        self.indexList['RSI'] = ['RSI', 'SIGNAL']
        self.indexList['Binary Wave MACD'] = ['BWMACD','SIGNAL', 'OSCILLATOR']
        self.indexList['TSF'] = ['TSF','SIGNAL']
        self.indexList['ZigZag'] = ['ZigZag1', 'ZigZag2']
        self.indexList['Bollinger Band'] = ['Bol상', 'Bol하', 'Bol중']

    def OnComboChanged(self):
        indexname = self.comboStg.currentText()
        if (indexname == '지표선택 없음') :
            return
        print(indexname)


        chartData = {}
        lineName = self.indexList[indexname]

        self.objCpIndex.makeIndex(indexname, lineName, chartData)

        for key, values in chartData.items() :
            print(key)
            print(values)

    def btnPrint_clicked(self):
        self.objCpIndex.listAllIndex()


    # 지표 변수 값을 조정함
    def btnStochSlow_clicked(self) :
        chartData = {}
        self.objCpIndex.makeStochasticSlow(chartData)

        for key, values in chartData.items() :
            print(key)
            print(values)


    def btnExit_clicked(self):
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
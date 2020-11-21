# 대신증권 API
# 실시간 분차트 데이터 만들기

# 실시간 시세를 받아, 1분 차트 데이터를 만드는 예제입니다.
# 예제에서는 코스피 200 에 속하는 200 종목을 실시간으로 요청한 후 받은 시세를 기준으로 1분 차트 데이터를 만듭니다.
# (200 종목 분 차트 데이터 생성)
#
# ■ 분 차트  기준 : 14:10분 봉의 겨우 14:09:00~14:09:59초 사이에 들어온 데이터를 하나의 봉으로 만듭니다
# ■ 예제에서는 시세 수신이 있을 경우에만 데이터를 추가하여 분데이터를 만듭니다. (해당 시간 체결이 없을 경우에는 분 데이터를 채우지 않음)
# ■ 과거 데이터 조회 없이 예제 실행 이후 들어온 데이터로만 분 차트를 만드는 예제입니다.
#
# ※ 주의사항: 본 예제는 PLUS 활용을 돕기 위해 예제로만 제공됩니다.
# 또한 제공된 코드에는 장 운영 시간 변경 등의 예외 코드가 반영 되어 있지 않습니다.

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

    # # 주문 관련 초기화
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False

    return True


################################################
# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

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

            if exFlag != ord('2'):
                return

            item = {}
            item['code'] = code
            item['time'] = timess
            item['diff'] = diff
            item['cur'] = cprice
            item['vol'] = vol

            # 현재가 업데이트
            self.caller.updateCurData(item)

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


class CMinchartData:
    def __init__(self):
        self.minDatas = {}
        self.objCur = {}

    def stop(self):
        for k, v in self.objCur.items():
            v.Unsubscribe()

    def addCode(self, code):
        if (code in self.minDatas):
            return

        self.minDatas[code] = []
        self.objCur[code] = CpPBStockCur()
        self.objCur[code].Subscribe(code, self)

    def updateCurData(self, item):
        code = item['code']
        time = item['time']
        cur = item['cur']
        self.makeMinchart(code, time, cur)

    def makeMinchart(self, code, time, cur):
        hh, mm = divmod(time, 10000)
        mm, tt = divmod(mm, 100)
        mm += 1
        if (mm == 60):
            hh += 1
            mm = 0

        hhmm = hh * 100 + mm
        if hhmm > 1530:
            hhmm = 1530
        bFind = False
        minlen = len(self.minDatas[code])
        if (minlen > 0):
            # 0 : 시간 1 : 시가 2: 고가 3: 저가 4: 종가
            if (self.minDatas[code][-1][0] == hhmm):
                item = self.minDatas[code][-1]
                bFind = True
                item[4] = cur
                if (item[2] < cur):
                    item[2] = cur
                if (item[3] > cur):
                    item[3] = cur

        if bFind == False:
            self.minDatas[code].append([hhmm, cur, cur, cur, cur])

        #        print(code, self.minDatas[code])
        return

    def print(self, code):
        print('====================================================-')
        print('분데이터 print', code, g_objCodeMgr.CodeToName(code))
        print('시간,시가,고가,저가,종가')
        for item in self.minDatas[code]:
            hh, mm = divmod(item[0], 100)
            print("%02d:%02d,%d,%d,%d,%d" % (hh, mm, item[1], item[2], item[3], item[4]))


# print(code, self.minDatas[code])


################################################
# 테스트를 위한 메인 화면
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        self.minData = CMinchartData()

        # 코스피 200 종목 가져와 추가
        self.codelist = g_objCodeMgr.GetGroupCodeList(180)
        for code in self.codelist:
            print(code, g_objCodeMgr.CodeToName(code))
            self.minData.addCode(code)

        self.setWindowTitle("주식 분 차트 생성")
        self.setGeometry(300, 300, 300, 180)

        nH = 20

        btnPrint = QPushButton('print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

    def btnPrint_clicked(self):
        for i in range(len(self.codelist)):
            self.minData.print(self.codelist[i])
            if i > 10: break
        return

    def btnExit_clicked(self):
        self.minData.stop()
        exit()
        return


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
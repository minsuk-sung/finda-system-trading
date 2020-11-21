# 대신증권 API
# 체결기준 주식 당일 매매 손익 예제
# 체결가를 기준으로 당일 매매의 평가손익 및 실현손익을 조회할 수 있는 서비스 예제 입니다.
# ■ 사용 서비스
#   - CpTrade.CpTd6032 : 체결기준 주식 당일 매매 손익

import sys
from PyQt5.QtWidgets import *
import win32com.client
import ctypes

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')


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


# Cp6032 : 주식 잔고 조회
class Cp6032:
    def __init__(self):
        acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = g_objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        print(acc, accFlag[0])

        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6032")
        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째

    # 실제적인 6032 통신 처리
    def request6032(self, caller):
        sumJango = 0
        sumSellM = 0
        sumRate = 0.0

        bIsFist = True
        while True:
            self.objRq.BlockRequest()
            # 통신 및 통신 에러 처리
            rqStatus = self.objRq.GetDibStatus()
            rqRet = self.objRq.GetDibMsg1()
            print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                return False

            cnt = self.objRq.GetHeaderValue(0)
            print('데이터 조회 개수', cnt)

            # 헤더 정보는 한번만 처리
            if bIsFist == True:
                sumJango = self.objRq.GetHeaderValue(1)
                sumSellM = self.objRq.GetHeaderValue(2)
                sumRate = self.objRq.GetHeaderValue(3)
                print('잔량평가손익', sumJango, '매도실현손익', sumSellM, '수익률', sumRate)
                bIsFist = False

            for i in range(cnt):
                item = {}
                item['종목코드'] = self.objRq.GetDataValue(12, i)  # 종목코드
                item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
                item['신용일자'] = self.objRq.GetDataValue(1, i)
                item['전일잔고'] = self.objRq.GetDataValue(2, i)
                item['금일매수수량'] = self.objRq.GetDataValue(3, i)
                item['금일매도수량'] = self.objRq.GetDataValue(4, i)
                item['금일잔고'] = self.objRq.GetDataValue(5, i)
                item['평균매입단가'] = self.objRq.GetDataValue(6, i)
                item['평균매도단가'] = self.objRq.GetDataValue(7, i)
                item['현재가'] = self.objRq.GetDataValue(8, i)
                item['잔량평가손익'] = self.objRq.GetDataValue(9, i)
                item['매도실현손익'] = self.objRq.GetDataValue(10, i)
                item['수익률'] = self.objRq.GetDataValue(11, i)

                print(item)
                caller.data6032.append(item)
            if (self.objRq.Continue == False):
                break
        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # plus 상태 체크
        if InitPlusCheck() == False:
            exit()

        self.setWindowTitle("체결기준 주식당일매매손익 예제")
        self.setGeometry(300, 300, 300, 180)

        # 6033 잔고 object
        self.obj6032 = Cp6032()
        self.data6032 = []

        nH = 20

        btnPrint = QPushButton('Print', self)
        btnPrint.move(20, nH)
        btnPrint.clicked.connect(self.btnPrint_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        # 잔고 요청
        self.request6032()

    def request6032(self):
        if self.obj6032.request6032(self) == False:
            return

    def btnPrint_clicked(self):
        print('체결기준 당일매매손익')
        for item in self.data6032:
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
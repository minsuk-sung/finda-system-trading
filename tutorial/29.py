# 대신증권 API
# 주식 IOC/FOK 주문 테스트 예제

# ※ IOC 및 FOK 주문
#   >> IOC (Immediate - Or-Cancel Order) : 주문즉시 체결 그리고 잔량 자동취소
#     호가접수시점에서 호가한 수량 중 매매계약을 체결할 수 있는 수량에 대하여는 매매거래를 성립시키고,
#     매매계약이 체결되지 아니한 수량은 취소하는 조건
#
#  >> FOK ( Fill- Or- kill Order) : 주문즉시 전부체결 또는 전부자동취소
#     호가의 접수시점에서 호가한 수량의 전부에 대햐여 매매계약을 체결할 수 있는 경우에는 매매거래를
#     성립시키고 그러하지 아니한 경우에는 당해수량의 전부를 취소하는 조건
#
#  ☞ 보통(지정가), 시장가, 최유리지정가에 한해서만 주문조건 부여가 가능합니다.

# IOC 및 FOK 주문의 경우 주문 즉시 체결 또는 취소가 되는 주문으로 아래 예제 코드에서는
# 실시간 주문  체결에서 이를 확인할 수 있도록 로그를 남기고 있습니다.
#
# 신규 IOC 매수 주문을 낸 경우, 실시간 주문 체결은 접수
# -> 취소 확인 순(체결이 되지 않을 경우) 으로 실시간 주문 체결이 오는 것을 아래 로그에서 확인할 수 있습니다.
#
# 신규 매수 주문조건구분: IOC 종목코드: A008800 가격: 474 수량: 1
# 통신상태 0 10766 매수주문이 접수되었습니다.(ordss.cststkord)
#
# conclution
# {'체결플래그': '접수', '주문번호': 24337, '주문수량': 1, '주문가격': 474, '원주문': 0,
# '종목코드': 'A008800', '종목명': '행남생활건강', '매수매도': '매수', '신용대출': '해당없음',
# '정정취소': '정상주문', '현금신용': '현금', '주문조건': 'IOC', '체결기준잔고수량': 0, '대출일': 0,
# '주문호가구분': '보통', '매도가능수량': 0}

# conclution
# {'체결플래그': '확인', '주문번호': 24337, '주문수량': 1, '주문가격': 474, '원주문': 0,
# '종목코드': 'A008800', '종목명': '행남생활건강', '매수매도': '매수', '신용대출': '해당없음',
# '정정취소': '취소주문', '현금신용': '현금', '주문조건': 'IOC', '체결기준잔고수량': 0, '대출일': 0,
# '주문호가구분': '보통', '매도가능수량': 0}
#
# ※ 주의: 현재 모의투자 시스템에서는 IOC/FOK 주문을 지원하지 않습니다
# ※ 주의: 제공된 코드는 개발자의 코딩을 돕기 위한 단순 예제로만 제공 됩니다.
# 주문의 경우 반드시 사용자가 코드를 먼저 확인 후 주의 해서 테스트 하시기 바랍니다.

ame  # 서비스가 다른 이벤트를 구분하기 위한 이름
self.parent = parent  # callback 을 위해 보관

# 구분값 : 텍스트로 변경하기 위해 딕셔너리 이용
self.dicflag12 = {'1': '매도', '2': '매수'}
self.dicflag14 = {'1': '체결', '2': '확인', '3': '거부', '4': '접수'}
self.dicflag15 = {'00': '해당없음', '01': '유통융자', '02': '자기융자', '03': '유통대주',
                  '04': '자기대주', '05': '주식담보대출'}
self.dicflag16 = {'1': '정상주문', '2': '정정주문', '3': '취소주문'}
self.dicflag17 = {'1': '현금', '2': '신용', '3': '선물대용', '4': '공매도'}
self.dicflag18 = {'01': '보통', '02': '임의', '03': '시장가', '05': '조건부지정가'}
self.dicflag19 = {'0': '없음', '1': 'IOC', '2': 'FOK'}


# PLUS 로 부터 실제로 시세를 수신 받는 이벤트 핸들러
def OnReceived(self):
    print(self.name)
    if self.name == 'conclution':
        # 주문 체결 실시간 업데이트
        conc = {}

        # 체결 플래그
        conc['체결플래그'] = self.dicflag14[self.client.GetHeaderValue(14)]

        conc['주문번호'] = self.client.GetHeaderValue(5)  # 주문번호
        conc['주문수량'] = self.client.GetHeaderValue(3)  # 주문/체결 수량
        conc['주문가격'] = self.client.GetHeaderValue(4)  # 주문/체결 가격
        conc['원주문'] = self.client.GetHeaderValue(6)
        conc['종목코드'] = self.client.GetHeaderValue(9)  # 종목코드
        conc['종목명'] = g_objCodeMgr.CodeToName(conc['종목코드'])

        conc['매수매도'] = self.dicflag12[self.client.GetHeaderValue(12)]

        flag15 = self.client.GetHeaderValue(15)  # 신용대출구분코드
        if (flag15 in self.dicflag15):
            conc['신용대출'] = self.dicflag15[flag15]
        else:
            conc['신용대출'] = '기타'

        conc['정정취소'] = self.dicflag16[self.client.GetHeaderValue(16)]
        conc['현금신용'] = self.dicflag17[self.client.GetHeaderValue(17)]
        conc['주문조건'] = self.dicflag19[self.client.GetHeaderValue(19)]

        conc['체결기준잔고수량'] = self.client.GetHeaderValue(23)
        conc['대출일'] = self.client.GetHeaderValue(20)
        flag18 = self.client.GetHeaderValue(18)
        if (flag18 in self.dicflag18):
            conc['주문호가구분'] = self.dicflag18[flag18]
        else:
            conc['주문호가구분'] = '기타'

        conc['매도가능수량'] = self.client.GetHeaderValue(22)

        print(conc)

        return


# CpPBConclusion: 실시간 주문 체결 수신 클래그
class CpPBConclusion:
    def __init__(self):
        self.name = 'conclution'
        self.obj = win32com.client.Dispatch('DsCbo1.CpConclusion')

    def Subscribe(self, parent):
        self.parent = parent
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, parent)
        self.obj.Subscribe()

    def Unsubscribe(self):
        self.obj.Unsubscribe()


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print('PLUS가 정상적으로 연결되지 않음. ')
            return
        self.objStockMst = win32com.client.Dispatch('DsCbo1.StockMst')
        return

    def Request(self, code, caller):
        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print('통신상태', self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        caller.curData = {}

        caller.curData['code'] = code
        caller.curData['종목명'] = g_objCodeMgr.CodeToName(code)
        caller.curData['현재가'] = self.objStockMst.GetHeaderValue(11)  # 종가
        caller.curData['대비'] = self.objStockMst.GetHeaderValue(12)  # 전일대비
        caller.curData['기준가'] = self.objStockMst.GetHeaderValue(27)  # 기준가
        caller.curData['거래량'] = self.objStockMst.GetHeaderValue(18)  # 거래량
        caller.curData['예상플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        caller.curData['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        caller.curData['예상대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결대비

        # 10차호가
        for i in range(10):
            key1 = '매도호가%d' % (i + 1)
            key2 = '매수호가%d' % (i + 1)
            caller.curData[key1] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            caller.curData[key2] = (self.objStockMst.GetDataValue(1, i))  # 매수호가

        # print(caller.curData)

        return True


class CpRPOrder:
    def __init__(self):
        # 연결 여부 체크
        if (g_objCpStatus.IsConnect == 0):
            print('PLUS가 정상적으로 연결되지 않음. ')
            return

        # 주문 초기화
        if (g_objCpTrade.TradeInit(0) != 0):
            print('주문 초기화 실패')
            return

        self.acc = g_objCpTrade.AccountNumber[0]  # 계좌번호
        self.accFlag = g_objCpTrade.GoodsList(self.acc, 1)  # 주식상품 구분
        print(self.acc, self.accFlag[0])

        # 주문 object 생성
        self.objBuySell = win32com.client.Dispatch('CpTrade.CpTd0311')  # 매수

        # 주문 조건 구분 '0': 없음 1: IOC 2: FOK
        self.dicFlag = {'IOC': '1', 'FOK': '2', '없음': '0'}

    def sellOrder(self, code, price, amount, flag):
        print('신규 매도', '주문조건구분:', flag, '종목코드:', code, '가격:', price, '수량:', amount)

        self.objBuySell.SetInputValue(0, '1')  # 1 매도 2 매수
        self.objBuySell.SetInputValue(1, self.acc)  # 계좌번호
        self.objBuySell.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objBuySell.SetInputValue(3, code)  # 종목코드
        self.objBuySell.SetInputValue(4, amount)  # 수량
        self.objBuySell.SetInputValue(5, price)  # 주문단가
        orderflag = self.dicFlag[flag]  # 주문 조건 구분 '0': 없음 1: IOC 2: FOK
        self.objBuySell.SetInputValue(7, orderflag)  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objBuySell.SetInputValue(8, '01')  # 주문호가 구분코드 - 01: 보통 03 시장가 05 조건부지정가

        # 주문 요청
        self.objBuySell.BlockRequest()

        rqStatus = self.objBuySell.GetDibStatus()
        rqRet = self.objBuySell.GetDibMsg1()
        print('통신상태', rqStatus, rqRet)
        if rqStatus != 0:
            return False

        return True

    def buyOrder(self, code, price, amount, flag):
        print('신규 매수', '주문조건구분:', flag, '종목코드:', code, '가격:', price, '수량:', amount)

        self.objBuySell.SetInputValue(0, '2')  # 1 매도 2 매수
        self.objBuySell.SetInputValue(1, self.acc)  # 계좌번호
        self.objBuySell.SetInputValue(2, self.accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objBuySell.SetInputValue(3, code)  # 종목코드
        self.objBuySell.SetInputValue(4, amount)  # 수량
        self.objBuySell.SetInputValue(5, price)  # 주문단가
        orderflag = self.dicFlag[flag]  # 주문 조건 구분 '0': 없음 1: IOC 2: FOK
        self.objBuySell.SetInputValue(7, orderflag)  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objBuySell.SetInputValue(8, '01')  # 주문호가 구분코드 - 01: 보통 03 시장가 05 조건부지정가

        # 주문 요청
        self.objBuySell.BlockRequest()

        rqStatus = self.objBuySell.GetDibStatus()
        rqRet = self.objBuySell.GetDibMsg1()
        print('통신상태', rqStatus, rqRet)
        if rqStatus != 0:
            return False
        return True


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('IOC/FOK 주문 테스트')
        self.setGeometry(300, 300, 300, 230)

        # 주문 통신 object
        self.objRpOrder = CpRPOrder()
        # 현재가 통신 object
        self.curData = {}
        self.objCur = CpRPCurrentPrice()

        # 주문체결은 미리 실시간 요청
        self.conclution = CpPBConclusion()
        self.conclution.Subscribe(self)

        nH = 20
        # 코드 입력기
        self.codeEdit = QLineEdit('', self)
        self.codeEdit.move(20, nH)
        self.codeEdit.textChanged.connect(self.codeEditChanged)
        self.codeEdit.setText('')
        self.label = QLabel('종목코드', self)
        self.label.move(140, nH)
        self.code = ''
        nH += 50

        btnIOCSell = QPushButton('매도IOC', self)
        btnIOCSell.move(20, nH)
        btnIOCSell.resize(200, 30)
        btnIOCSell.clicked.connect(self.btnIOCSell_clicked)
        nH += 50

        btnIOCBuy = QPushButton('매수IOC', self)
        btnIOCBuy.move(20, nH)
        btnIOCBuy.resize(200, 30)
        btnIOCBuy.clicked.connect(self.btnIOCBuy_clicked)
        nH += 50

        btnFOKSell = QPushButton('매도FOK', self)
        btnFOKSell.move(20, nH)
        btnFOKSell.resize(200, 30)
        btnFOKSell.clicked.connect(self.btnFOKSell_clicked)
        nH += 50

        btnFOKBuy = QPushButton('매수FOK', self)
        btnFOKBuy.move(20, nH)
        btnFOKBuy.resize(200, 30)
        btnFOKBuy.clicked.connect(self.btnFOKBuy_clicked)
        nH += 50

        btnExit = QPushButton('종료', self)
        btnExit.move(20, nH)
        btnExit.resize(200, 30)
        btnExit.clicked.connect(self.btnExit_clicked)
        nH += 50

        self.setGeometry(300, 300, 300, nH)

    def btnIOCSell_clicked(self):
        return self.sellOrder('IOC')

    def btnIOCBuy_clicked(self):
        return self.buyOrder('IOC')

    def btnFOKSell_clicked(self):
        return self.sellOrder('FOK')

    def btnFOKBuy_clicked(self):
        return self.buyOrder('FOK')

    # 종료
    def btnExit_clicked(self):
        exit()

    def codeEditChanged(self):
        code = self.codeEdit.text()
        self.setCode(code)

    def setCode(self, code):
        if len(code) < 6:
            return

        print(code)
        if not (code[0] == 'A'):
            code = 'A' + code

        name = g_objCodeMgr.CodeToName(code)
        if len(name) == 0:
            print('종목코드 확인')
            return

        self.label.setText(name)
        self.code = code

    # 매수는 매수3호가에 IOC/FOK 주문을 낸다.
    def buyOrder(self, flag):
        # 현재가 통신
        if (self.objCur.Request(self.code, self) == False):
            w = QWidget()
            QMessageBox.warning(w, '오류', '현재가 통신 오류 발생/주문 중단')
            return

        self.objRpOrder.buyOrder(self.code, self.curData['매수호가3'], 1, flag)

    # 매도는 매도3호가에 IOC/FOK 주문을 낸다.
    def sellOrder(self, flag):
        # 현재가 통신
        if (self.objCur.Request(self.code, self) == False):
            w = QWidget()
            QMessageBox.warning(w, '오류', '현재가 통신 오류 발생/주문 중단')
            return

        self.objRpOrder.sellOrder(self.code, self.curData['매도호가3'], 1, flag)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
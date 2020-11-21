# 대신증권 API
# 주식 현재가 10차 호가 화면 구현하기

# 그림과 같은 10차 호가 UI 를 파이썬으로 구현하는 예제입니다
# 10차 호가를 위한 UI 디자인은 QT DESIGNER 를 사용했습니다.
# 개발을 위한 코드와 ui 파일은 첨부 파일을 통해 내려 받을 수 있습니다

import sys
from PyQt5 import QtWidgets
from PyQt5 import uic
from PyQt5.QtCore import pyqtSlot
import win32com.client

# cp object
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")


# 현재가 정보 저장 구조체
class stockPricedData:
    def __init__(self):
        self.dicEx = {ord('0'): "동시호가/장중 아님", ord('1'): "동시호가", ord('2'): "장중"}
        self.code = ""
        self.name = ""
        self.cur = 0  # 현재가
        self.diff = 0  # 대비
        self.diffp = 0  # 대비율
        self.offer = [0 for _ in range(10)]  # 매도호가
        self.bid = [0 for _ in range(10)]  # 매수호가
        self.offervol = [0 for _ in range(10)]  # 매도호가 잔량
        self.bidvol = [0 for _ in range(10)]  # 매수호가 잔량
        self.totOffer = 0  # 총매도잔량
        self.totBid = 0  # 총매수 잔량
        self.vol = 0  # 거래량
        self.baseprice = 0  # 기준가

        # 예상체결가 정보
        self.exFlag = ord('2')
        self.expcur = 0  # 예상체결가
        self.expdiff = 0  # 예상 대비
        self.expdiffp = 0  # 예상 대비율
        self.expvol = 0  # 예상 거래량
        self.objCur = CpPBStockCur()
        self.objOfferbid = CpPBStockBid()

    def __del__(self):
        self.objCur.Unsubscribe()
        self.objOfferbid.Unsubscribe()

    # 전일 대비 계산
    def makediffp(self):
        lastday = 0
        if (self.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            if self.baseprice > 0:
                lastday = self.baseprice
            else:
                lastday = self.expcur - self.expdiff
            if lastday:
                self.expdiffp = (self.expdiff / lastday) * 100
            else:
                self.expdiffp = 0
        else:
            if self.baseprice > 0:
                lastday = self.baseprice
            else:
                lastday = self.cur - self.diff
            if lastday:
                self.diffp = (self.diff / lastday) * 100
            else:
                self.diffp = 0

    def getCurColor(self):
        diff = self.diff
        if (self.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            diff = self.expdiff
        if (diff > 0):
            return 'color: red'
        elif (diff == 0):
            return 'color: black'
        elif (diff < 0):
            return 'color: blue'


# CpEvent: 실시간 이벤트 수신 클래스
class CpEvent:
    def set_params(self, client, name, rpMst, parent):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.parent = parent  # callback 을 위해 보관
        self.rpMst = rpMst

    # PLUS 로 부터 실제로 시세를 수신 받는 이벤트 핸들러
    def OnReceived(self):
        if self.name == "stockcur":
            # 현재가 체결 데이터 실시간 업데이트
            self.rpMst.exFlag = self.client.GetHeaderValue(19)  # 예상체결 플래그
            code = self.client.GetHeaderValue(0)
            diff = self.client.GetHeaderValue(2)
            cur = self.client.GetHeaderValue(13)  # 현재가
            vol = self.client.GetHeaderValue(9)  # 거래량

            # 예제는 장중만 처리 함.
            if (self.rpMst.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
                # 예상체결가 정보
                self.rpMst.expcur = cur
                self.rpMst.expdiff = diff
                self.rpMst.expvol = vol
            else:
                self.rpMst.cur = cur
                self.rpMst.diff = diff
                self.rpMst.makediffp()
                self.vol = vol

            self.rpMst.makediffp()
            # 현재가 업데이트
            self.parent.monitorPriceChange()

            return

        elif self.name == "stockbid":
            # 현재가 10차 호가 데이터 실시간 업데이c
            code = self.client.GetHeaderValue(0)
            dataindex = [3, 7, 11, 15, 19, 27, 31, 35, 39, 43]
            obi = 0
            for i in range(10):
                self.rpMst.offer[i] = self.client.GetHeaderValue(dataindex[i])
                self.rpMst.bid[i] = self.client.GetHeaderValue(dataindex[i] + 1)
                self.rpMst.offervol[i] = self.client.GetHeaderValue(dataindex[i] + 2)
                self.rpMst.bidvol[i] = self.client.GetHeaderValue(dataindex[i] + 3)

            self.rpMst.totOffer = self.client.GetHeaderValue(23)
            self.rpMst.totBid = self.client.GetHeaderValue(24)
            # 10차 호가 변경 call back 함수 호출
            self.parent.monitorOfferbidChange()
            return


# SB/PB 요청 ROOT 클래스
class CpPublish:
    def __init__(self, name, serviceID):
        self.name = name
        self.obj = win32com.client.Dispatch(serviceID)
        self.bIsSB = False

    def Subscribe(self, var, rpMst, parent):
        if self.bIsSB:
            self.Unsubscribe()

        if (len(var) > 0):
            self.obj.SetInputValue(0, var)

        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, rpMst, parent)
        self.obj.Subscribe()
        self.bIsSB = True

    def Unsubscribe(self):
        if self.bIsSB:
            self.obj.Unsubscribe()
        self.bIsSB = False


# CpPBStockCur: 실시간 현재가 요청 클래스
class CpPBStockCur(CpPublish):
    def __init__(self):
        super().__init__("stockcur", "DsCbo1.StockCur")


# CpPBStockBid: 실시간 10차 호가 요청 클래스
class CpPBStockBid(CpPublish):
    def __init__(self):
        super().__init__("stockbid", "Dscbo1.StockJpBid")


# SB/PB 요청 ROOT 클래스
class CpPBConnection:
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpUtil.CpCybos")
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, "connection", None)


# CpRPCurrentPrice:  현재가 기본 정보 조회 클래스
class CpRPCurrentPrice:
    def __init__(self):
        if (g_objCpStatus.IsConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        return

    def Request(self, code, rtMst, callbackobj):
        # 현재가 통신
        rtMst.objCur.Unsubscribe()
        rtMst.objOfferbid.Unsubscribe()

        self.objStockMst.SetInputValue(0, code)
        ret = self.objStockMst.BlockRequest()
        if self.objStockMst.GetDibStatus() != 0:
            print("통신상태", self.objStockMst.GetDibStatus(), self.objStockMst.GetDibMsg1())
            return False

        # 수신 받은 현재가 정보를 rtMst 에 저장
        rtMst.code = code
        rtMst.name = g_objCodeMgr.CodeToName(code)
        rtMst.cur = self.objStockMst.GetHeaderValue(11)  # 종가
        rtMst.diff = self.objStockMst.GetHeaderValue(12)  # 전일대비
        rtMst.baseprice = self.objStockMst.GetHeaderValue(27)  # 기준가
        rtMst.vol = self.objStockMst.GetHeaderValue(18)  # 거래량
        rtMst.exFlag = self.objStockMst.GetHeaderValue(58)  # 예상플래그
        rtMst.expcur = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        rtMst.expdiff = self.objStockMst.GetHeaderValue(56)  # 예상체결대비
        rtMst.makediffp()

        rtMst.totOffer = self.objStockMst.GetHeaderValue(71)  # 총매도잔량
        rtMst.totBid = self.objStockMst.GetHeaderValue(73)  # 총매수잔량

        # 10차호가
        for i in range(10):
            rtMst.offer[i] = (self.objStockMst.GetDataValue(0, i))  # 매도호가
            rtMst.bid[i] = (self.objStockMst.GetDataValue(1, i))  # 매수호가
            rtMst.offervol[i] = (self.objStockMst.GetDataValue(2, i))  # 매도호가 잔량
            rtMst.bidvol[i] = (self.objStockMst.GetDataValue(3, i))  # 매수호가 잔량

        rtMst.objCur.Subscribe(code, rtMst, callbackobj)
        rtMst.objOfferbid.Subscribe(code, rtMst, callbackobj)


class Form(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.ui = uic.loadUi("hoga.ui", self)
        self.ui.show()
        self.objMst = CpRPCurrentPrice()
        self.item = stockPricedData()

        self.setCode("000660")

    @pyqtSlot()
    def slot_codeupdate(self):
        code = self.ui.editCode.toPlainText()
        self.setCode(code)

    def slot_codechanged(self):
        code = self.ui.editCode.toPlainText()
        self.setCode(code)

    def monitorPriceChange(self):
        self.displyHoga()

    def monitorOfferbidChange(self):
        self.displyHoga()

    def setCode(self, code):
        if len(code) < 6:
            return

        print(code)
        if not (code[0] == "A"):
            code = "A" + code

        name = g_objCodeMgr.CodeToName(code)
        if len(name) == 0:
            print("종목코드 확인")
            return

        self.ui.label_name.setText(name)

        if (self.objMst.Request(code, self.item, self) == False):
            return
        self.displyHoga()

    def displyHoga(self):
        self.ui.label_offer10.setText(format(self.item.offer[9], ','))
        self.ui.label_offer9.setText(format(self.item.offer[8], ','))
        self.ui.label_offer8.setText(format(self.item.offer[7], ','))
        self.ui.label_offer7.setText(format(self.item.offer[6], ','))
        self.ui.label_offer6.setText(format(self.item.offer[5], ','))
        self.ui.label_offer5.setText(format(self.item.offer[4], ','))
        self.ui.label_offer4.setText(format(self.item.offer[3], ','))
        self.ui.label_offer3.setText(format(self.item.offer[2], ','))
        self.ui.label_offer2.setText(format(self.item.offer[1], ','))
        self.ui.label_offer1.setText(format(self.item.offer[0], ','))

        self.ui.label_offer_v10.setText(format(self.item.offervol[9], ','))
        self.ui.label_offer_v9.setText(format(self.item.offervol[8], ','))
        self.ui.label_offer_v8.setText(format(self.item.offervol[7], ','))
        self.ui.label_offer_v7.setText(format(self.item.offervol[6], ','))
        self.ui.label_offer_v6.setText(format(self.item.offervol[5], ','))
        self.ui.label_offer_v5.setText(format(self.item.offervol[4], ','))
        self.ui.label_offer_v4.setText(format(self.item.offervol[3], ','))
        self.ui.label_offer_v3.setText(format(self.item.offervol[2], ','))
        self.ui.label_offer_v2.setText(format(self.item.offervol[1], ','))
        self.ui.label_offer_v1.setText(format(self.item.offervol[0], ','))

        self.ui.label_bid10.setText(format(self.item.bid[9], ','))
        self.ui.label_bid9.setText(format(self.item.bid[8], ','))
        self.ui.label_bid8.setText(format(self.item.bid[7], ','))
        self.ui.label_bid7.setText(format(self.item.bid[6], ','))
        self.ui.label_bid6.setText(format(self.item.bid[5], ','))
        self.ui.label_bid5.setText(format(self.item.bid[4], ','))
        self.ui.label_bid4.setText(format(self.item.bid[3], ','))
        self.ui.label_bid3.setText(format(self.item.bid[2], ','))
        self.ui.label_bid2.setText(format(self.item.bid[1], ','))
        self.ui.label_bid1.setText(format(self.item.bid[0], ','))

        self.ui.label_bid_v10.setText(format(self.item.bidvol[9], ','))
        self.ui.label_bid_v9.setText(format(self.item.bidvol[8], ','))
        self.ui.label_bid_v8.setText(format(self.item.bidvol[7], ','))
        self.ui.label_bid_v7.setText(format(self.item.bidvol[6], ','))
        self.ui.label_bid_v6.setText(format(self.item.bidvol[5], ','))
        self.ui.label_bid_v5.setText(format(self.item.bidvol[4], ','))
        self.ui.label_bid_v4.setText(format(self.item.bidvol[3], ','))
        self.ui.label_bid_v3.setText(format(self.item.bidvol[2], ','))
        self.ui.label_bid_v2.setText(format(self.item.bidvol[1], ','))
        self.ui.label_bid_v1.setText(format(self.item.bidvol[0], ','))

        cur = self.item.cur
        diff = self.item.diff
        diffp = self.item.diffp
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            cur = self.item.expcur
            diff = self.item.expdiff
            diffp = self.item.expdiffp

        strcur = format(cur, ',')
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            strcur = "*" + strcur

        curcolor = self.item.getCurColor()
        self.ui.label_cur.setStyleSheet(curcolor)
        self.ui.label_cur.setText(strcur)
        strdiff = str(diff) + "  " + format(diffp, '.2f')
        strdiff += "%"
        self.ui.label_diff.setText(strdiff)
        self.ui.label_diff.setStyleSheet(curcolor)

        self.ui.label_totoffer.setText(format(self.item.totOffer, ','))
        self.ui.label_totbid.setText(format(self.item.totBid, ','))


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = Form()
    sys.exit(app.exec())
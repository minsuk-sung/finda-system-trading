# 대신증권 API
# 주식 현재가(10차호가/시간대별/일자별) 구현하기 예제

# 주식 현재가 화면을 구성하는 10차 호가, 시간대별, 일자별 데이터를 구현한 파이썬 예제입니다
# 화면 UI 는 PYQT 를 이용하였고 첨부된 파일에서 소스와 UI 를 받아 확인 가능합니다.

import sys

import pandas
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import *
import win32com.client
from pandas import Series, DataFrame
import locale

# cp object
g_objCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
g_objCpStatus = win32com.client.Dispatch("CpUtil.CpCybos")
g_objCpTrade = win32com.client.Dispatch("CpTrade.CpTdUtil")
locale.setlocale(locale.LC_ALL, '')


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
        self.tvol = 0  # 순간 체결량
        self.baseprice = 0  # 기준가
        self.high = 0
        self.low = 0
        self.open = 0
        self.volFlag = ord('0')  # 체결매도/체결 매수 여부
        self.time = 0
        self.sum_buyvol = 0
        self.sum_sellvol = 0
        self.vol_str = 0

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
                self.rpMst.vol = vol
                self.rpMst.open = self.client.GetHeaderValue(4)
                self.rpMst.high = self.client.GetHeaderValue(5)
                self.rpMst.low = self.client.GetHeaderValue(6)
                self.rpMst.tvol = self.client.GetHeaderValue(17)
                self.rpMst.volFlag = self.client.GetHeaderValue(14)  # '1'  매수 '2' 매도
                self.rpMst.time = self.client.GetHeaderValue(18)
                self.rpMst.sum_buyvol = self.client.GetHeaderValue(16)  # 누적매수체결수량 (체결가방식)
                self.rpMst.sum_sellvol = self.client.GetHeaderValue(15)  # 누적매도체결수량 (체결가방식)
                if (self.rpMst.sum_sellvol):
                    self.rpMst.volstr = self.rpMst.sum_buyvol / self.rpMst.sum_sellvol * 100
                else:
                    self.rpMst.volstr = 0

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


# CpWeekList:  일자별 리스트 구하기
class CpWeekList:
    def __init__(self):
        self.objWeek = win32com.client.Dispatch("Dscbo1.StockWeek")
        return

    def Request(self, code, caller):
        # 현재가 통신
        self.objWeek.SetInputValue(0, code)
        # 데이터들
        dates = []
        opens = []
        highs = []
        lows = []
        closes = []
        diffs = []
        vols = []
        diffps = []
        foreign_vols = []
        foreign_diff = []
        foreign_p = []

        # 누적 개수 - 100 개까지만 하자
        sumCnt = 0
        while True:
            ret = self.objWeek.BlockRequest()
            if self.objWeek.GetDibStatus() != 0:
                print("통신상태", self.objWeek.GetDibStatus(), self.objWeek.GetDibMsg1())
                return False

            cnt = self.objWeek.GetHeaderValue(1)
            sumCnt += cnt
            if cnt == 0:
                break

            for i in range(cnt):
                dates.append(self.objWeek.GetDataValue(0, i))
                opens.append(self.objWeek.GetDataValue(1, i))
                highs.append(self.objWeek.GetDataValue(2, i))
                lows.append(self.objWeek.GetDataValue(3, i))
                closes.append(self.objWeek.GetDataValue(4, i))

                temp = self.objWeek.GetDataValue(5, i)
                diffs.append(temp)
                vols.append(self.objWeek.GetDataValue(6, i))

                temp2 = self.objWeek.GetDataValue(10, i)
                if (temp < 0):
                    temp2 *= -1
                diffps.append(temp2)

                foreign_vols.append(self.objWeek.GetDataValue(7, i))  # 외인보유
                foreign_diff.append(self.objWeek.GetDataValue(8, i))  # 외인보유 전일대비
                foreign_p.append(self.objWeek.GetDataValue(9, i))  # 외인비중

            if (sumCnt > 100):
                break

            if self.objWeek.Continue == False:
                break

        if len(dates) == 0:
            return False

        caller.rpWeek = None
        weekCol = {'close': closes,
                   'diff': diffs,
                   'diffp': diffps,
                   'vol': vols,
                   'open': opens,
                   'high': highs,
                   'low': lows,
                   'for_v': foreign_vols,
                   'for_d': foreign_diff,
                   'for_p': foreign_p,
                   }
        caller.rpWeek = DataFrame(weekCol, index=dates)
        return True


# CpStockBid:  시간대별 조회
class CpStockBid:
    def __init__(self):
        self.objSBid = win32com.client.Dispatch("Dscbo1.StockBid")
        return

    def Request(self, code, caller):
        # 현재가 통신
        self.objSBid.SetInputValue(0, code)
        self.objSBid.SetInputValue(2, 80)  # 요청개수 (최대 80)
        self.objSBid.SetInputValue(3, ord('C'))  # C 체결가 비교 방식 H 호가 비교방식

        times = []
        curs = []
        diffs = []
        tvols = []
        offers = []
        bids = []
        vols = []
        offerbidFlags = []  # 체결 상태 '1' 매수 '2' 매도
        volstrs = []  # 체결강도
        marketFlags = []  # 장구분 '1' 동시호가 예상체결' '2' 장중

        # 누적 개수 - 100 개까지만 하자
        sumCnt = 0
        while True:
            ret = self.objSBid.BlockRequest()
            if self.objSBid.GetDibStatus() != 0:
                print("통신상태", self.objSBid.GetDibStatus(), self.objSBid.GetDibMsg1())
                return False

            cnt = self.objSBid.GetHeaderValue(2)
            sumCnt += cnt
            if cnt == 0:
                break

            strcur = ""
            strflag = ""
            strflag2 = ""
            for i in range(cnt):
                cur = self.objSBid.GetDataValue(4, i)
                times.append(self.objSBid.GetDataValue(9, i))
                diffs.append(self.objSBid.GetDataValue(1, i))
                vols.append(self.objSBid.GetDataValue(5, i))
                tvols.append(self.objSBid.GetDataValue(6, i))
                offers.append(self.objSBid.GetDataValue(2, i))
                bids.append(self.objSBid.GetDataValue(3, i))
                flag = self.objSBid.GetDataValue(7, i)
                if (flag == ord('1')):
                    strflag = "체결매수"
                else:
                    strflag = "체결매도"
                offerbidFlags.append(strflag)
                volstrs.append(self.objSBid.GetDataValue(8, i))
                flag = self.objSBid.GetDataValue(10, i)
                if (flag == ord('1')):
                    strflag2 = "예상체결"
                    # strcur = '*' + str(cur)
                else:
                    strflag2 = "장중"
                    # strcur = str(cur)
                marketFlags.append(strflag2)
                curs.append(cur)

            if (sumCnt > 100):
                break

            if self.objSBid.Continue == False:
                break

        if len(times) == 0:
            return False

        caller.rpStockBid = None
        sBidCol = {'time': times,
                   'cur': curs,
                   'diff': diffs,
                   'vol': vols,
                   'tvol': tvols,
                   'offer': offers,
                   'bid': bids,
                   'flag': offerbidFlags,
                   'market': marketFlags,
                   'volstr': volstrs}
        caller.rpStockBid = DataFrame(sBidCol)
        print(caller.rpStockBid)
        return True


class Form(QtWidgets.QDialog):
    def __init__(self, parent=None):
        QtWidgets.QDialog.__init__(self, parent)
        self.ui = uic.loadUi("hoga_1.ui", self)
        self.ui.show()
        self.objMst = CpRPCurrentPrice()
        self.item = stockPricedData()

        # 일자별
        self.objWeek = CpWeekList()
        self.rpWeek = DataFrame()  # 일자별 데이터프레임

        # 시간대별
        self.rpStockBid = DataFrame()
        self.objStockBid = CpStockBid()
        self.todayIndex = 0

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
        self.updateWeek()
        self.updateStockBid()

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

        # 일자별
        self.ui.tableWeek.clearContents()
        if (self.objWeek.Request(code, self) == True):
            print(self.rpWeek)
            self.displyWeek()

        # 시간대별
        self.ui.tableStockBid.clearContents()
        if (self.objStockBid.Request(code, self) == True):
            self.displyStockBid()

    # 10차 호가 UI 채우기
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

    # 일자별 리스트 UI 채우기
    def displyWeek(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return
        self.ui.tableWeek.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpWeek.iterrows():
            datas = [index, row['close'], row['diff'], row['diffp'], row['vol'], row['open'], row['high'], row['low'],
                     row['for_v'], row['for_d'], row['for_p']]
            for col in range(len(datas)):
                val = ''
                if (col == 0):  # 일자
                    # 20170929 ==> 2017/09/29
                    yyyy = int(datas[col] / 10000)
                    mm = int(datas[col] - (yyyy * 10000))
                    dd = mm % 100
                    mm = mm / 100
                    val = '%04d/%02d/%02d' % (yyyy, mm, dd)
                elif (col == 3 or col == 10):  # 대비율
                    val = locale.format('%.2f', datas[col], 1)
                    val += "%"

                else:
                    val = locale.format('%d', datas[col], 1)

                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableWeek.setItem(nRow, col, item)

            if (nRow == 0):
                self.todayIndex = index
            nRow += 1

            self.tableWeek.resizeColumnsToContents()
        return

    # 일자별 리스트 UI 채우기 - 오늘 날짜 업데이트
    def updateWeek(self):
        rowcnt = len(self.rpWeek.index)
        if rowcnt == 0:
            return

        # 오늘 날짜 데이터 업데이트
        self.rpWeek.set_value(self.todayIndex, 'close', self.item.cur)
        self.rpWeek.set_value(self.todayIndex, 'open', self.item.open)
        self.rpWeek.set_value(self.todayIndex, 'high', self.item.high)
        self.rpWeek.set_value(self.todayIndex, 'low', self.item.low)
        self.rpWeek.set_value(self.todayIndex, 'vol', self.item.vol)
        self.rpWeek.set_value(self.todayIndex, 'diff', self.item.diff)
        self.rpWeek.set_value(self.todayIndex, 'diffp', self.item.diffp)

        datas = [self.todayIndex, self.item.cur, self.item.diff, self.item.diffp, self.item.vol,
                 self.item.open, self.item.high, self.item.low]
        for col in range(len(datas)):
            val = ''
            if (col == 0):  # 일자
                # 20170929 ==> 2017/09/29
                yyyy = int(datas[col] / 10000)
                mm = int(datas[col] - (yyyy * 10000))
                dd = mm % 100
                mm = mm / 100
                val = '%04d/%02d/%02d' % (yyyy, mm, dd)
            elif (col == 3):  # 대비율
                val = locale.format('%.2f', datas[col], 1)
                val += "%"

            else:
                val = locale.format('%d', datas[col], 1)

            item = QTableWidgetItem(val)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.ui.tableWeek.setItem(0, col, item)

        return

    # 시간대별 리스트 UI  채우기
    def displyStockBid(self):
        rowcnt = len(self.rpStockBid.index)
        if rowcnt == 0:
            return
        self.ui.tableStockBid.setRowCount(rowcnt)

        nRow = 0

        for index, row in self.rpStockBid.iterrows():
            # 행 내에 표시할 데이터 - 컬럼 순
            datas = [row['time'], row['cur'], row['diff'], row['offer'], row['bid'], row['vol'], row['tvol'],
                     row['tvol'], row['volstr']]
            market = row['market']
            for col in range(len(datas)):
                val = ''
                if col == 0:  # 시각
                    # 155925 ==> 15:59:25
                    hh = int(datas[col] / 10000)
                    mm = int(datas[col] - (hh * 10000))
                    ss = mm % 100
                    mm = mm / 100
                    val = '%02d:%02d:%02d' % (hh, mm, ss)
                elif col == 6:  # 체결매도
                    market = row['flag']
                    if (market == "체결매도"):
                        val = locale.format('%d', datas[col], 1)
                elif col == 7:  # 체결매수
                    market = row['flag']
                    if (market == "체결매수"):
                        val = locale.format('%d', datas[col], 1)
                elif col == 8:  # 체결강도
                    val = locale.format('%.2f', datas[col], 1)
                elif col == 1:  # 현재가
                    val = locale.format('%d', datas[col], 1)
                    if (market == "예상체결"):
                        val = '*' + val
                else:  # 기타
                    val = locale.format('%d', datas[col], 1)
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.ui.tableStockBid.setItem(nRow, col, item)
            nRow += 1

        self.tableStockBid.resizeColumnsToContents()
        return

    def updateStockBid(self):
        rowcnt = len(self.rpStockBid.index)
        if rowcnt == 0:
            return
        if (self.item.exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            return

        buyvol = sellvol = 0
        if self.item.volFlag == ord('1'):
            buyvol = self.item.tvol
        if self.item.volFlag == ord('2'):
            sellvol = self.item.tvol
        line = DataFrame({"time": self.item.time,
                          "cur": self.item.cur,
                          "diff": self.item.diff,
                          "offer": self.item.offer[0],
                          "bid": self.item.bid[0],
                          "vol": self.item.vol,
                          "tvol": buyvol,
                          "tvol": sellvol,
                          "volstr": self.item.volstr},
                         index=[0])

        self.rpStockBid = pandas.concat([line, self.rpStockBid.ix[:]]).reset_index(drop=True)

        # 행 내에 표시할 데이터 - 컬럼 순
        datas = [self.item.time, self.item.cur, self.item.diff, self.item.offer[0], self.item.bid[0],
                 self.item.vol, sellvol, buyvol, self.item.volstr]
        self.ui.tableStockBid.insertRow(0)
        for col in range(len(datas)):
            val = ''
            if col == 0:  # 시각
                # 155925 ==> 15:59:25
                hh = int(datas[col] / 10000)
                mm = int(datas[col] - (hh * 10000))
                ss = mm % 100
                mm = mm / 100
                val = '%02d:%02d:%02d' % (hh, mm, ss)
            elif col == 6:  # 체결매도
                val = locale.format('%d', datas[col], 1)
            elif col == 7:  # 체결매수
                val = locale.format('%d', datas[col], 1)
            elif col == 8:  # 체결강도
                val = locale.format('%.2f', datas[col], 1)
            else:  # 기타
                val = locale.format('%d', datas[col], 1)

            item = QTableWidgetItem(val)
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.ui.tableStockBid.setItem(0, col, item)

        return


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = Form()
    sys.exit(app.exec())
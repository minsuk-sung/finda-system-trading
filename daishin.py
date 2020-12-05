# Made by Minsuk Sung
# Contact: mssung94@gmail.com
# Homepage: minsuksung-ai.tistory.com
import os
import ctypes
import win32com.client
# from pywinauto import application
from datetime import datetime
from slack import Slack
import time

class Daishin:
    def __init__(self,msg_on):
        print('대신증권 API 초기화')
        
        self.objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
        self.objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')  # 코스닥, 코스피 관련
        self.objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        self.objRq = win32com.client.Dispatch("CpTrade.CpTd6033")  # 주식 잔고 조회 관련
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")  # 주식 현재가 조회 관련
        self.objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")  # 주식 매수/매도 주문 관련
        self.objStockWeek = win32com.client.Dispatch("DsCbo1.StockWeek")  # 일별 데이터
        self.objStockCur = win32com.client.Dispatch("DsCbo1.StockCur")
        self.objCpCash = win32com.client.Dispatch("CpTrade.CpTdNew5331A")

        self.slack = Slack()
        self.today = datetime.now()
        self.channel_list = {'TEST': '#test'}
        self.orderType = {'1':'매도', '2':'매수'}
        self.msg_on = msg_on

        # 프로세스가 관리자 권한으로 실행 여부
        if ctypes.windll.shell32.IsUserAnAdmin():
            print('정상: 관리자권한으로 실행된 프로세스입니다.')
        else:
            print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
            exit(0)

        # 연결 여부 체크
        if self.objCpStatus.IsConnect == 0:
            print("CREON PLUS가 정상적으로 연결되지 않음. ")
            exit(0)

        # 주문 관련 초기화
        if self.objCpTrade.TradeInit(0) != 0:
            print("주문 초기화 실패")
            exit(0)

        self.acc_no = self.objCpTrade.AccountNumber[0]  # 나의 계좌번호
        self.kospi = self.objCodeMgr.GetStockListByMarket(1)  # 코스피
        self.kosdaq = self.objCodeMgr.GetStockListByMarket(2)  # 코스닥

        self.slack.notification(
            pretext=f"",
            title=f"대신증권 트레이딩 시스템 동작",
            fallback=f"대신증권 트레이딩 시스템 동작",  # 미리보기로 볼 수 있는
            text=f"[INFO] 현재 시각 {self.today.year}년 {self.today.month}월 {self.today.day}일 {self.today.hour}시 {self.today.minute}분 트레이딩 시스템이 동작 되었습니다.",
            channel=self.channel_list['TEST'],
            msg_on=self.msg_on
        )

    def get_connect_state(self):
        return self.objCpStatus.IsConnect

    def get_account_info(self):
        accFlag = self.objCpTrade.GoodsList(self.acc_no, 1)  # 주식상품 구분
        self.objRq.SetInputValue(0, self.acc_no)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)
        self.objRq.SetInputValue(3, 2)  # 수익률구분코드 - ( "1" : 100% 기준, "2": 0% 기준)

        self.objRq.BlockRequest()  # 이거 안해주면 정보 못 가져옴

        self.objCpCash.SetInputValue(0, self.acc_no)
        self.objCpCash.SetInputValue(1, accFlag[0])
        self.objCpCash.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        res = {}

        res['계좌명'] = self.objRq.GetHeaderValue(0)
        res['결제잔고수량'] = self.objRq.GetHeaderValue(1)
        res['체결잔고수량'] = self.objRq.GetHeaderValue(2)
        res['총 평가금액'] = self.objRq.GetHeaderValue(3)
        res['평가손익'] = self.objRq.GetHeaderValue(4)
        res['대출금액'] = self.objRq.GetHeaderValue(6)
        # res['수신개수'] = self.objRq.GetHeaderValue(7)
        res['수익률'] = self.objRq.GetHeaderValue(8)
        res['D+2 예상예수금'] = self.objRq.GetHeaderValue(9)
        res['총평가 내 대주평가금액'] = self.objRq.GetHeaderValue(10)
        res['총평가 내 잔고평가금액'] = self.objRq.GetHeaderValue(11)
        res['대주금액'] = self.objRq.GetHeaderValue(12)

        res['주문 가능 금액'] = self.objCpCash.GetHeaderValue(9)

        txt = f"""
01. 현재 결제잔고수량: {res['결제잔고수량']:,}
02. 체결잔고수량: {res['체결잔고수량']:,}
03. 총 평가금액: {res['총 평가금액']:,}
04. 평가손익: {res['평가손익']:,}
05. 대출금액: {res['대출금액']:,}
06. 수익률: {res['수익률']:,}
07. D+2 예상예수금: {res['D+2 예상예수금']:,}
08. 총평가 내 대주평가금액: {res['총평가 내 대주평가금액']:,}
09. 총평가 내 잔고평가금액: {res['총평가 내 잔고평가금액']:,}
10. 대주금액: {res['대주금액']:,}
"""

        self.slack.notification(
            pretext=f"",
            title=f"현재 계좌({self.acc_no}) 잔고 평가현황",
            fallback=f"현재 계좌({self.acc_no}) 잔고 평가현황",  # 미리보기로 볼 수 있는
            text=txt,
            channel=self.channel_list['TEST'],
            msg_on=self.msg_on
        )

        return res

    # 2번째 예제 활용
    def get_current_data(self, code):
        # 현재가 객체 구하기
        self.objStockMst.SetInputValue(0, code)  # 종목 코드 - 삼성전자
        self.objStockMst.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objStockMst.GetDibStatus()
        rqRet = self.objStockMst.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()
            
        # 반환할 값을 위한 딕셔너리
        res = {}

        # 현재가 정보 조회
        res['종목코드'] = self.objStockMst.GetHeaderValue(0)  # 종목코드
        res['종목명'] = self.objStockMst.GetHeaderValue(1)  # 종목명
        res['시간'] = self.objStockMst.GetHeaderValue(4)  # 시간
        res['종가'] = self.objStockMst.GetHeaderValue(11)  # 종가
        res['대비'] = self.objStockMst.GetHeaderValue(12)  # 대비
        res['시가'] = self.objStockMst.GetHeaderValue(13)  # 시가
        res['고가'] = self.objStockMst.GetHeaderValue(14)  # 고가
        res['저가'] = self.objStockMst.GetHeaderValue(15)  # 저가
        res['매도호가'] = self.objStockMst.GetHeaderValue(16)  # 매도호가
        res['매수호가'] = self.objStockMst.GetHeaderValue(17)  # 매수호가
        res['거래량'] = self.objStockMst.GetHeaderValue(18)  # 거래량
        res['거래대금'] = self.objStockMst.GetHeaderValue(19)  # 거래대금

        # 예상 체결관련 정보
        res['예상체결가 구분 플래그'] = self.objStockMst.GetHeaderValue(58)  # 예상체결가 구분 플래그
        res['예상체결가'] = self.objStockMst.GetHeaderValue(55)  # 예상체결가
        res['예상체결가 전일대비'] = self.objStockMst.GetHeaderValue(56)  # 예상체결가 전일대비
        res['예상체결수량'] = self.objStockMst.GetHeaderValue(57)  # 예상체결수량

        # print("코드", res['종목코드'])
        # print("이름", res['종목명'])
        # print("시간", res['시간'])
        # print("종가", res['종가'])
        # print("대비", res['대비'])
        # print("시가", res['시가'])
        # print("고가", res['고가'])
        # print("저가", res['저가'])
        # print("매도호가", res['매도호가'])
        # print("매수호가", res['매수호가'])
        # print("거래량", res['거래량'])
        # print("거래대금", res['거래대금'])
        #
        # if (res['예상체결가 구분 플래그'] == ord('0')):
        #     print("장 구분값: 동시호가와 장중 이외의 시간")
        # elif (res['예상체결가 구분 플래그'] == ord('1')):
        #     print("장 구분값: 동시호가 시간")
        # elif (res['예상체결가 구분 플래그'] == ord('2')):
        #     print("장 구분값: 장중 또는 장종료")

        # print("예상체결가 대비 수량")
        # print("예상체결가", res['예상체결가'])
        # print("예상체결가 대비", res['예상체결가 전일대비'])
        # print("예상체결수량", res['예상체결수량'])

        return res

    def get_daily_data(self, code, cnt):
        self.objStockWeek.SetInputValue(0, code)  # 종목 코드 - 삼성전자

        res = {}

        # 최초 데이터 요청
        # 데이터 요청
        self.objStockWeek.BlockRequest()

        # 통신 결과 확인
        ret = None
        rqStatus = self.objStockWeek.GetDibStatus()
        rqRet = self.objStockWeek.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            ret = False

        # 일자별 정보 데이터 처리
        count = self.objStockWeek.GetHeaderValue(1)  # 데이터 개수
        for i in range(count):
            date = self.objStockWeek.GetDataValue(0, i)  # 일자
            open = self.objStockWeek.GetDataValue(1, i)  # 시가
            high = self.objStockWeek.GetDataValue(2, i)  # 고가
            low = self.objStockWeek.GetDataValue(3, i)  # 저가
            close = self.objStockWeek.GetDataValue(4, i)  # 종가
            diff = self.objStockWeek.GetDataValue(5, i)  # 종가
            vol = self.objStockWeek.GetDataValue(6, i)  # 종가
            res[date] = [open, high, low, close, diff, vol]

        if ret == False:
            exit()

        ret = True

        # 연속 데이터 요청
        # 예제는 5번만 연속 통신 하도록 함.
        NextCount = 1
        while self.objStockWeek.Continue:  # 연속 조회처리
            NextCount += 1;
            if (NextCount > cnt):
                break
            # ret = RequestData(self.objStockWeek)
            rqStatus = self.objStockWeek.GetDibStatus()
            rqRet = self.objStockWeek.GetDibMsg1()
            # print("통신상태", rqStatus, rqRet)
            if rqStatus != 0:
                ret = False

            # 일자별 정보 데이터 처리
            count = self.objStockWeek.GetHeaderValue(1)  # 데이터 개수
            for i in range(count):
                date = self.objStockWeek.GetDataValue(0, i)  # 일자
                open = self.objStockWeek.GetDataValue(1, i)  # 시가
                high = self.objStockWeek.GetDataValue(2, i)  # 고가
                low = self.objStockWeek.GetDataValue(3, i)  # 저가
                close = self.objStockWeek.GetDataValue(4, i)  # 종가
                diff = self.objStockWeek.GetDataValue(5, i)  # 종가
                vol = self.objStockWeek.GetDataValue(6, i)  # 종가
                res[date] = [open, high, low, close, diff, vol]

            if ret == False:
                exit()

        return  res

    def get_my_stocks(self):
        acc = self.acc_no
        accFlag = self.objCpTrade.GoodsList(acc, 1)  # 주식상품 구분

        self.objRq.SetInputValue(0, acc)  # 계좌번호
        self.objRq.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objRq.SetInputValue(2, 50)  # 요청 건수(최대 50)

        self.objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(7)
        # print(cnt)

        res = {}
        txt = ''
        txt = txt + f"-" * 100 + f"\n"
        txt = txt + f"종목코드 \t 종목명 \t 체결잔고수량 \t 체결장부단가 \t 평가금액 \t 평가손익 \n"

        # print("종목코드 종목명 체결잔고수량 체결장부단가 평가금액 평가손익")
        for i in range(cnt):
            item = {}
            code = self.objRq.GetDataValue(12, i)  # 종목코드
            item['종목명'] = self.objRq.GetDataValue(0, i)  # 종목명
            # retcode.append(code)
            # if len(retcode) >= 200:  # 최대 200 종목만,
            #    break
            item['신용구분'] = self.objRq.GetDataValue(1, i)  # 신용구분
            item['대출일'] = self.objRq.GetDataValue(2, i)  # 대출일
            item['체결잔고수량'] = self.objRq.GetDataValue(7, i)  # 체결잔고수량
            item['체결장부단가'] = self.objRq.GetDataValue(17, i)  # 체결장부단가
            item['평가금액'] = self.objRq.GetDataValue(9, i)  # 평가금액(천원미만은 절사 됨)
            item['수익률'] = self.objRq.GetDataValue(11, i)  # 평가손익
            item['평가손익'] = self.objRq.GetDataValue(10, i)  # 평가손익

            # print(code, item['종목명'], item['대출일'], item['체결잔고수량'], item['체결장부단가'], item['체결장부단가'], item['평가금액'], item['평가손익'])
            txt = txt + f"-" * 100 + f"\n"
            txt = txt + f"{code} \t {item['종목명']} \t {item['체결잔고수량']} \t {item['체결장부단가']} \t {item['평가금액']} \t {item['평가손익']:.4f} \n"
            res[code] = item

        self.slack.notification(
            pretext="",
            title=f"현재 계좌({self.acc_no}) 보유 종목 현황",
            fallback=f"현재 계좌({self.acc_no}) 보유 종목 현황",  # 미리보기로 볼 수 있는
            text=txt,
            channel=self.channel_list['TEST'],
            msg_on=self.msg_on
        )

        return res

    def sendOrder(self, bs_type, code, volume):
        # 주식 매수 주문
        acc = self.objCpTrade.AccountNumber[0]  # 계좌번호
        accFlag = self.objCpTrade.GoodsList(acc, 1)  # 주식상품 구분
        # print(acc, accFlag[0])
        self.objStockOrder.SetInputValue(0, bs_type)  # 1: 매도 / 2: 매수
        self.objStockOrder.SetInputValue(1, acc)  # 계좌번호
        self.objStockOrder.SetInputValue(2, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
        self.objStockOrder.SetInputValue(3, code)  # 종목코드 - A003540 - 대신증권 종목
        self.objStockOrder.SetInputValue(4, volume)  # 매수수량 10주
        # self.objStockOrder.SetInputValue(5, price)  # 주문단가  - 14,100원 시장가 매수/매도의 경우 가격 의미 없음
        self.objStockOrder.SetInputValue(7, "0")  # 주문 조건 구분 코드, 0: 기본 1: IOC 2:FOK
        self.objStockOrder.SetInputValue(8, "03")  # 주문호가 구분코드 - 01: 보통 / 02: 임의 / 03: 시장가

        # 매수 주문 요청
        self.objStockOrder.BlockRequest()

        rqStatus = self.objStockOrder.GetDibStatus()
        rqRet = self.objStockOrder.GetDibMsg1()
        # print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()

        self.slack.notification(
            pretext="",
            title=f"{code} {volume}주 {self.orderType[bs_type]} 완료",
            fallback=f"{code} {volume}주 {self.orderType[bs_type]} 완료",  # 미리보기로 볼 수 있는
            text=f"{rqRet}",
            channel=self.channel_list['TEST'],
            msg_on=self.msg_on
        )


if __name__ == '__main__':

    import pandas as pd

    daishin = Daishin(msg_on=False)
    # print(daishin.acc_no)
    # print(f"현재 연결된 계좌번호: {daishin.acc_no}")
    # print(f"현재 코스피 종목수: {len(daishin.kospi)}")
    # print(f"현재 코스닥 종목수: {len(daishin.kosdaq)}")
    # print(f"삼성전자의 현재가 정보: {daishin.get_current_data('A005930')}")
    # print(pd.DataFrame(daishin.get_daily_data('A005930',5)).T)
    # daishin.get_account_info()
    # daishin.get_my_stocks()
    # daishin.sendOrder("2", "035420", 5)
    print(daishin.kospi)

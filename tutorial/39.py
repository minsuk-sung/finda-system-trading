# 대신증권 API
# 투자자별 매매 종합 예제
# 투자자별 매매종합서비스(CpSysDib.CpSvrNew7221) 를 이용하여
# 거래소, 코스닥, KOSPI200 선물/옵션,주식선물, 업종,프로그램매매,개별주식선물(종목) 등의 투자자별 매매동향 데이터를 조회 하는 파이썬 예제입니다.

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
    #     pr"주문 초기화 실패")
    #     return False

    return True

class Cp7221:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.CpSvrNew7221")
        self.InvestIndex = {
            0: '거래소주식',
            1: '코스닥주식',
            2: '선물',
            3: '옵션콜',
            4: '옵션풋',
            5: '주식콜',
            6: '주식풋',
            7: '스타지수선물',
            8: '주식선물',
            9: '채권선물 3년국채(오픈예정)',
            10: '채권선물 5년국채(오픈예정)',
            11: '채권선물 10년국체(오픈예정)',
            12: '금리선물 CD(오픈예정)',
            13: '금리선물통안증권(오픈예정)',
            14: '통화선물미국달러(오픈예정)',
            15: '통화선물엔(오픈예정)',
            16: '통화선물유로(오픈예정)',
            17: '금속상품선물금(오픈예정)',
            18: '농산물파생선물돈육(오픈예정)',
            19: '통화콜옵션미국달러(오픈예정)',
            20: '통화풋옵션미국달러(오픈예정)',
            21: 'CME선물',
            22: '미니금선물'
        }

    def Request(self):
        self.objRq.SetInputValue(0, ord('1'))  # 옵션금액 선물계약
        self.objRq.BlockRequest()

        # 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False

        cnt =  self.objRq.GetHeaderValue(1)
        for key, value in self.InvestIndex.items():
            dicInvest = {}
            dicInvest['개인매도'] = self.objRq.GetDataValue(0, key)
            dicInvest['개인매수'] = self.objRq.GetDataValue(1, key)
            dicInvest['개인순매수'] = self.objRq.GetDataValue(2, key)
            dicInvest['외국인매도'] = self.objRq.GetDataValue(3, key)
            dicInvest['외국인매수'] = self.objRq.GetDataValue(4, key)
            dicInvest['외국인순매수'] = self.objRq.GetDataValue(5, key)
            dicInvest['기관매도'] = self.objRq.GetDataValue(6, key)
            dicInvest['기관매수'] = self.objRq.GetDataValue(7, key)
            dicInvest['기관순매수'] = self.objRq.GetDataValue(8, key)

            print(value)
            print(dicInvest)


if __name__ == "__main__":
    InitPlusCheck()
    objMarket = Cp7221()
    objMarket.Request()

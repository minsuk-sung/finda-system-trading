# 대신증권 API
# 차트 데이터 구하는 예제

import win32com.client

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 차트 객체 구하기
objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")

objStockChart.SetInputValue(0, 'A005930')  # 종목 코드 - 삼성전자
objStockChart.SetInputValue(1, ord('2'))  # 개수로 조회
objStockChart.SetInputValue(4, 100)  # 최근 100일 치
objStockChart.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
objStockChart.SetInputValue(6, ord('D'))  # '차트 주가 - 일간 차트 요청
objStockChart.SetInputValue(9, ord('1'))  # 수정주가 사용
objStockChart.BlockRequest()

len = objStockChart.GetHeaderValue(3)

print("날짜", "시가", "고가", "저가", "종가", "거래량")
print("빼기빼기==============================================-")

for i in range(len):
    day = objStockChart.GetDataValue(0, i)
    open = objStockChart.GetDataValue(1, i)
    high = objStockChart.GetDataValue(2, i)
    low = objStockChart.GetDataValue(3, i)
    close = objStockChart.GetDataValue(4, i)
    vol = objStockChart.GetDataValue(5, i)
    print(day, open, high, low, close, vol)
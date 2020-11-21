# 대신증권 API
# 데이터 요청 방법 2가지 BlockRequest 와 Request 방식 비교 예제
# 플러스 API 에서 데이터를 요청하는 방법은 크게 2가지가 있습니다
#
# BlockRequest 방식 - 가장 간단하게 데이터 요청해서 수신 가능
# Request 호출 후 Received 이벤트로 수신 받기
#
# 아래는 위 2가지를 비교할 수 있도록 만든 예제 코드입니다

# 일반적인 데이터 요청에는 BlockRequest  방식이 가장 간단합니다
# 다만, BlockRequest  함수 내에서도 동일 하게 메시지펌핑을 하고 있어 해당 통신이 마치기 전에 실시간 시세를 수신 받거나
# 다른 이벤트에 의해 재귀 호출 되는 문제가 있을 경우 함수 호출이 실패할 수 있습니다 
# 복잡한 실시간 시세 수신 중에 통신을 해야 하는 경우에는 Request 방식을 이용해야 합니다.

import pythoncom
from PyQt5.QtWidgets import *
import win32com.client

import win32event

g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')

StopEvent = win32event.CreateEvent(None, 0, 0, None)


class CpEvent:
    def set_params(self, client, name, caller):
        self.client = client  # CP 실시간 통신 object
        self.name = name  # 서비스가 다른 이벤트를 구분하기 위한 이름
        self.caller = caller  # callback 을 위해 보관

    def OnReceived(self):
        # 실시간 처리 - 현재가 주문 체결
        if self.name == 'stockmst':
            print('recieved')
            win32event.SetEvent(StopEvent)
            return


class CpCurReply:
    def __init__(self, objEvent):
        self.name = "stockmst"
        self.obj = objEvent

    def Subscribe(self):
        handler = win32com.client.WithEvents(self.obj, CpEvent)
        handler.set_params(self.obj, self.name, None)


def MessagePump(timeout):
    waitables = [StopEvent]
    while 1:
        rc = win32event.MsgWaitForMultipleObjects(
            waitables,
            0,  # Wait for all = false, so it waits for anyone
            timeout,  # (or win32event.INFINITE)
            win32event.QS_ALLEVENTS)  # Accepts all input

        if rc == win32event.WAIT_OBJECT_0:
            # Our first event listed, the StopEvent, was triggered, so we must exit
            print('stop event')
            break

        elif rc == win32event.WAIT_OBJECT_0 + len(waitables):
            # A windows message is waiting - take care of it. (Don't ask me
            # why a WAIT_OBJECT_MSG isn't defined < WAIT_OBJECT_0...!).
            # This message-serving MUST be done for COM, DDE, and other
            # Windowsy things to work properly!
            print('pump')
            if pythoncom.PumpWaitingMessages():
                break  # we received a wm_quit message
        elif rc == win32event.WAIT_TIMEOUT:
            print('timeout')
            return
            pass
        else:
            print('exception')
            raise RuntimeError("unexpected win32wait return value")


code = 'A005930'

##############################################################
# 1. BlockRequest
print('#####################################')
objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
objStockMst.SetInputValue(0, code)
objStockMst.BlockRequest()
print('BlockRequest 로 수신 받은 데이터')
item = {}
item['종목명'] = g_objCodeMgr.CodeToName(code)
item['현재가'] = objStockMst.GetHeaderValue(11)  # 종가
item['대비'] = objStockMst.GetHeaderValue(12)  # 전일대비
print(item)

print('')
##############################################################
# 2. Request ==> 메시지 펌프 ==>  OnReceived 이벤트 수신
print('#####################################')
objReply = CpCurReply(objStockMst)
objReply.Subscribe()

code = 'A005930'
objStockMst.SetInputValue(0, code)
objStockMst.Request()
MessagePump(10000)
item = {}
item['종목명'] = g_objCodeMgr.CodeToName(code)
item['현재가'] = objStockMst.GetHeaderValue(11)  # 종가
item['대비'] = objStockMst.GetHeaderValue(12)  # 전일대비
print(item)


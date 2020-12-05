import os
import sys
import time
import locale
import warnings
import configparser
from pywinauto import application
from PyQt5 import uic
from PyQt5.QtCore import QTime, QTimer, Qt, QDateTime, QDate
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox, QTableWidgetItem
from daishin import Daishin

locale.setlocale(locale.LC_ALL, 'en_US')
form_class = uic.loadUiType("pytrader.ui")[0]
warnings.filterwarnings(action='ignore')

class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.isLogin = False
        self.user_name = None

        if not self.isLogin:
            self.pushButton_5.clicked.connect(self.login)  # 매수/매도 주문
        else:
            QMessageBox.warning(self, "알림", f"[로그인] 현재 {self.user_name}님으로 로그인된 상태입니다.")
            self.show()

        self.pushButton.clicked.connect(self.send_order) # 매수/매도 주문
        self.pushButton_2.clicked.connect(self.show_portfolio_info) # 현재 보유 종목 조회
        self.pushButton_3.clicked.connect(self.not_implemented)
        self.pushButton_6.clicked.connect(self.logout)  # 매수/매도 주문

    def login(self):
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')

        time.sleep(5)

        config = configparser.ConfigParser()
        config.read('config.ini')

        user_id = config['DEFAULT']['ID']
        user_pwd = config['DEFAULT']['PWD']
        user_pwdcert = config['DEFAULT']['PWD_CERT']

        creon = application.Application()
        creon.start(f'C:\CREON\STARTER\coStarter.exe /prj:cp '
                    f'/id:{user_id} /pwd:{user_pwd} /pwdcert:{user_pwdcert} /autostart')
        time.sleep(30)  # 전체적으로 실행될때까지 시간이 필요

        self.daishin = Daishin(msg_on=False)
        self.show_account_info()
        self.show_portfolio_info()

        self.timer = QTimer(self)
        self.timer.start(500)  # 0.5초
        self.timer.timeout.connect(self.timeout)

        self.timer2 = QTimer(self)
        self.timer2.start(500)
        self.timer2.timeout.connect(self.timeout2)

        self.isLogin = True
        self.user_name = self.daishin.get_account_info()['계좌명']
        self.label_10.setText(f"{self.user_name}님 환영합니다! 오늘도 좋은 하루 되세요!")

    def logout(self):
        os.system('taskkill /IM coStarter* /F /T')
        os.system('taskkill /IM CpStart* /F /T')
        os.system('taskkill /IM DibServer* /F /T')
        os.system('wmic process where "name like \'%coStarter%\'" call terminate')
        os.system('wmic process where "name like \'%CpStart%\'" call terminate')
        os.system('wmic process where "name like \'%DibServer%\'" call terminate')

    def timeout(self):
        market_start_time = QTime(9, 0, 0)
        current_date = QDate.currentDate()
        current_time = QTime.currentTime()

        # 사실 이 부분은 아직 왜 쓰는건지 모르겠네
        # if current_time > market_start_time and self.trade_stocks_done is False:
        #     self.trade_stocks()
        #     self.trade_stocks_done = True

        text_time = current_date.toString(Qt.DefaultLocaleLongDate) + ' ' + current_time.toString('hh:mm:ss')
        time_msg = "현재 날짜 및 시간: " + text_time

        if (current_time >= QTime(9, 0, 0)) and (current_time <= QTime(15, 30, 0)):
            order_cond = "주문가능"
        else:
            order_cond = "주문불가능"
            self.statusbar.setStyleSheet("color: red;")

        state = self.daishin.get_connect_state()
        if state == 1:
            state_msg = f"서버 연결 중({order_cond})"
        else:
            state_msg = "서버 미 연결 중"

        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def timeout2(self):
        if self.checkBox.isChecked():
            self.show_account_info()
            self.show_portfolio_info()

    def change_account_form(self, sAccNo):
        return f"{sAccNo[0:3]}-{sAccNo[3:]}(01)"

    def show_account_info(self):
        acc_info = self.daishin.get_account_info()
        self.lineEdit.setText(self.change_account_form(self.daishin.acc_no))  # 계좌번호 표시
        self.lineEdit_2.setText(str(acc_info['총 평가금액']))  # 총 평가금액 "{:,d}".format(1234567)
        self.lineEdit_3.setText(str(round(acc_info['수익률'], 3)) + '%')  # 수익률
        if acc_info['수익률'] > 0:
            self.lineEdit_3.setStyleSheet("color: red;")
        elif acc_info['수익률'] < 0:
            self.lineEdit_3.setStyleSheet("color: blue;")

        self.lineEdit_5.setText(str(acc_info['주문 가능 금액']))

    def show_portfolio_info(self):
        res = self.daishin.get_my_stocks()
        self.tableWidget_2.setRowCount(len(res))
        self.tableWidget_2.setColumnCount(6)
        for i, (code, info) in enumerate(res.items()):

            # 종목명
            item = QTableWidgetItem(str(info['종목명']))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_2.setItem(i, 0, item)

            # 종목코드
            item = QTableWidgetItem(str(code))
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_2.setItem(i, 1, item)

            # 체결잔고수량
            item = QTableWidgetItem(f"{info['체결잔고수량']:,}")
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_2.setItem(i, 2, item)

            # 현재가
            item = QTableWidgetItem(f"{self.daishin.get_current_data(code)['종가']:,}")
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_2.setItem(i, 3, item)

            # 체결장부단가
            item = QTableWidgetItem(f"{info['체결장부단가']:,.2f}")
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_2.setItem(i, 4, item)

            # 수익률
            item = QTableWidgetItem(f"{info['수익률']:.2f}%")
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.tableWidget_2.setItem(i, 5, item)
            if info['수익률'] >= 0:
                self.tableWidget_2.item(i, 5).setForeground(QColor(255, 0, 0))
            else:
                self.tableWidget_2.item(i, 5).setForeground(QColor(0, 0, 255))

        self.tableWidget_2.resizeRowsToContents()

    def send_order(self):
        current_time = QTime.currentTime()
        if (current_time >= QTime(9, 0, 0)) and (current_time <= QTime(15, 30, 0)):
            order_type_lookup = {'신규매도': 1, '신규매수': 2}
            order_type = self.comboBox_2.currentText()
            code = self.lineEdit_4.text()
            volume = self.spinBox.value()
            self.daishin.sendOrder(f"{order_type_lookup[order_type]}", str(code), int(volume))
            QMessageBox.warning(self, "알림", f"[수동주문] {str(code)}의 {order_type}({int(volume)}주)가 완료되었습니다.")
            self.show()
        else:
            self.spinBox.clear()
            self.spinBox.setValue(0)
            self.spinBox_2.clear()
            self.spinBox_2.setValue(0)
            QMessageBox.warning(self, "경고", "지금은 정규매매장이 종료되었습니다.")
            self.show()

    def not_implemented(self):
        # QT 경고 메세지: https: // dbrang.tistory.com / 948
        # https://has3ong.tistory.com/188
        QMessageBox.warning(self, "경고", "해당 모듈은 아직 개발되지 않았습니다.")
        self.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
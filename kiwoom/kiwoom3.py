from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from config.errorCode import *
from datetime import datetime
import pandas as pd
import csv


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.create_kiwoom_instance()
        self.set_signal_slots()

        ####### event loop를 실행하기 위한 변수 모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.tr_event_loop = QEventLoop()
        #########################################

        self.stock_data = []

    # OCX 방식을 python에 사용할 수 있게 반환해 주는 함수
    def create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 레지스트리에 저장된 API 모듈 불러오기

    # event와 slot 연결
    def set_signal_slots(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.tr_data_slot)  # 트랜잭션 요청 관련 이벤트

    # 로그인 signal 보내기
    def login_signal(self):
        print(f"로그인 시작 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널
        self.login_event_loop.exec_()  # 이벤트 루프 실행

    # 로그인 slot
    def login_slot(self, err_code):
        print(f"로그인 결과: {errors(err_code)} ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
        self.login_event_loop.exit()  # 로그인 완료시 이벤트 루프 종료

    def get_stock_name(self, code):
        stock_name = self.dynamicCall('GetMasterCodeName(str)', code)
        return stock_name

    def set_input_value(self, id, value):
        self.dynamicCall('SetInputValue(str, str)', id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.dynamicCall('CommRqData(str, str, int, str)', rqname, trcode, next, screen_no)
        self.tr_event_loop.exec_()

    def comm_get_data(self, code, real_type, field_name, index, item_name):
        ret = self.dynamicCall('CommGetData(str, str, str, int, str)', code, real_type, field_name, index, item_name)
        return ret.strip()

    def get_repeat_cnt(self, trcode, rqname):
        ret = self.dynamicCall('GetRepeatCnt(str, str)', trcode, rqname)
        return ret

    def get_csv_data(self, file_name, header):
        data = pd.read_csv(file_name, encoding = 'cp949', dtype = object, keep_default_na = False, names = header)
        return data

    def tr_data_slot(self, screen_no, rqname, trcode, record_name, next):
        if rqname == '시작날짜잡기':
            self.opt10081_start_point(rqname, trcode)

        try:
            self.tr_event_loop.exit()
        except AttributeError:
            pass

    def request_start_point(self, stock_code):
        self.set_input_value('종목코드', stock_code)
        self.set_input_value('수정주가구분', 1)
        self.comm_rq_data('시작날짜잡기', 'opt10081', 0, '1001')

    def opt10081_start_point(self, rqname, trcode):
        self.stock_data = []
        cnt = self.get_repeat_cnt(trcode, rqname)

        for i in range(cnt):
            date = self.comm_get_data(trcode, '', rqname, i, '일자')
            open_price = self.comm_get_data(trcode, '', rqname, i, '시가')
            close_price = self.comm_get_data(trcode, '', rqname, i, '현재가')
            low_price = self.comm_get_data(trcode, '', rqname, i, '저가')
            self.stock_data.append((date, int(open_price), int(close_price), int(low_price)))

    def set_start_point(self, data, highest_date):
        for idx, daily_info in enumerate(data):
            if daily_info[0] == highest_date:
                highest_idx = idx
                data = data[highest_idx:]
                break

        if len(data) == 1:
            fibo_start_date, fibo_start_price = data[0][0], data[0][3]
        else:
            fibo_start_date, fibo_start_price = data[0][0], min(data[0][3], data[1][2])
            for i in range(1, min(len(data), 7) - 1):
                rise_percentage = (data[i][2] - data[i + 1][2]) * 100 / data[i + 1][2]
                if rise_percentage >= 10:
                    fibo_start_date = data[i][0]
                    fibo_start_price = min(data[i][3], data[i + 1][2])

        status = '후보' if data[0][0] == fibo_start_date else '상한가 알파'

        return fibo_start_date, fibo_start_price, status, data[0][2]

    def write_file(self, write_data):
        f = open(f"set_start_point_{datetime.now().strftime('%Y-%m-%d')}.csv", 'a', newline='')
        writer = csv.writer(f)
        writer.writerow(write_data)
        f.close()

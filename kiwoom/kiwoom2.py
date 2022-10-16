from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *
from datetime import datetime
import csv
import pandas as pd


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        ####### event loop를 실행하기 위한 변수 모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.start_price_event_loop = QEventLoop()
        #########################################

        ########### 전체 종목 관리
        self.buffer = []
        self.write_buffer = []
        ###########################

        ####### 요청 스크린 번호
        self.screen_deposit = "2000"  # 계좌 관련한 스크린 번호
        self.screen_stock = "4000"  # 계산용 스크린 번호
        ########################################

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 반환해 주는 함수 실행
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음
        self.login_signal()  # 로그인 요청 관련 함수
        self.read_csv_file()
        self.write_file()


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 레지스트리에 저장된 API 모듈 불러오기

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot)  # 트랜잭션 요청 관련 이벤트

    def login_signal(self):
        print(f"로그인 시작 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널
        self.login_event_loop.exec_()  # 이벤트 루프 실행

    def login_slot(self, err_code):
        print(f"로그인 결과: {errors(err_code)} ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
        self.login_event_loop.exit()  # 로그인 완료시 이벤트 루프 종료


    # status 구분
    # 0: 오늘 상한가 달아서 시작 날짜 정해야 할때
    # 10: 아직 매수 안함. 계속 추적
    # 11: 상한가 알파
    # 12: 3시확
    # 2: 매수함
    # 3: 후시세 감시중


    def read_csv_file(self):
        data = pd.read_csv("C:\\Users\\박성훈\\Desktop\\상한가 종목.csv", encoding='cp949', dtype = object)
        for i in range(len(data)):
            QTest.qWait(3600)
            date, stock_name, stock_code = str(data.iloc[i]['날짜']), str(data.iloc[i]['종목명']), str(data.iloc[i]['종목코드']).zfill(6)
            self.buffer = [date, stock_name, stock_code]

            self.dynamicCall("SetInputValue(str, str)", "종목코드", stock_code)
            self.dynamicCall("SetInputValue(str, str)", "기준일자", '20200801')
            self.dynamicCall("SetInputValue(str, str)", "수정주가구분", "1")
            self.dynamicCall("CommRqData(str, str, str, str)", "시작날짜조회", "opt10081", '0', self.screen_stock)

            self.start_price_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

        if sRQName == '시작날짜조회':
            code = self.dynamicCall("GetCommData(str, str, str, str)", sTrCode, sRQName, '0', "종목코드").strip()
            data = []
            cnt = self.dynamicCall("GetRepeatCnt(str, str)", sTrCode, sRQName)

            # for i in range(min(cnt, 7)): # 상한가 찍은 기준봉 1 + 그전 5봉 탐색 + 5봉전 상승률 구하기 1
            for i in range(cnt):
                date = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "일자").strip()
                start_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "시가").strip()
                end_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "현재가").strip()
                low_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "저가").strip()
                data.append((date, int(start_price), int(end_price), int(low_price)))
            print(data)
            fibo_start_date, fibo_start_price = data[0][0], min(data[0][3], data[1][2])
            for i in range(1, min(cnt, 7) - 1):
                rise_percentage = (data[i][2] - data[i+1][2]) * 100 / data[i+1][2]
                if rise_percentage >= 10:
                    fibo_start_date = data[i][0]
                    fibo_start_price = min(data[i][3], data[i+1][2])

            self.buffer.extend([fibo_start_date, fibo_start_price])

            print(self.buffer)
            self.write_buffer.append(self.buffer)
            self.start_price_event_loop.exit()

    def write_file(self):
        f = open(f"out2_{datetime.now().strftime('%Y-%m-%d')}.csv", 'w', newline='')
        writer = csv.writer(f)
        writer.writerows(self.write_buffer)
        f.close()

    def stop_screen_cancel(self, sScrNo = None):
        self.dynamicCall("DisconnectRealData(str)", sScrNo)  # 스크린 번호 연결 끊기



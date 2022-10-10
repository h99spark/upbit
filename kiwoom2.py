from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *
from datetime import datetime
import csv


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        ####### event loop를 실행하기 위한 변수 모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.get_deposit_event_loop = QEventLoop()  # 예수금 요청용 이벤트 루프
        self.calculator_event_loop = QEventLoop()
        #########################################

        ########### 전체 종목 관리
        self.all_stocks = []
        self.buffer = []
        ###########################

        ####### 계좌 관련된 변수
        self.account_num = ''  # 계좌번호
        self.deposit = 0  # 예수금

        ####### 요청 스크린 번호
        self.screen_deposit = "2000"  # 계좌 관련한 스크린 번호
        self.screen_stock = "4000"  # 계산용 스크린 번호
        self.screen_real_stock = "5000"  # 종목별 할당할 스크린 번호
        self.screen_meme_stock = "6000"  # 종목별 할당할 주문용 스크린 번호
        ########################################

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 반환해 주는 함수 실행
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음

        self.login_signal()  # 로그인 요청 관련 함수

        self.get_code_list()
        self.get_daily_data()
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

    def get_code_list(self):
        code_list_kospi = self.dynamicCall("GetCodeListByMarket(str)", '0')
        code_list_kosdaq = self.dynamicCall("GetCodeListByMarket(str)", '10')
        kospi_list = code_list_kospi.split(';')[:-1]
        kosdaq_list = code_list_kosdaq.split(';')[:-1]

        for stock_code in kospi_list:
            stock_name = self.dynamicCall("GetMasterCodeName(str)", stock_code)
            self.all_stocks.append((stock_code, stock_name))

        for stock_code in kosdaq_list:
            stock_name = self.dynamicCall("GetMasterCodeName(str)", stock_code)
            self.all_stocks.append((stock_code, stock_name))

        print(len(self.all_stocks))


    def get_daily_data(self):
        idx = 1
        for code, name in self.all_stocks:
            QTest.qWait(4000)
            print(idx, code, name)
            self.dynamicCall("SetInputValue(str, str)", "종목코드", code)
            self.dynamicCall("SetInputValue(str, str)", "기준일자", None)
            self.dynamicCall("SetInputValue(str, str)", "수정주가구분", "1")
            self.dynamicCall("CommRqData(str, str, str, str)", "주식일봉차트조회", "opt10081", '0', self.screen_stock)
            idx += 1

            self.calculator_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

        if sRQName == '주식일봉차트조회':
            code = self.dynamicCall("GetCommData(str, str, str, str)", sTrCode, sRQName, '0', "종목코드").strip()
            name = self.dynamicCall("GetMasterCodeName(str)", code)
            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            data = []

            try:
                for i in range(cnt):
                    date = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "일자").strip()
                    start_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "시가").strip()
                    end_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "현재가").strip()
                    low_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "저가").strip()
                    data.append((date, int(start_price), int(end_price), int(low_price)))

                for i in range(cnt - 1):
                    rise_percentage = (data[i][2] - data[i+1][2]) * 100 / data[i+1][2]
                    if rise_percentage >= 29:
                        print(name, data[i][0], rise_percentage)
                        self.buffer.append([data[i][0], name, code])
            except:
                pass

            self.calculator_event_loop.exit()

    def write_file(self):
        f = open(f"out2_{datetime.now().strftime('%Y-%m-%d')}.csv", 'a', newline='')
        writer = csv.writer(f)
        writer.writerows(self.buffer)
        f.close()

    def stop_screen_cancel(self, sScrNo = None):
        self.dynamicCall("DisconnectRealData(str)", sScrNo)  # 스크린 번호 연결 끊기



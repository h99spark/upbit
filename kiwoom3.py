from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *
from PyQt5.QtTest import *
from datetime import datetime
import pandas as pd


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()

        ####### event loop를 실행하기 위한 변수 모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.get_deposit_event_loop = QEventLoop()  # 예수금 요청용 이벤트 루프
        self.calculator_event_loop = QEventLoop()
        #########################################

        ########### 전체 종목 관리
        self.all_stock_dict = {}
        self.all_stocks = {}
        ###########################

        ####### 계좌 관련된 변수
        self.account_num = ''  # 계좌번호
        self.deposit = 0  # 예수금

        self.total_profit_loss_money = 0  # 총평가손익금액
        self.total_profit_loss_rate = 0.0  # 총수익률(%)

        self.account_stock_dict = {}
        self.not_account_stock_dict = {}
        ########################################

        ######## 종목 정보 가져오기
        self.portfolio_stock_dict = {}
        ########################

        ########### 종목 분석 용
        self.calcul_data = []
        ##########################################

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
        self.read_csv_file()

        # self.detail_account_mystock()  # 계좌평가잔고내역 가져오기
        # QTimer.singleShot(5000, self.not_concluded_account)  # 5초 뒤에 미체결 종목들 가져오기 실행
        # #########################################
        #
        # QTest.qWait(10000)
        # self.read_code()
        # self.screen_number_setting()
        #
        # QTest.qWait(10000)

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

    def read_csv_file(self):
        data = pd.read_csv("C:\\Users\\박성훈\\PycharmProjects\\pythonProject\\out2_2022-10-10.csv", encoding='cp949')
        for i in range(len(data)):
            QTest.qWait(3600)
            date, stock_name, stock_code = str(data.iloc[i]['날짜']), str(data.iloc[i]['종목명']), str(data.iloc[i]['종목코드']).zfill(6)
            print(stock_name, stock_code, date, end = ' ')
            self.dynamicCall("SetInputValue(str, str)", "종목코드", stock_code)
            self.dynamicCall("SetInputValue(str, str)", "기준일자", date)
            self.dynamicCall("SetInputValue(str, str)", "수정주가구분", "1")
            self.dynamicCall("CommRqData(str, str, str, str)", "주식일봉차트조회", "opt10081", '0', self.screen_stock)

            self.calculator_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

        if sRQName == '주식일봉차트조회':
            code = self.dynamicCall("GetCommData(str, str, str, str)", sTrCode, sRQName, '0', "종목코드").strip()
            data = []
            for i in range(7): # 상한가 찍은 기준봉 1 + 그전 5봉 탐색 + 5봉전 상승률 구하기 1
                date = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "일자").strip()
                start_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "시가").strip()
                end_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "현재가").strip()
                low_price = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "저가").strip()
                volume = self.dynamicCall("GetCommData(str, str, int, str)", sTrCode, sRQName, i, "거래대금").strip()
                data.append((date, int(start_price), int(end_price), int(low_price), int(volume)))

            start_date, fibo_start_price = data[0][0], min(data[0][3], data[1][2])
            for i in range(1, 6):
                rise_percentage = (data[i][2] - data[i+1][2]) * 100 / data[i+1][2]
                if rise_percentage >= 10:
                    start_date = data[i][0]
                    fibo_start_price = min(data[i][3], data[i+1][2])

            print(f'start_date: {start_date}, start price: {fibo_start_price}')

            self.calculator_event_loop.exit()


        # elif sRQName == "계좌평가잔고내역요청":
        #     total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0,
        #                                        "총매입금액")  # 출력 : 000000000746100
        #     self.total_buy_money = int(total_buy_money)
        #     total_profit_loss_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
        #                                                0, "총평가손익금액")  # 출력 : 000000000009761
        #     self.total_profit_loss_money = int(total_profit_loss_money)
        #     total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
        #                                               0, "총수익률(%)")  # 출력 : 000000001.31
        #     self.total_profit_loss_rate = float(total_profit_loss_rate)
        #
        #     print(
        #         "계좌평가잔고내역요청 싱글데이터 : %s - %s - %s" % (total_buy_money, total_profit_loss_money, total_profit_loss_rate))
        #
        #     rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
        #
        #     for i in range(rows):
        #         code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
        #         code = code.strip()[1:]
        #
        #         code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
        #         stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
        #                                           "보유수량")
        #         buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
        #         learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
        #                                       "수익률(%)")
        #         current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
        #                                          "현재가")
        #         total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
        #                                                i, "매입금액")
        #         possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
        #                                              "매매가능수량")
        #
        #         print("종목번호: %s - 종목명: %s - 보유수량: %s - 매입가:%s - 수익률: %s - 현재가: %s" % (
        #             code, code_nm, stock_quantity, buy_price, learn_rate, current_price))
        #
        #         if code in self.account_stock_dict:
        #             pass
        #         else:
        #             self.account_stock_dict[code] = {}
        #
        #         code_nm = code_nm.strip()
        #         stock_quantity = int(stock_quantity.strip())
        #         buy_price = int(buy_price.strip())
        #         learn_rate = float(learn_rate.strip())
        #         current_price = int(current_price.strip())
        #         total_chegual_price = int(total_chegual_price.strip())
        #         possible_quantity = int(possible_quantity.strip())
        #
        #         self.account_stock_dict[code].update({"종목명": code_nm})
        #         self.account_stock_dict[code].update({"보유수량": stock_quantity})
        #         self.account_stock_dict[code].update({"매입가": buy_price})
        #         self.account_stock_dict[code].update({"수익률(%)": learn_rate})
        #         self.account_stock_dict[code].update({"현재가": current_price})
        #         self.account_stock_dict[code].update({"매입금액": total_chegual_price})
        #         self.account_stock_dict[code].update({"매매가능수량": possible_quantity})
        #
        #     print("sPreNext : %s" % sPrevNext)
        #     print("계좌에 가지고 있는 종목은 %s " % rows)
        #
        #     if sPrevNext == "2":
        #         self.detail_account_mystock(sPrevNext="2")
        #     else:
        #         self.detail_account_info_event_loop.exit()
    #
    #     elif sRQName == "실시간미체결요청":
    #         rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
    #
    #         for i in range(rows):
    #             code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
    #
    #             code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
    #             order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
    #             order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                             "주문상태")  # 접수,확인,체결
    #             order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                               "주문수량")
    #             order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                            "주문가격")
    #             order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                            "주문구분")  # -매도, +매수, -매도정정, +매수정정
    #             not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                             "미체결수량")
    #             ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                            "체결량")
    #
    #             code = code.strip()
    #             code_nm = code_nm.strip()
    #             order_no = int(order_no.strip())
    #             order_status = order_status.strip()
    #             order_quantity = int(order_quantity.strip())
    #             order_price = int(order_price.strip())
    #             order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
    #             not_quantity = int(not_quantity.strip())
    #             ok_quantity = int(ok_quantity.strip())
    #
    #             if order_no in self.not_account_stock_dict:
    #                 pass
    #             else:
    #                 self.not_account_stock_dict[order_no] = {}
    #
    #             self.not_account_stock_dict[order_no].update({'종목코드': code})
    #             self.not_account_stock_dict[order_no].update({'종목명': code_nm})
    #             self.not_account_stock_dict[order_no].update({'주문번호': order_no})
    #             self.not_account_stock_dict[order_no].update({'주문상태': order_status})
    #             self.not_account_stock_dict[order_no].update({'주문수량': order_quantity})
    #             self.not_account_stock_dict[order_no].update({'주문가격': order_price})
    #             self.not_account_stock_dict[order_no].update({'주문구분': order_gubun})
    #             self.not_account_stock_dict[order_no].update({'미체결수량': not_quantity})
    #             self.not_account_stock_dict[order_no].update({'체결량': ok_quantity})
    #
    #             print("미체결 종목 : %s " % self.not_account_stock_dict[order_no])
    #
    #         self.detail_account_info_event_loop.exit()
    #
    #     elif sRQName == "주식일봉차트조회":
    #         code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
    #         code = code.strip()
    #         # data = self.dynamicCall("GetCommDataEx(QString, QString)", sTrCode, sRQName)
    #
    #         cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
    #         print("남은 일자 수 %s" % cnt)
    #
    #         for i in range(cnt):
    #             data = []
    #
    #             current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                              "현재가")  # 출력 : 000070
    #             value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                      "거래량")  # 출력 : 000070
    #             trading_value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                              "거래대금")  # 출력 : 000070
    #             date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                     "일자")  # 출력 : 000070
    #             start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                            "시가")  # 출력 : 000070
    #             high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                           "고가")  # 출력 : 000070
    #             low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
    #                                          "저가")  # 출력 : 000070
    #
    #             data.append("")
    #             data.append(current_price.strip())
    #             data.append(value.strip())
    #             data.append(trading_value.strip())
    #             data.append(date.strip())
    #             data.append(start_price.strip())
    #             data.append(high_price.strip())
    #             data.append(low_price.strip())
    #             data.append("")
    #
    #             self.calcul_data.append(data.copy())
    #
    #         if sPrevNext == "2":
    #             self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
    #         else:
    #             print("총 일수 %s" % len(self.calcul_data))
    #
    #             pass_success = False
    #
    #             # 120일 이평선을 그릴만큼의 데이터가 있는지 체크
    #             if self.calcul_data == None or len(self.calcul_data) < 120:
    #                 pass_success = False
    #             else:
    #                 # 120일 이평선의 최근 가격 구함
    #                 total_price = 0
    #                 for value in self.calcul_data[:120]:
    #                     total_price += int(value[1])
    #                 moving_average_price = total_price / 120
    #
    #                 # 오늘자 주가가 120일 이평선에 걸쳐있는지 확인
    #                 bottom_stock_price = False
    #                 check_price = None
    #                 if int(self.calcul_data[0][7]) <= moving_average_price and moving_average_price <= int(
    #                         self.calcul_data[0][6]):
    #                     print("오늘의 주가가 120 이평선에 걸쳐있는지 확인")
    #                     bottom_stock_price = True
    #                     check_price = int(self.calcul_data[0][6])
    #
    #                 # 과거 일봉 데이터를 조회하면서 120일 이동평균선보다 주가가 계속 밑에 존재하는지 확인
    #                 prev_price = None
    #                 if bottom_stock_price == True:
    #                     moving_average_price_prev = 0
    #                     price_top_moving = False
    #                     idx = 1
    #
    #                     while True:
    #                         if len(self.calcul_data[idx:]) < 120:  # 120일 치가 있는지 계속 확인
    #                             print("120일 치가 없음")
    #                             break
    #
    #                         total_price = 0
    #                         for value in self.calcul_data[idx:120 + idx]:
    #                             total_price += int(value[1])
    #                         moving_average_price_prev = total_price / 120
    #
    #                         if moving_average_price_prev <= int(self.calcul_data[idx][6]) and idx <= 20:
    #                             print("20일 동안 주가가 120일 이평선과 같거나 위에 있으면 조건 통과 못 함")
    #                             price_top_moving = False
    #                             break
    #
    #                         elif int(self.calcul_data[idx][
    #                                      7]) > moving_average_price_prev and idx > 20:  # 120일 이평선 위에 있는 구간 존재
    #                             print("120일치 이평선 위에 있는 구간 확인됨")
    #                             price_top_moving = True
    #                             prev_price = int(self.calcul_data[idx][7])
    #                             break
    #
    #                         idx += 1
    #
    #                     # 해당부분 이평선이 가장 최근의 이평선 가격보다 낮은지 확인
    #                     if price_top_moving == True:
    #                         if moving_average_price > moving_average_price_prev and check_price > prev_price:
    #                             print("포착된 이평선의 가격이 오늘자 이평선 가격보다 낮은 것 확인")
    #                             print("포착된 부분의 일봉 저가가 오늘자 일봉의 고가보다 낮은지 확인")
    #                             pass_success = True
    #
    #             if pass_success == True:
    #                 print("조건부 통과됨")
    #
    #                 code_nm = self.dynamicCall("GetMasterCodeName(QString)", code)
    #
    #                 f = open("files/condition_stock.txt", "a", encoding="utf8")
    #                 f.write("%s\t%s\t%s\n" % (code, code_nm, str(self.calcul_data[0][1])))
    #                 f.close()
    #
    #             elif pass_success == False:
    #                 print("조건부 통과 못 함")
    #
    #             self.calcul_data.clear()
    #             self.calculator_event_loop.exit()

    def stop_screen_cancel(self, sScrNo = None):
        self.dynamicCall("DisconnectRealData(str)", sScrNo)  # 스크린 번호 연결 끊기


    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code)
        code_list = code_list.split(';')[:-1]
        return code_list

    def calculator_fnc(self):
        code_list = self.get_code_list_by_market("10")

        print("코스닥 갯수 %s " % len(code_list))

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)  # 스크린 연결 끊기

            print("%s / %s : KOSDAQ Stock Code : %s is updating... " % (idx + 1, len(code_list), code))
            self.day_kiwoom_db(code=code)

    # def day_kiwoom_db(self, code=None, date=None, sPrevNext="0"):
    #     QTest.qWait(3600)  # 3.6초마다 딜레이를 준다.
    #
    #     self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
    #     self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")
    #
    #     if date != None:
    #         self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)
    #
    #     self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식일봉차트조회", "opt10081", sPrevNext,
    #                      self.screen_calculation_stock)
    #
    #     self.calculator_event_loop.exec_()
    #
    # def read_code(self):
    #     if os.path.exists("files/condition_stock.txt"):  # 해당 경로에 파일이 있는지 체크한다.
    #         f = open("files/condition_stock.txt", "r", encoding="utf8")
    #
    #         lines = f.readlines()  # 파일에 있는 내용들이 모두 읽어와 진다.
    #         for line in lines:  # 줄바꿈된 내용들이 한줄 씩 읽어와진다.
    #             if line != "":
    #                 ls = line.split("\t")
    #
    #                 stock_code = ls[0]
    #                 stock_name = ls[1]
    #                 stock_price = int(ls[2].split("\n")[0])
    #                 stock_price = abs(stock_price)
    #
    #                 self.portfolio_stock_dict.update({stock_code: {"종목명": stock_name, "현재가": stock_price}})
    #         f.close()
    #
    # def merge_dict(self):
    #     self.all_stock_dict.update({"계좌평가잔고내역": self.account_stock_dict})
    #     self.all_stock_dict.update({'미체결종목': self.not_account_stock_dict})
    #     self.all_stock_dict.update({'포트폴리오종목': self.portfolio_stock_dict})
    #
    # def screen_number_setting(self):
    #     screen_overwrite = []
    #
    #     # 계좌평가잔고내역에 있는 종목들
    #     for code in self.account_stock_dict.keys():
    #         if code not in screen_overwrite:
    #             screen_overwrite.append(code)
    #
    #     # 미체결에 있는 종목들
    #     for order_number in self.not_account_stock_dict.keys():
    #         code = self.not_account_stock_dict[order_number]['종목코드']
    #
    #         if code not in screen_overwrite:
    #             screen_overwrite.append(code)
    #
    #     # 포트폴리오에 있는 종목들
    #     for code in self.portfolio_stock_dict.keys():
    #         if code not in screen_overwrite:
    #             screen_overwrite.append(code)
    #
    #     # 스크린 번호 할당
    #     cnt = 0
    #     for code in screen_overwrite:
    #         temp_screen = int(self.screen_real_stock)
    #         meme_screen = int(self.screen_meme_stock)
    #
    #         if (cnt % 50) == 0:
    #             temp_screen += 1
    #             self.screen_real_stock = str(temp_screen)
    #
    #         if (cnt % 50) == 0:
    #             meme_screen += 1
    #             self.screen_meme_stock = str(meme_screen)
    #
    #         if code in self.portfolio_stock_dict.keys():
    #             self.portfolio_stock_dict[code].update({"스크린번호": str(self.screen_real_stock)})
    #             self.portfolio_stock_dict[code].update({"주문용스크린번호": str(self.screen_meme_stock)})
    #
    #         elif code not in self.portfolio_stock_dict.keys():
    #             self.portfolio_stock_dict.update(
    #                 {code: {"스크린번호": str(self.screen_real_stock), "주문용스크린번호": str(self.screen_meme_stock)}})
    #
    #         cnt += 1
    #
    #     print(self.portfolio_stock_dict)

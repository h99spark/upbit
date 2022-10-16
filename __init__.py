import sys
from kiwoom.kiwoom3 import *
from PyQt5.QtWidgets import *

if __name__ == "__main__":
    app = QApplication(sys.argv)  # PyQt5로 실행할 파일명을 자동 설정
    kiwoom = Kiwoom()  # 키움 클래스 객체화
    kiwoom.login_signal()  # 로그인 시도

    file_name = "C:\\Users\\박성훈\\PycharmProjects\\pythonProject\\시작(2022-10-16).csv"
    header = ['날짜', '종목명', '종목코드', '상태', '시작날짜', '시작가격', '최고날짜', '최고가격']
    kiwoom.all_stocks = kiwoom.get_csv_data(file_name, header)

    for i in range(len(kiwoom.all_stocks)):
        row = kiwoom.all_stocks.iloc[i]
        highest_date, stock_name, stock_code = row['날짜'], row['종목명'], row['종목코드'].zfill(6)
        status, start_date, start_price = row['상태'], row['시작날짜'], row['시작가격']

        # if not status:
        #     QTest.qWait(3600)
        #     kiwoom.request_start_point(stock_code)
        #     start_date, start_price, status, highest_price = kiwoom.set_start_point(kiwoom.stock_data, highest_date)
        # elif status == '후보':
        #     QTest.qWait(3600)
        QTest.qWait(3600)
        kiwoom.request_start_point(stock_code)
        start_date, start_price, status, highest_price = kiwoom.set_start_point(kiwoom.stock_data, highest_date)


        print(f'{stock_name}({stock_code}) : {highest_date} // {start_date}({start_price}) // {highest_date}({highest_price})')
        kiwoom.write_file([highest_date, stock_name, stock_code, status, start_date, start_price, highest_date, highest_price])




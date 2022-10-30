import win32com.client
import pandas as pd
import time

class GetData:

    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.StockChart")
        self.obj2 = win32com.client.Dispatch("CpUtil.CpCybos")

        self.all_stocks = {}

        self.row = []
        self.header = ['날짜', '시간', '시가', '고가', '저가', '종가']
        self.df = None

    # 파일 읽어서 모든 종목명, 종목코드 가져와서 저장
    def read_file(self, file_name):
        data_type = {"날짜": 'str', '종목명': 'str', '종목코드': 'str'}
        df = pd.read_csv(file_name, encoding='utf-8', dtype=data_type)

        for i in range(len(df)):
            name, code = df.iloc[i]['종목명'], 'A' + df.iloc[i]['종목코드'].zfill(6)
            if name not in self.all_stocks:
                self.all_stocks[name] = code

        print(f"총 종목 수: {len(self.all_stocks)}")

    # 1초당 요청 개수 제한 걸리지 않게 쉬었다 가기
    def request_refill(self):
        remain_request_count = self.obj2.GetLimitRemainCount(1)
        if remain_request_count == 0:
            print('남은 요청이 모두 소진되었습니다. 잠시 대기합니다.')

            while True:
                time.sleep(1)
                remain_request_count = self.obj2.GetLimitRemainCount(1)
                if remain_request_count > 0:
                    print(f'작업을 재개합니다. (남은 요청 : {remain_request_count}')
                    return

    # 종목코드로 분봉 가져와서 dataframe에 저장
    def get_ohlc(self, code, minute, num):

        self.obj.SetInputValue(0, code) # 종목코드
        self.obj.SetInputValue(1, ord('2')) # 1: 기간, 2: 개수로 받기
        # obj.objStockChart.SetInputValue(2, toDate)  # To 날짜
        # obj.objStockChart.SetInputValue(3, fromDate)  # From 날짜
        self.obj.SetInputValue(4, num)  # 요청개수
        self.obj.SetInputValue(5, [0, 1, 2, 3, 4, 5])  # 0: 날짜, 1: 시간, 2~5: ohlc
        self.obj.SetInputValue(6, ord('m')) # 분봉 차트
        self.obj.SetInputValue(7, minute)  # 5분 간격
        self.obj.SetInputValue(9, ord('1'))  # 1: 수정주가

        totlen = 0
        row = []
        while True:

            self.obj.BlockRequest()
            rqStatus = self.obj.GetDibStatus()
            rqRet = self.obj.GetDibMsg1()
            if rqStatus != 0:
                print("통신상태", rqStatus, rqRet)
                exit()

            cnt = self.obj.GetHeaderValue(3)
            totlen += cnt
            for i in range(cnt):
                date = self.obj.GetDataValue(0, i)
                time = str(self.obj.GetDataValue(1, i)).zfill(4)
                time = time[:2] + ':' + time[2:]
                open = self.obj.GetDataValue(2, i)
                high = self.obj.GetDataValue(3, i)
                low = self.obj.GetDataValue(4, i)
                close = self.obj.GetDataValue(5, i)
                row.append([date, time, open, high, low, close])

            if not self.obj.Continue:
                break
            if totlen > num:
                break

        row.sort(key = lambda x: (x[0], x[1]))
        self.df = pd.DataFrame(row, columns=self.header)

    # dataframe을 csv 파일에 저장
    def write_csvfile(self, name):
        print(f'파일 쓰는 중: ./분봉/{name}.csv 크기: {self.len(df)}')
        self.df.to_csv(f"./분봉/{name}.csv", mode='w', index=None, encoding='utf-8-sig')


if __name__ == '__main__':
    cb = GetData()

    cb.read_file('C:\\Users\\박성훈\\Desktop\\상한가 종목(22년하).csv')

    for name, code in cb.all_stocks.items():

        cb.request_refill()

        # 종목코드, 몇분봉, 몇개나
        cb.get_ohlc(code, 5, 30000)

        cb.write_csvfile(name)
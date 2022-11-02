import pandas as pd
import csv
import os

class ThreeExtension:

    def __init__(self, file_name, name, code, start_date, start_price, standard_date):

        self.point_bought = [False, False, False] # 5타점 샀는지 여부, 6타점 샀는지 여부, 7타점 샀는지 여부
        self.points = [0, 0, 0, 0, 0] # 4타점, 5타점, 6타점, 7타점, 8타점

        self.name = name
        self.code = code

        self.is_sold = False # 팔았는지 여부

        self.file_name = file_name
        self.dates = [] # 날짜 모음
        self.info = {} # 분봉 데이터 모음(시각, 시가, 고가, 저가, 종가)

        self.start_date = start_date # 피보나치 시작 날짜
        self.highest_date = standard_date # 피보나치 최고 날짜
        self.standard_date = standard_date # 상한가 날짜
        self.start_price = start_price # 피보나치 시작 가격
        self.highest_price = start_price # 피보나치 최고 가격

        self.buy_price = 0.0
        self.sell_price = 0.0

        # 0: 후보, 1: 상한가 알파, 2: 3시확
        self.status = 0 if self.start_date == self.highest_date else 1

        self.result = [self.standard_date, self.name, self.code]

        print("")
        print(file_name, start_date, start_price, date)

    # csv 파일 불러와서 데이터를 딕셔너리에 저장
    def set_data(self):
        data_type = {"날짜": 'str', '시간': 'str', '시가': 'int', '고가': 'int', '저가': 'int', '종가': 'int'}
        df = pd.read_csv(self.file_name, encoding='utf-8', dtype=data_type)
        self.dates = sorted(list(set(df['날짜'])))

        for i in range(len(df)):
            row = df.iloc[i]
            date, time = row['날짜'].replace('/', ''), row['시간']
            open, high, low, close = row['시가'], row['고가'], row['저가'], row['종가']

            if date not in self.info:
                self.info[date] = [(time, open, high, low, close)]
            else:
                self.info[date].append((time, open, high, low, close))

        for key in self.info.keys():
            self.info[key].sort(key=lambda x: x[0])

    # 타점 업데이트
    def update_points(self):
        self.points[0] = self.start_price + (self.highest_price - self.start_price) * 0.6
        self.points[1] = self.start_price + (self.highest_price - self.start_price) * 0.5
        self.points[2] = self.start_price + (self.highest_price - self.start_price) * 0.4
        self.points[3] = self.start_price + (self.highest_price - self.start_price) * 0.3
        self.points[4] = self.start_price + (self.highest_price - self.start_price) * 0.2

    # 상한가 단 날 피보나치 초기화하기
    def init_fibo(self):
        self.highest_price = self.info[self.standard_date][-1][2]
        self.update_points()

    # 상한가 다음 날부터 하루씩 검사
    def traverse_dates(self):
        standard_idx = self.dates.index(self.standard_date)
        for date_idx in range(standard_idx + 1, len(self.dates)):
            if not self.traverse_daily(self.dates[date_idx]):
                return

    # 하루치 분봉들 검사
    def traverse_daily(self, date):

        for time, open, high, low, close in self.info[date]:

            self.extend_fibo(date, high)

            if not self.buy_decision(low, date, time):
                self.sell_decision(low, high, date, time)

        return False if self.examine_rotten(date) or self.finish_after_look(date) else True

    # 고점 갱신하면 피보나치 확장. 추가로 후보면 상한가 알파 승격, 상한가 알파면 3시확 승격
    def extend_fibo(self, date, high):
        highest_idx = self.dates.index(self.highest_date)
        cur_idx = self.dates.index(date)

        if high > self.highest_price:
            self.highest_date, self.highest_price = date, high
            self.update_points()

            if self.status == 0:
                self.status = 1
                print(f'{date}: 상한가 알파 달성')
            elif self.status == 1 and highest_idx != cur_idx:
                self.status = 2
                print(f'{date}: 3시확 달성')

    # 5일 이상 고점 갱신 실패하면 종료(후보) or 5일 이상 자리 안오면 종료(상한가 알파) or 3주 이상 자리 안오면 종료(3시확)
    def examine_rotten(self, date):
        highest_idx = self.dates.index(self.highest_date)
        cur_idx = self.dates.index(date)

        if self.point_bought == [False, False, False]:
            if self.status == 0 and highest_idx + 4 <= cur_idx:
                print(f'{date}: 고점 갱신 실패')
                return True
            elif self.status == 1 and highest_idx + 4 <= cur_idx:
                print(f'{date}: 5일째 자리 안 옴')
                return True
            elif self.status == 2 and highest_idx + 14 <= cur_idx:
                print(f'{date}: 3주째 자리 안 옴')
                return True

        return False

    # 매도 후 고점 5일 후 지나면 종료
    # TODO: 매도 후 고점 5일 이내에 고점 갱신하면 3시확으로 후시세 감시
    def finish_after_look(self, date):
        highest_idx = self.dates.index(self.highest_date)
        cur_idx = self.dates.index(date)

        if self.is_sold:
            if highest_idx + 4 <= cur_idx:
                print(f'{date}: 후시세 감시 종료')
                return True
            # elif highest_idx == cur_idx:
            #     self.is_sold = False
            #     self.point_bought = [False, False, False]
            #     return False

        return False

    # 매수 결정
    # TODO: 정확한 매수 가격(호가 반영)
    def buy_decision(self, low, date, time):
        buy = False

        # if self.status in (1, 2) and not self.is_sold:
        if self.status == 2 and not self.is_sold:
            # 5타점 매수
            if self.point_bought == [False, False, False] and low <= self.points[1]:
                self.point_bought[0] = True
                self.buy_price = self.points[1]
                print(f'{date} {time} - 5타점 매수(평단: {self.buy_price})')
                buy = True

            # 6타점 매수
            if self.point_bought == [True, False, False] and low <= self.points[2]:
                self.point_bought[1] = True
                self.buy_price = (self.points[1] + self.points[2]) / 2
                print(f'{date} {time} - 6타점 매수(평단: {self.buy_price})')
                buy = True

            # 7타점 매수
            if self.point_bought == [True, True, False] and low <= self.points[3]:
                self.point_bought[2] = True
                self.buy_price = self.points[2]
                print(f'{date} {time} - 7타점 매수(평단: {self.buy_price})')
                buy = True

        return buy

    # 매도 결정
    # TODO: 정확한 매도 가격(호가 반영)
    def sell_decision(self, low, high, date, time):

        # 5타점 매수 후 4타점 매도
        if self.point_bought == [True, False, False] and not self.is_sold and self.points[0] <= high:
            self.is_sold = True
            self.sell_price = self.points[0]
            print(f'{date} {time} - 4타점 매도(이익) (매도단가: {self.sell_price})')
            print(f'수익률: {(self.sell_price - self.buy_price) * 100/ self.buy_price:.2f}%')

        # 5, 6타점 매수 후 4.5타점 매도
        elif self.point_bought == [True, True, False] and not self.is_sold and (self.points[0] + self.points[1]) / 2 <= high:
            self.is_sold = True
            self.sell_price = (self.points[0] + self.points[1]) / 2
            print(f'{date} {time} - 4.5타점 매도(이익) (매도단가: {self.sell_price})')
            print(f'수익률: {(self.sell_price - self.buy_price) * 100/ self.buy_price:.2f}%')

        # 5, 6, 7타점 매수 후 6타점 매도
        elif self.point_bought == [True, True, True] and not self.is_sold and self.points[2] <= high:
            self.is_sold = True
            self.sell_price = self.points[2]
            print(f'{date} {time} - 6타점 매도(본전) (매도단가: {self.sell_price})')
            print(f'수익률: {(self.sell_price - self.buy_price) * 100/ self.buy_price:.2f}%')

        # 5, 6, 7타점 매수 후 8타점 매도
        elif self.point_bought == [True, True, True] and not self.is_sold and low <= self.points[4]:
            self.is_sold = True
            self.sell_price = self.points[4]
            print(f'{date} {time} - 8타점 매도(손해) (매도단가: {self.sell_price})')
            print(f'수익률: {(self.sell_price - self.buy_price) * 100/ self.buy_price:.2f}%')

    def write_result(self):
        if self.status != 2:
            return
        else:
            self.result.append('3시확')

        self.result.extend((self.start_date, self.start_price, self.highest_date, self.highest_price))

        if not self.is_sold:
            if self.point_bought == [False, False, False]:
                self.result.extend((' ', ' ', ' ', 'O')) # 자리안옴
            else:
                self.result.extend((' ', ' ', ' ', ' ')) # 진행중
        else:
            if self.point_bought == [True, False, False] or self.point_bought == [True, True, False]:
                self.result.extend(('O', ' ', ' ', ' ')) # 이득
            else:
                if self.buy_price > self.sell_price:
                    self.result.extend((' ', ' ', 'O', ' ')) # 손해
                else:
                    self.result.extend((' ', 'O', ' ', '')) # 본전

        print(self.result)
        with open('./3시확 통계.csv', 'a', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(self.result)


if __name__ == '__main__':

    ###################################################################
    read_file = 'C:\\Users\\User\\Desktop\\상한가 종목(22년 하반기).csv'
    write_file = './3시확 통계.csv'
    ###################################################################

    with open(write_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['날짜', '종목명', '종목코드', '기법', '시작날짜', '시작가격', '최고날짜', '최고가격', '이득', '본전', '손해', '자리안옴'])

    data_type = {"날짜": 'str', '종목명': 'str', '종목코드': 'str', '시작날짜': 'str', '시작가격': 'int'}
    df = pd.read_csv(read_file, encoding='utf-8', dtype=data_type)

    for i in range(len(df)):
        date = df.iloc[i]['날짜']
        name, code = df.iloc[i]['종목명'], df.iloc[i]['종목코드']
        start_date, start_price = df.iloc[i]['시작날짜'], df.iloc[i]['시작가격']
        file_name = f'./분봉\\{name}.csv'

        te = ThreeExtension(file_name, name, code, start_date, start_price, date)
        te.set_data()
        te.init_fibo()
        te.traverse_dates()
        te.write_result()
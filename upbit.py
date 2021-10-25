import pyupbit
import time
import requests
from bs4 import BeautifulSoup
from datetime import datetime

access = "xHN0w0UsB8a09wdYZXrurJ5xMXxsdv1i8bW6uOh9"
secret = "Ig674PmtvONPAqy7GLxVWVUe9Um4OR7rxmTNgVg9"

upbit = pyupbit.Upbit(access, secret)


# 거래대금 상위 30개만 뽑기
def transaction_top():
    coin_array = []
    url = "https://www.coingecko.com/ko/거래소/upbit"
    resp = requests.get(url)

    bs = BeautifulSoup(resp.text, 'html.parser')
    selector = "tbody > tr > td > a"
    columns = bs.select(selector)

    ticker_in_krw = [x.text.strip() for x in columns if x.text.strip()[-3:] == "KRW"]
    for ticker in ticker_in_krw:
        coin = "KRW-" + ticker.split('/')[0]
        coin_array.append(coin)

    return coin_array[:30]


def price_range(coin):
    data_count = 5
    high_low_list = []
    volume_list = []

    df = pyupbit.get_ohlcv(coin, interval="minute1", count=data_count + 1)

    for i in range(data_count):
        high_low_list.append(df.iloc[i]['high'])
        high_low_list.append(df.iloc[i]['low'])
        volume_list.append(df.iloc[i]['volume'])

    return max(high_low_list), min(high_low_list), sum(volume_list) / data_count


# 거래량은 직전 5분동안의 평균의 7배 이상
# 강한 양봉 / 전 고가보다 일정 %이상 상승해야

def buy_decision(coin, high, avg_volume):
    df = pyupbit.get_ohlcv(coin, interval="minute1", count=1)
    current_volume = df.iloc[0]['volume']

    if current_volume > avg_volume * 7:
        current_price = pyupbit.get_current_price(coin)
        print(str(time.strftime('%m-%d %H:%M:%S')), "///   coin name: ", coin, "   ///   current price: ", current_price)



coin_array = transaction_top()
print(coin_array, "\n", "\n")

while True:
    for coin in coin_array:
        high, low, avg_volume = price_range(coin)
        buy_decision(coin, high, avg_volume)
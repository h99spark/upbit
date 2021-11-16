import pyupbit
import time
import requests
from bs4 import BeautifulSoup
import sys
import multiprocessing
import os

access = "xHN0w0UsB8a09wdYZXrurJ5xMXxsdv1i8bW6uOh9"
secret = "Ig674PmtvONPAqy7GLxVWVUe9Um4OR7rxmTNgVg9"

upbit = pyupbit.Upbit(access, secret)


def initialize():
    print("%s: Initialize start" % str(time.strftime('%m-%d %H:%M:%S')))
    my_balance = upbit.get_balances()

    if len(my_balance) >= 2:
        for i in range(1, len(my_balance)):
            coin = 'KRW-' + my_balance[i]['currency']
            have = my_balance[i]['balance']
            print("%s Initialize: sell coin %s" % (str(time.strftime('%m-%d %H:%M:%S')), coin))
            upbit.sell_market_order(coin, have)

    print("%s Initialize complete" % str(time.strftime('%m-%d %H:%M:%S')))


def get_coin_list():
    coin_array = []
    url = "https://www.coingecko.com/ko/거래소/upbit"
    coin_num = 15

    response = requests.get(url)
    if response.ok:
        print("%s: Access coingecko successfully" % str(time.strftime('%m-%d %H:%M:%S')))
    else:
        print("%s: Cannot access coingecko" % str(time.strftime('%m-%d %H:%M:%S')))
        sys.exit()

    bs = BeautifulSoup(response.text, 'html.parser')
    selector = "tbody > tr > td > a"
    columns = bs.select(selector)

    ticker_in_krw = [x.text.strip() for x in columns if x.text.strip()[-3:] == "KRW"]
    for ticker in ticker_in_krw:
        coin = "KRW-" + ticker.split('/')[0]
        coin_array.append(coin)

    return coin_array[:coin_num]


def get_tick(coin_array):
    coin_dict = dict()

    for coin in coin_array:
        price = pyupbit.get_current_price(coin)

        if price >= 2000000:
            coin_dict[coin] = 1000.0
        elif price >= 1000000:
            coin_dict[coin] = 500.0
        elif price >= 500000:
            coin_dict[coin] = 100.0
        elif price >= 100000:
            coin_dict[coin] = 50.0
        elif price >= 10000:
            coin_dict[coin] = 10.0
        elif price >= 1000:
            coin_dict[coin] = 5.0
        elif price >= 100:
            coin_dict[coin] = 1.0
        elif price >= 10:
            coin_dict[coin] = 0.1
        elif price >= 0:
            coin_dict[coin] = 0.01

    return coin_dict


def buy_decision(coin_array):

    while True:
        for coin in coin_array:
            volume_list = []

            df = pyupbit.get_ohlcv(coin, interval='minute1', count =3)
            current_price = pyupbit.get_current_price(coin)

            if df.iloc[0]['volume'] < df.iloc[1]['volume'] < df.iloc[2]['volume']:
                volume_increase = True
            else:
                volume_increase = True

            if df.iloc[0]['open'] < df.iloc['close'] and df.iloc[1]['open'] < df.iloc[1]['close'] and df.iloc[2]['open'] < df.iloc[2]['close']:
                continuous_red_candle = True
            else:
                continuous_red_candle = False

            if df.iloc[0]['open'] <= df.iloc[1]['open'] <= df.iloc[2]['open'] \
                and df.iloc[0]['open'] * 1.005 < df.iloc[2]['open'] \
                and df.iloc[0]['open'] + 3 * coin_tick_dict < df.iloc[2]['open']:
                price_increase = True
            else:
                price_increase = False

            if volume_increase and continuous_red_candle and price_increase:
                return coin

def buy_coin(coin):
    print("%s : Buy %s" % (str(time.strftime('%m-%d %H:%M:%S')), coin))
    upbit.buy_market_order(coin, 10000)
    pool.terminate()


def sell_decision(coin):
    sell_good_ratio = 1.02
    sell_bad_ratio = 0.98
    stop_loss_time = 600
    oversee_time = 300
    coin = ''

    while True:
        my_balance = upbit.get_balances()
        if(len(my_balance)) >= 2:
            buy_price = float(my_balance[1]['avg_buy_price'])
            coin = 'KRW-' + my_balance[1]['currency']
            print("%s: Sell decision start: %s" % (str(time.strftime('%m-%d %H:%M:%S')), coin))
            break
        else:
            time.sleep(0.1)

    start_time = time.time()
    while True:
        current_price = pyupbit.get_current_price(coin)

        if 0 <= time.time() - start_time <= oversee_time:
            if current_price >= buy_price * sell_good_ratio:
                break
            elif current_price <= buy_price * sell_bad_ratio:
                break
            else:
                continue
        elif oversee_time < time.time() - start_time <= stop_loss_time:
            if current_price >= buy_price * 1.005:
                break
            elif current_price <= buy_price * sell_bad_ratio:
                break
            else:
                continue
        elif stop_loss_time < time.time() - start_time:
            break
        time.sleep(0.1)

    upbit.sell_market_order(coin, my_balance[1]['balance'])
    print("%s: Sell process complete: %s(%s)" % (str(time.strftime('%m-%d %H:%M:%S')), coin))


if __name__ == '__main__':

    print("\nProgram start: %s" % str(time.strftime('%m-%d %H:%M:%S')))

    initialize()
    coin_array = get_coin_list()

    coin_tick_dict = get_tick(coin_array)

    cpu_num = 5
    pool = multiprocessing.Pool(cpu_num)
    for i in range(cpu_num):
        pool.apply_asnyc(buy_decision, args=(coin_array[i * 3: (i + 1) * 3],), callback = buy_decision)
    pool.close()
    pool.join()

    sell_decision()

    print("Program end: %s" % str(time.strftime('%m-%d %H:%M:%S')))
    time.sleep(180)
    #os.system("nohup python3 -u upbit.py >> output.txt &")
    os.system("python3 /home/ubuntu/upbit/upbit.py")
    sys.exit()

import asyncio
import websockets
import json
import pandas as pd
import pyupbit
import time
import datetime
import schedule
import requests
import numpy as np
import slack_sdk
import os
import xlrd
from openpyxl import Workbook
from slack_sdk import WebClient
from pandas import DataFrame
from pandas import Series

# 로그인
access = "Your Upbit access key"
secret = "Your Upbit secret key"
myToken = "xoxb-Slack Token"

# 변수 설정
maxx = 68           #RSI 최대
minn = 28           #RSI 최소
sellrate1 = 1.2     #목표수익률     
sellrate2 = 1.02     #목표수익률
interval = "days"    # interval
ticker = pyupbit.get_tickers(fiat="KRW")    #ticker
resettime = 300     #재시작 시간(초)

top1coin = ['KRW-BTC']
top1name = ['BTC']
buy_average1 = 100000000

#--------------------------엑셀 저장------------------------

base_dir = "C:/cryptoauto"
file_nm1 = "topcoin1.xlsx"
xlxs_dir1 = os.path.join(base_dir, file_nm1)

file_nm2 = "topcoin2.xlsx"
xlxs_dir2 = os.path.join(base_dir, file_nm2)

file_nm3 = "topcoin3.xlsx"
xlxs_dir3 = os.path.join(base_dir, file_nm3)

file_nm4 = "topcoin4.xlsx"
xlxs_dir4 = os.path.join(base_dir, file_nm4)



#--------------------------함수 설정-------------------------

# 시작 시간 조회
def get_start_time(tickers, interval):             
    df = pyupbit.get_ohlcv(tickers, interval="days", count=1)
    start_time = df.index[0]
    return start_time

# 현재 가격 가져오기
def get_current_price(tickers):                      
    return pyupbit.get_orderbook(tickers)[0]["orderbook_units"][0]["ask_price"]

# 잔고 조회
def get_balance(currency):                          
    balances = upbit.get_balances()
    for b in balances:
        if b['currency'] == currency:
            if b['balance'] is not None:
                return float(b['balance'])
            else:
                return 0

# 매수평균가
def get_buy_average(currency):                      
    balances = upbit.get_balances()
    for b in balances:
        if b['currency'] == currency:
            if b['avg_buy_price'] is not None:
                return float(b['avg_buy_price'])
            else:
                return 0


# 최근 거래 체결 날짜 가져오기
def get_trade_time(ticker):                         
    df = pd.DataFrame(upbit.get_order(ticker, state="done"))
    trade_done = df.iloc[0]["created_at"]
    trade_done_time = datetime.datetime.strptime(trade_done[:-6], "%Y-%m-%dT%H:%M:%S")
    return trade_done_time

# RSI 구하기
def rsi(ohlc: pd.DataFrame, period: int = 14):              
    delta = ohlc["close"].diff() 
    ups, downs = delta.copy(), delta.copy() 
    ups[ups < 0] = 0 
    downs[downs > 0] = 0 
    
    AU = ups.ewm(com = period-1, min_periods = period).mean() 
    AD = downs.abs().ewm(com = period-1, min_periods = period).mean() 
    RS = AU/AD 
    
    return pd.Series(100 - (100/(1 + RS)), name = "RSI") 

# 20일 이동 평균선 조회
def get_ma20(ticker):                           
    df = pyupbit.get_ohlcv(ticker, interval='days', count=20)
    ma20 = df['close'].rolling(window=20, min_periods=1).mean().iloc[-1]
    print(ma20)
    return ma20

# 20시간 이동 평균선 조회
def get_ma20b(ticker):                           
    df = pyupbit.get_ohlcv(ticker, interval='minutes60', count=20)
    ma20b = df['close'].rolling(window=20, min_periods=1).mean().iloc[-1]
    print(ma20b)
    return ma20b

# 10시간 이동 평균선 조회
def get_ma20c(ticker):                           
    df = pyupbit.get_ohlcv(ticker, interval='minutes30', count=20)
    ma20b = df['close'].rolling(window=20, min_periods=1).mean().iloc[-1]
    print(ma20b)
    return ma20b

# 시작 시간 조회
def get_start_time(ticker):
    df = pyupbit.get_ohlcv(ticker, interval="days", count=1)
    start_time = df.index[0]
    return start_time

# 슬랙 메시지 전송
def post_message(token, channel, text):         
    requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )

# 업비트 실시간 통신 : 코인명 / 현재가격 / 등락률 / 거래대금 > 엑셀 저장
async def upbit_websocket_today():
    wb = await websockets.connect("wss://api.upbit.com/websocket/v1", ping_interval=None)
    coinlist = pyupbit.get_tickers(fiat="KRW")  # 코인종류 목록

    namelist = []
    cplist = []
    atplist = []
    scrlist = []
    currencylist =[]

    for coin in coinlist :
        await wb.send(json.dumps([{"ticket":"test"},{"type":"ticker","codes":[coin]},{"format":"SIMPLE"}]))
        if wb.open:
            result = await wb.recv()
            result = json.loads(result)
            name = result.get('cd')
            current_pirce = result.get('tp')
            atp = result.get('atp24h')
            scra = result.get('scr')
            current = name.split('-')[1]
            print(name, "/", current_pirce, '/', scra, "/", atp, "/")

            namelist.append(name)
            cplist.append(current_pirce)
            scrlist.append(scra*100)
            atplist.append(atp)
            currencylist.append(current)

            topcoin = {'코인코드' : namelist,
                        '현재가' : cplist,
                        '등락률' : scrlist,
                       '거래대금' : atplist,
                       '코인이름' : currencylist}

        else :
            loop.close()

    df = pd.DataFrame(topcoin)
    top1 = df.sort_values('등락률', ascending=False)
    top1.to_excel(xlxs_dir1,
                sheet_name='Sheet1',
                na_rep ='NaN',
                float_format = "%.2f",
                header = True,
                index = True,
                index_label="id",
                startrow=0,
                startcol=0,
                )

async def upbit_websocket_always():
    wb = await websockets.connect("wss://api.upbit.com/websocket/v1", ping_interval=None)
    coinlist = pyupbit.get_tickers(fiat="KRW")  # 코인종류 목록

    namelist = []
    cplist = []
    atplist = []
    scrlist = []
    currencylist =[]

    for coin in coinlist :
        await wb.send(json.dumps([{"ticket":"test"},{"type":"ticker","codes":[coin]},{"format":"SIMPLE"}]))
        if wb.open:
            result = await wb.recv()
            result = json.loads(result)
            name = result.get('cd')
            current_pirce = result.get('tp')
            atp = result.get('atp24h')
            scra = result.get('scr')
            current = name.split('-')[1]
            print(name, "/", current_pirce, '/', scra, "/", atp, "/")

            namelist.append(name)
            cplist.append(current_pirce)
            scrlist.append(scra*100)
            atplist.append(atp)
            currencylist.append(current)

            topcoin = {'코인코드' : namelist,
                        '현재가' : cplist,
                        '등락률' : scrlist,
                       '거래대금' : atplist,
                       '코인이름' : currencylist}

        else :
            loop.close()

    df = pd.DataFrame(topcoin)

    top2 = df.sort_values('등락률', ascending=False)
    top2.to_excel(xlxs_dir2,
                sheet_name='Sheet1',
                na_rep ='NaN',
                float_format = "%.2f",
                header = True,
                index = True,
                index_label="id",
                startrow=0,
                startcol=0,
                )
#-------------------------------------------------------------




#--------------------------AutoTrade--------------------------
while True:

    # 시간 설정
    now = datetime.datetime.now()
    start_time = get_start_time("KRW-BTC")
    end_time = start_time + datetime.timedelta(days=1)
    schedule.run_pending()

    # 필수 변수 초기화
    coinname = []
    currentprice = []
    targetprice = []
    percentige = []
    currencyname = []
    rsilist = []
    ma20list = []
    currencylist =[]
    rsi1coin = []
    coin2list = []
   
# [1. 일일 단타]-------------------------------------------------------------------------#
# 목표 : 20%
#  1) 정보 수집 / 09:00:10
#  2) 조사 시작 / 09:00:20
#  3) 매매 시작 / 09:00:30
#  4) 매도 <- 해당 코드는 [3. 상시 단타]에 포함
#   ① 익일 08:59, 시장가 매도
#   ② 20% 달성 시, 시장가 익절
#   ③ -18% 도달 시, 지정가 손절
# -------------------------------------------------------------------------------------#
    if (start_time + datetime.timedelta(seconds=10) < now < start_time + datetime.timedelta(seconds=30)) :

        # 업비트 현황 조사
        loop = asyncio.get_event_loop()
        loop.run_until_complete(upbit_websocket_today())

        # 거래량 > 등락률 순으로 상위 1개 코인 선택
        top1 = pd.read_excel('topcoin1.xlsx')               
        buyone = ((top1['거래대금'] > 150000) & (top1['현재가'] < 20000 ))
        buythis = top1[buyone]
        top1coin = buythis['코인코드'].head(1).values
        top1name = buythis['코인이름'].head(1).values
        print(top1coin, top1name)

        # 로그인
        upbit = pyupbit.Upbit(access, secret)
        #시작 메세지 슬랙 전송
        post_message(myToken,"#hjs-autoupbit", "지성! 일일단타 시작해볼게요!")
        time.sleep(2)
        total = get_balance('KRW')


        #A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
        if len(top1coin) == 0 :
            post_message(myToken,"#hjs-autoupbit", "지성! 지금 없어요...5분 후에 또 볼게요!")
            print("쉴게! 303")
            time.sleep(resettime)

        #B. 매매 시작
        if len(top1coin) != 0 :
            # 순위 조사
            current_price = get_current_price(top1coin[0])
            currencylist = top1coin[0].split('-')[1]
            time.sleep(1) 

            ma1 = get_ma20(top1coin[0])
            ma = current_price - ma1
            time.sleep(2)

            if (ma > 0):        #해당 코인이 볼린더 상단이면 자동 매매 시작
                #자동 매매 시작
                while True :
                    total = get_balance('KRW')
                    current_price = get_current_price(top1coin[0])
                    coin = get_balance(top1name[0])
                    ma1 = get_ma20(top1coin[0])
                    time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                    print(top1coin[0], current_price)

                    #매매 알고리즘
                    if ((total > 5000) and (ma > 0)):                       #해당 코인이 볼린더 상단이면 보유금의 40% 구매
                        upbit.buy_market_order(top1coin[0], total*0.4000)
                        buy_average1 = current_price
                        post_message(myToken,"#hjs-autoupbit", "지성! 일일 단타로" + top1name + "샀어요!")

                        time.sleep(30)
                        break

            elif (ma < 0) :             #해당 코인이 볼린더 상단이면 코인 재조사 시작
                post_message(myToken,"#hjs-autoupbit", "○지성! 지금 없어요...5분 후에 다시 볼게요!")
                print("쉴게! 338")
                time.sleep(resettime)


# [2. 9시 단타]-------------------------------------------------------------------------#
# 목표 : 익절
#  1) 정보 수집 / 09:01:30
#  2) 조사 시작 / 09:02:45
#  3) 매매 시작 / 09:03:00
#  4) 매도
#   ① 09:10, 시장가 매도
# -------------------------------------------------------------------------------------#
    elif start_time + datetime.timedelta(seconds=40) < now < start_time + datetime.timedelta(seconds=70):
        # 업비트 현황 조사
        loop = asyncio.get_event_loop()
        loop.run_until_complete(upbit_websocket_always())

        # 등락률이 높은 코인 중 1분봉(Interval) RSI가 68 이상이고, 현재가가 20000원 이하인 코인 1개 선택
        top2 = pd.read_excel('topcoin2.xlsx')
        coin2list = top2['코인코드'].values  # 코인종류 목록

        for ticker in coin2list:
            currencylist = ticker.split('-')[1]
            current_price = get_current_price(ticker)
            data = pyupbit.get_ohlcv(ticker, interval='minutes1') 
            now_rsi = rsi(data, 14).iloc[-1]
            time.sleep(0.5)
            print(currencylist)

            coinname.append(ticker)
            currencyname.append(currencylist)
            currentprice.append(current_price)
            rsilist.append(now_rsi)

            topcoin = {'코인코드': coinname,
                        'RSI' : rsilist,
                        '현재가' : currentprice,
                        '코인이름' : currencyname}

        top2 = pd.DataFrame(topcoin)
        top2.to_excel(xlxs_dir2,
                    sheet_name='Sheet1',
                    na_rep ='NaN',
                    float_format = "%.2f",
                    header = True,
                    index = True,
                    index_label="id",
                    startrow=0,
                    startcol=0,
                    )

        #RSI 범위 내 필터링하기
        top3 = pd.read_excel('topcoin2.xlsx')               
        buyone = ((top3['RSI'] > maxx) & (top3['현재가'] < 20000 ))
        buythis = top3[buyone]
        top4coin = buythis['코인코드'].head(2).values
        top4name = buythis['코인이름'].head(2).values

        # 로그인
        upbit = pyupbit.Upbit(access, secret)
        time.sleep(2)
        total = get_balance('KRW')

        #A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
        if len(top4coin) == 0 :
            post_message(myToken,"#hjs-autoupbit", "△지성! 지금 없어요...5분 후에 또 볼게요!")
            print("쉴게! 409")
            time.sleep(resettime)

        #B. 매매 시작
        if len(top4coin) != 0 :
            # 순위 조사
            top3coin = list(set(top4coin) - set(top1coin))
            top2coin = top3coin[0]
            top2name = top3coin[0].split('-')[1]
            print(top4coin, top4name)
            print(top2coin, top2name)

            current_price2 = get_current_price(top2coin)

            data = pyupbit.get_ohlcv(top2coin, interval='minutes10') 
            now_rsi = rsi(data, 14).iloc[-1]

            ma2 = get_ma20b(top2coin)
            ma = current_price2 - ma2
            time.sleep(2)

            if (ma > 0):        #해당 코인이 볼린더 상단이면 자동 매매 시작
                #자동 매매 시작
                while True :
                    total = get_balance('KRW')
                    current_price1 = get_current_price(top1coin[0])                    
                    current_price2 = get_current_price(top2coin)
                    coin1 = get_balance(top1name[0])
                    coin2 = get_balance(top2name)
                    now = datetime.datetime.now()
                    data = pyupbit.get_ohlcv(top2coin, interval='minutes1') 
                    now_rsi = rsi(data, 14).iloc[-1]
                    schedule.run_pending()
                    time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                    print(current_price2, now_rsi)

                    #매매 알고리즘
                    if ((total > 5000) and (ma > 0)):                       #상시 코인이 볼린더 상단이면 구매
                        upbit.buy_market_order(top2coin, total*0.9995)
                        buy_average2 = current_price2
                        post_message(myToken,"#hjs-autoupbit", "지성! 9시 단타로" + top2name + "샀어요!")
                        time.sleep(30)

                    elif start_time + datetime.timedelta(minutes=10) < now < start_time + datetime.timedelta(minutes=11):         #9:10분 강제 매도    
                        upbit.sell_market_order(top2coin, coin2)       
                        post_message(myToken,"#hjs-autoupbit", "지성! 9시 단타 끝!")
                        time.sleep(30)
                        break

            elif (ma < 0) :             #해당 코인이 볼린더 상단이면 코인 재조사 시작
                post_message(myToken,"#hjs-autoupbit", "△지성! 지금 없어요...5분 후에 다시 볼게요!")
                print("쉴게! 454") 
                time.sleep(resettime)     

# [3. 상시 단타]-------------------------------------------------------------------------#
# 목표 : 2%
#  1) 정보 수집 / 09:10:30
#  2) 조사 시작 / 09:11:45
#  3) 매매 시작 / 09:12:00
#  4) 매도
#   ① 2% 달성 시, 지정가 익절
#   ② -18% 도달 시, 지정가 손절
#   (③ 익일 08:59, 강제 시장가 매도)
# -------------------------------------------------------------------------------------#
    else :
        # 업비트 현황 조사
        loop = asyncio.get_event_loop()
        loop.run_until_complete(upbit_websocket_always())

        # 등락률이 높은 코인 중 60분봉(Interval) RSI가 28 이하이고, 현재가가 20000원 이하인 코인 1개 선택
        top3 = pd.read_excel('topcoin2.xlsx')
        coin2list = top3['코인코드'].values  # 코인종류 목록

        for ticker in coin2list:
            currencylist = ticker.split('-')[1]
            current_price = get_current_price(ticker)
            data = pyupbit.get_ohlcv(ticker, interval='minutes60') 
            now_rsi = rsi(data, 14).iloc[-1]
            time.sleep(0.5)
            print(currencylist)

            coinname.append(ticker)
            currencyname.append(currencylist)
            currentprice.append(current_price)
            rsilist.append(now_rsi)

            topcoin = {'코인코드': coinname,
                        'RSI' : rsilist,
                        '현재가' : currentprice,
                        '코인이름' : currencyname}

        top3 = pd.DataFrame(topcoin)
        top3.to_excel(xlxs_dir2,
                    sheet_name='Sheet1',
                    na_rep ='NaN',
                    float_format = "%.2f",
                    header = True,
                    index = True,
                    index_label="id",
                    startrow=0,
                    startcol=0,
                    )

        #RSI 범위 내 필터링하기
        top2 = pd.read_excel('topcoin2.xlsx')               
        buyone = ((top2['RSI'] < minn) & (top2['현재가'] < 20000 ))
        buythis = top2[buyone]
        top4coin = buythis['코인코드'].head(2).values
        top4name = buythis['코인이름'].head(2).values

        # 로그인
        upbit = pyupbit.Upbit(access, secret)
        time.sleep(2)
        total = get_balance('KRW')

        #A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
        if len(top4coin) == 0 :
            post_message(myToken,"#hjs-autoupbit", "△지성! 지금 없어요...5분 후에 또 볼게요!")
            print("쉴게! 521")
            time.sleep(resettime)


        #B. 매매 시작
        if len(top4coin) != 0 :
            # 순위 조사
            top3coin = list(set(top4coin) - set(top1coin))
            top2coin = top3coin[0]
            top2name = top3coin[0].split('-')[1]
            print(top4coin, top4name)
            print(top2coin, top2name)

            current_price2 = get_current_price(top2coin)

            data = pyupbit.get_ohlcv(top2coin, interval='minutes60') 
            now_rsi = rsi(data, 14).iloc[-1]

            ma2 = get_ma20b(top2coin)
            ma = current_price2 - ma2
            time.sleep(2)

            if (len(top4coin) != 0):        #해당 코인이 볼린더 상단이면 자동 매매 시작
                #자동 매매 시작
                while True :
                    total = get_balance('KRW')
                    current_price1 = get_current_price(top1coin[0])                    
                    current_price2 = get_current_price(top2coin)
                    coin1 = get_balance(top1name[0])
                    coin2 = get_balance(top2name)
                    now = datetime.datetime.now()
                    data = pyupbit.get_ohlcv(top2coin, interval='minutes60') 
                    now_rsi = rsi(data, 14).iloc[-1]
                    ma1 = get_ma20b(top2coin)
                    schedule.run_pending()
                    time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                    print(current_price2, now_rsi)

                    #매매 알고리즘
                    if (total > 5000):                       #상시 코인이 볼린더 상단이면 구매
                        upbit.buy_market_order(top2coin, total*0.9995)
                        buy_average2 = current_price2
                        post_message(myToken,"#hjs-autoupbit", "지성!" + top2name + "샀어요!")
                        time.sleep(30)

                    if (current_price2 > (buy_average2 * sellrate2)):                                       #상시 코인가격이 2% 상승하면 지정가 익절
                        upbit.sell_limit_order(top2coin, current_price2, coin2)       
                        post_message(myToken,"#hjs-autoupbit", "지성! 오케이! 2% 하나 더 찾아볼게요!")
                        time.sleep(10)
                        break

                    if (current_price1 > (buy_average1 * sellrate1)):                                       #일일 코인가격이 20% 상승하면 시장가 익절
                        upbit.sell_limit_order(top1coin[0], current_price1, coin1)       
                        post_message(myToken,"#hjs-autoupbit", "지성! 오케이! 20% 완료!")
                        time.sleep(10)
                        break

                    elif end_time - datetime.timedelta(minutes=1) < now < end_time:       #08:59시에 전량 시장가 매도
                        upbit.sell_market_order(top1coin[0], coin1)       
                        upbit.sell_market_order(top2coin, coin2)       
                        post_message(myToken,"#hjs-autoupbit", "지성! 좋은 아침, 오늘꺼 준비할게!")
                        time.sleep(71)
                        break

            elif (len(top4coin) == 0) :             #해당 코인이 볼린더 상단이면 코인 재조사 시작
                post_message(myToken,"#hjs-autoupbit", "△지성! 지금 없어요...5분 후에 다시 볼게요!")
                print("쉴게! 587")
                time.sleep(resettime)
#-------------------------------------------------------------

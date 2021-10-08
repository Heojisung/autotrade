# --------------- [간단 설명] -------------------------------------------------
# - 업비트 Websocket을 통해 전체 코인 정보 수신
# - 등락률을 내림차순 정렬
# - 시간대별로 "거래대금" or "RSI" or "볼린저밴드"를 기준으로 매매
#  1) [09:00, 13:00 ,17:00]  15분간 2회식 초단타 매매
#  2) [17:15 ~ 익일 08:59] 30분봉 RSI 기반으로 매매
#  3) [나머지 시간] 60분봉 RSI 기반으로 매매
# -----------------------------------------------------------------------------

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
maxx = 65           #RSI 최대
minn = 29           #RSI 최소
sellrate = 1.03     #목표수익률     
interval = "days"    # interval
ticker = pyupbit.get_tickers(fiat="KRW")    #ticker
resettime = 300     #재시작 시간(초)

#--------------------------엑셀 저장-------------------------

base_dir = "C:/cryptoauto"
file_nm = "topcoin.xlsx"
xlxs_dir = os.path.join(base_dir, file_nm)

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
async def upbit_websocket():
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
    top2.to_excel(xlxs_dir,
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
    coin3list = []


#-------------------------- 1) 9시 1분/6분, 13시 1분/6분, 17시 1분/6분에 초단타 매매 시작-----------------------------------------------#
    if (start_time + datetime.timedelta(minutes=1) < now < start_time + datetime.timedelta(minutes=9)) or (start_time + datetime.timedelta(minutes=241) < now < start_time + datetime.timedelta(minutes=249)) or (start_time + datetime.timedelta(minutes=481) < now < start_time + datetime.timedelta(minutes=489)) :

        # 업비트 현황 조사
        loop = asyncio.get_event_loop()
        loop.run_until_complete(upbit_websocket())

        # 거래량 > 등락률 순으로 상위 1개 코인 선택
        top1 = pd.read_excel('topcoin.xlsx')               
        buyone = ((top1['거래대금'] > 150000) & (top1['현재가'] < 20000 ))
        buythis = top1[buyone]
        top1coin = buythis['코인코드'].head(1).values
        top1name = buythis['코인이름'].head(1).values
        print(top1coin, top1name)

        # 로그인
        upbit = pyupbit.Upbit(access, secret)
        #시작 메세지 슬랙 전송
        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "사볼게요!")
        time.sleep(2)
        total = get_balance('KRW')


        #A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
        if len(top1coin) == 0 :
            post_message(myToken,"#hjs-autoupbit", "○__! 지금 없어요...5분 후에 또 볼게요!")
            time.sleep(resettime)

        #B. 매매 시작
        if len(top1coin) != 0 :
            # 순위 조사
            current_price = get_current_price(top1coin[0])
            currencylist = top1coin[0].split('-')[1]
            time.sleep(1) 

            data = pyupbit.get_ohlcv(top1coin[0], interval) 
            now_rsi = rsi(data, 14).iloc[-1]

            ma1 = get_ma20(top1coin[0])
            ma = current_price - ma1
            time.sleep(2)

            if (ma > 0):        #해당 코인이 볼린더 상단이면 자동 매매 시작
                #자동 매매 시작
                while True :
                    total = get_balance('KRW')
                    current_price = get_current_price(top1coin[0])
                    coin = get_balance(top1name[0])
                    now = datetime.datetime.now()
                    data = pyupbit.get_ohlcv(top1coin[0], interval='day') 
                    now_rsi = rsi(data, 14).iloc[-1]
                    ma1 = get_ma20(top1coin[0])
                    time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                    print(current_price, now_rsi)

                    #매매 알고리즘
                    if ((total > 5000) and (ma > 0)):                       #해당 코인이 볼린더 상단이면 구매
                        upbit.buy_market_order(top1coin[0], total*0.9995)
                        buy_average = current_price
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "샀어요! 시작해볼게요!")
                        time.sleep(30)

                    if start_time + datetime.timedelta(minutes=6) < now < start_time + datetime.timedelta(minutes=7):       # 09:06분 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이거 팔고, 다음꺼 살게!")
                        time.sleep(61)
                        break

                    if start_time + datetime.timedelta(minutes=15) < now < start_time + datetime.timedelta(minutes=16):       # 09:15분 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이거 팔고, 다음꺼 살게!")
                        time.sleep(61)
                        break

                    if start_time + datetime.timedelta(minutes=246) < now < start_time + datetime.timedelta(minutes=247):       # 13:06분 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이거 팔고, 다음꺼 살게!")
                        time.sleep(61)
                        break

                    if start_time + datetime.timedelta(minutes=255) < now < start_time + datetime.timedelta(minutes=256):       # 13:15분 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이거 팔고, 다음꺼 살게!")
                        time.sleep(61)
                        break

                    if start_time + datetime.timedelta(minutes=486) < now < start_time + datetime.timedelta(minutes=487):       # 17:06분 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이거 팔고, 다음꺼 살게!")
                        time.sleep(61)
                        break

                    if start_time + datetime.timedelta(minutes=495) < now < start_time + datetime.timedelta(minutes=496):       # 17:15분 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이거 팔고, 다음꺼 살게!")
                        time.sleep(61)
                        break

                    elif (current_price < ma1):                                       #볼린저밴드 하단 통과 시 손절 매도
                        upbit.sell_market_order(top1coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "○__!" + top1name + "이건 손절할게!")
                        time.sleep(30)
                        break

            elif (ma < 0) :             #해당 코인이 볼린더 상단이면 코인 재조사 시작
                post_message(myToken,"#hjs-autoupbit", "○__! 지금 없어요...5분 후에 다시 볼게요!")
                time.sleep(resettime)


#----------------------------- 2) 17:15 ~ 익일 08:59 -------------------------------------------------------------#
    if start_time + datetime.timedelta(minutes=495) < now < end_time - datetime.timedelta(minutes=2):
        # 업비트 현황 조사
        loop = asyncio.get_event_loop()
        loop.run_until_complete(upbit_websocket())

        # 등락률이 높은 코인 중 60분봉(Interval) RSI가 28.5 이하이고, 현재가가 20000원 이하인 코인 1개 선택
        top3 = pd.read_excel('topcoin.xlsx')
        coin3list = top3['코인코드'].values  # 코인종류 목록

        for ticker in coin3list:
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

        top2 = pd.DataFrame(topcoin)
        top2.to_excel(xlxs_dir,
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
        top4 = pd.read_excel('topcoin.xlsx')               
        buy4one = ((top4['RSI'] < minn) & (top4['현재가'] < 20000 ))
        buy4this = top4[buy4one]
        top4coin = buy4this['코인코드'].head(1).values
        top4name = buy4this['코인이름'].head(1).values
        print(top4coin, top4name)

        # 로그인
        upbit = pyupbit.Upbit(access, secret)
        #시작 메세지 슬랙 전송
        post_message(myToken,"#hjs-autoupbit", "△__!" + top4name + "사볼게요!")
        time.sleep(2)
        total = get_balance('KRW')

        #A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
        if len(top4coin) == 0 :
            post_message(myToken,"#hjs-autoupbit", "△__! 지금 없어요...5분 후에 또 볼게요!")
            time.sleep(resettime)

        #B. 매매 시작
        if len(top4coin) != 0 :
            # 순위 조사
            current_price = get_current_price(top4coin[0])
            currencylist = top4coin[0].split('-')[1]
            time.sleep(1) 

            data = pyupbit.get_ohlcv(top4coin[0], interval='minutes60') 
            now_rsi = rsi(data, 14).iloc[-1]

            ma2 = get_ma20b(top4coin[0])
            ma = current_price - ma2
            time.sleep(2)

            if (ma > 0):        #해당 코인이 볼린더 상단이면 자동 매매 시작
                #자동 매매 시작
                while True :
                    total = get_balance('KRW')
                    current_price = get_current_price(top4coin[0])
                    coin = get_balance(top4name[0])
                    now = datetime.datetime.now()
                    data = pyupbit.get_ohlcv(top4coin[0], interval='minutes60') 
                    now_rsi = rsi(data, 14).iloc[-1]
                    ma1 = get_ma20b(top4coin[0])
                    schedule.run_pending()
                    time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                    print(current_price, now_rsi)

                    #매매 알고리즘
                    if ((total > 5000) and (ma > 0)):                       #해당 코인이 볼린더 상단이면 구매
                        upbit.buy_market_order(top4coin[0], total*0.9995)
                        buy_average = current_price
                        post_message(myToken,"#hjs-autoupbit", "△__!" + top4name + "샀어요! 시작해볼게요!")
                        time.sleep(30)

                    if (current_price > (buy_average * 1.1)):                                       #해당 코인가격이 10% 상승하면 시장가 익절
                        upbit.sell_market_order(top4coin[0], coin)
                        post_message(myToken,"#hjs-autoupbit", "△__! 오케이! 10% 하나 더 찾아볼게요!")
                        time.sleep(30)
                        break

                    if (now_rsi > maxx):                                       #RSI가 65 이상이면 시장가 익절
                        upbit.sell_market_order(top4coin[0], coin)
                        post_message(myToken,"#hjs-autoupbit", "△__! 오케이! 10% 하나 더 찾아볼게요!")
                        time.sleep(30)
                        break

                    elif end_time - datetime.timedelta(minutes=1) < now < end_time:       #08:59시에 전량 매도
                        upbit.sell_market_order(top4coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "△__!" + top4name + "이거 팔고, 다음꺼 볼게!")
                        time.sleep(121)
                        break

            elif (ma < 0) :             #해당 코인이 볼린더 상단이면 코인 재조사 시작
                post_message(myToken,"#hjs-autoupbit", "△__! 지금 없어요...5분 후에 다시 볼게요!")
                time.sleep(resettime)     


#----------------------------- 3) 나머지 시간 자동매매 -------------------------------------------------------------#
    else:
        # 업비트 현황 조사
        loop = asyncio.get_event_loop()
        loop.run_until_complete(upbit_websocket())

        # 등락률이 높은 코인 중 30분봉(Interval) RSI가 67 이상이고, 현재가가 20000원 이하인 코인 1개 선택
        top3 = pd.read_excel('topcoin.xlsx')
        coin3list = top3['코인코드'].values  # 코인종류 목록

        for ticker in coin3list:
            currencylist = ticker.split('-')[1]
            current_price = get_current_price(ticker)
            data = pyupbit.get_ohlcv(ticker, interval='minutes30') 
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
        top2.to_excel(xlxs_dir,
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
        top4 = pd.read_excel('topcoin.xlsx')               
        buy4one = ((top4['RSI'] > maxx) & (top4['현재가'] < 20000 ))
        buy4this = top4[buy4one]
        top4coin = buy4this['코인코드'].head(1).values
        top4name = buy4this['코인이름'].head(1).values
        print(top4coin, top4name)

        # 로그인
        upbit = pyupbit.Upbit(access, secret)
        #시작 메세지 슬랙 전송
        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "사볼게요!")
        time.sleep(2)
        total = get_balance('KRW')

        #A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
        if len(top4coin) == 0 :
            post_message(myToken,"#hjs-autoupbit", "□__! 지금 없어요...5분 후에 또 볼게요!")
            time.sleep(resettime)

        #B. 매매 시작
        if len(top4coin) != 0 :
            # 순위 조사
            current_price = get_current_price(top4coin[0])
            currencylist = top4coin[0].split('-')[1]
            time.sleep(1) 

            data = pyupbit.get_ohlcv(top4coin[0], interval='minutes30') 
            now_rsi = rsi(data, 14).iloc[-1]

            ma2 = get_ma20c(top4coin[0])
            ma = current_price - ma2
            time.sleep(2)

            if (ma > 0):        #해당 코인이 볼린더 상단이면 자동 매매 시작
                #자동 매매 시작
                while True :
                    total = get_balance('KRW')
                    current_price = get_current_price(top4coin[0])
                    coin = get_balance(top4name[0])
                    now = datetime.datetime.now()
                    data = pyupbit.get_ohlcv(top4coin[0], interval='minutes30') 
                    now_rsi = rsi(data, 14).iloc[-1]
                    ma1 = get_ma20c(top4coin[0])
                    schedule.run_pending()
                    time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                    print(current_price, now_rsi)

                    #매매 알고리즘
                    if ((total > 5000) and (ma > 0)):                       #해당 코인이 볼린더 상단이면 구매
                        upbit.buy_market_order(top4coin[0], total*0.9995)
                        buy_average = current_price
                        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "샀어요! 시작해볼게요!")
                        time.sleep(30)

                    if (current_price > (buy_average * sellrate)):                                       #해당 코인가격이 목표가 도달하면 시장가 익절
                        upbit.sell_market_order(top4coin[0], coin)
                        post_message(myToken,"#hjs-autoupbit", "□__! 오케이! 하나 더 찾아볼게요!")
                        time.sleep(30)
                        break

                    if end_time - datetime.timedelta(minutes=1) < now < end_time:       #08:59시에 전량 매도
                        upbit.sell_market_order(top4coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "이거 팔고, 다음꺼 볼게!")
                        time.sleep(121)
                        break

                    if start_time + datetime.timedelta(minutes=239) < now < start_time + datetime.timedelta(minutes=240):       #12:59시에 전량 매도
                        upbit.sell_market_order(top4coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "이거 팔고, 다음꺼 볼게!")
                        time.sleep(121)
                        break

                    if start_time + datetime.timedelta(minutes=479) < now < start_time + datetime.timedelta(minutes=480):       #16:59시에 전량 매도
                        upbit.sell_market_order(top4coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "이거 팔고, 다음꺼 볼게!")
                        time.sleep(121)
                        break                    

                    elif (current_price < ma2):                                       #볼린저밴드 하단 통과 시 손절 매도
                        upbit.sell_market_order(top4coin[0], coin)       
                        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "이건 손절할게!")
                        time.sleep(30)
                        break

                    if (current_price < (buy_average * 0.82)):                                       #해당 코인가격이 18% 빠지면 시장가 손절
                        upbit.sell_market_order(top4coin[0], coin)
                        post_message(myToken,"#hjs-autoupbit", "□__!" + top4name + "손절할게!")
                        time.sleep(30)
                        break

            elif (ma < 0) :             #해당 코인이 볼린더 상단이면 코인 재조사 시작
                post_message(myToken,"#hjs-autoupbit", "□__! 지금 없어요...5분 후에 다시 볼게요!")
                time.sleep(resettime)     

#-------------------------------------------------------------
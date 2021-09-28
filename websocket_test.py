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
from fbprophet import Prophet

# 로그인
access = "Your Upbit access"
secret = "Your Upbit secret"
myToken = "xoxb-Your Slack Token"

# 변수 설정
maxx = 71           #RSI 최대 
minn = 65           #RSI 최소
sellrate = 1.05     #목표수익률
interval = "minute240"    # interval
ticker = pyupbit.get_tickers(fiat="KRW")    #ticker
k = 0.1     # 변동률 k
resettime = 300     #재시작 시간(초)

#--------------------------필수 변수-------------------------

pcplist = []
coinname = []
currentprice = []
targetprice = []
gapprice = []
percentige = []
currencyname = []
rsilist = []
ma20list = []
currencylist =[]

base_dir = "C:/cryptoauto"
file_nm = "topcoin.xlsx"
xlxs_dir = os.path.join(base_dir, file_nm)

#--------------------------함수 설정-------------------------

# 변동성 돌파 전략으로 매수 목표가 정하기
def get_target_price(ticker, k):         
    df = pyupbit.get_ohlcv(ticker, interval="minute15", count=2)     
    target_price = df.iloc[0]['close'] + (df.iloc[0]['high'] - df.iloc[0]['low']) * k       
    return target_price

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

#RSI 구하기
def rsi(ohlc: pd.DataFrame, period: int = 14):              
    delta = ohlc["close"].diff() 
    ups, downs = delta.copy(), delta.copy() 
    ups[ups < 0] = 0 
    downs[downs > 0] = 0 
    
    AU = ups.ewm(com = period-1, min_periods = period).mean() 
    AD = downs.abs().ewm(com = period-1, min_periods = period).mean() 
    RS = AU/AD 
    
    return pd.Series(100 - (100/(1 + RS)), name = "RSI") 

#20일 이동 평균선 조회
def get_ma20(ticker):                           
    df = pyupbit.get_ohlcv(ticker, interval='days', count=20)
    ma20 = df['close'].rolling(window=20, min_periods=1).mean().iloc[-1]
    print(ma20)
    return ma20

#슬랙 메시지 전송
def post_message(token, channel, text):         
    requests.post("https://slack.com/api/chat.postMessage",
        headers={"Authorization": "Bearer "+token},
        data={"channel": channel,"text": text}
    )

# 당일 종가 예측
predicted_close_price = 0
def predict_price(ticker):
    """Prophet으로 당일 종가 가격 예측"""
    global predicted_close_price
    df = pyupbit.get_ohlcv(ticker, interval="minute60")
    df = df.reset_index()
    df['ds'] = df['index']
    df['y'] = df['close']
    data = df[['ds','y']]
    model = Prophet()
    model.fit(data)
    future = model.make_future_dataframe(periods=24, freq='H')
    forecast = model.predict(future)
    closeDf = forecast[forecast['ds'] == forecast.iloc[-1]['ds'].replace(hour=9)]
    if len(closeDf) == 0:
        closeDf = forecast[forecast['ds'] == data.iloc[-1]['ds'].replace(hour=9)]
    closeValue = closeDf['yhat'].values[0]
    predicted_close_price = closeValue

# 업비트 실시간 통신 : 코인명 / 현재가격 / 등락율 / 거래대금 > 엑셀 저장
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
            print(name, "/", current_pirce, '/', scra, "/", atp)

            namelist.append(name)
            cplist.append(current_pirce)
            scrlist.append(scra*100)
            atplist.append(atp)
            currencylist.append(current)

            topcoin = {'코인코드' : namelist,
                        '현재가' : cplist,
                        '등락율' : scrlist,
                       '거래대금' : atplist,
                       '코인이름' : currencylist}

        else :
            loop.close()

    df = pd.DataFrame(topcoin)
    top2 = df.sort_values('거래대금', ascending=False)
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

# 업비트 현황 조사
loop = asyncio.get_event_loop()
loop.run_until_complete(upbit_websocket())

# 거래량 > 등락율 순으로 상위 1개 코인 선택
top1 = pd.read_excel('topcoin.xlsx')               
buyone = (top1['등락율'] > 0)
buythis = top1[buyone]
top1coin = buythis['코인코드'].head(1).values
top1name = buythis['코인이름'].head(1).values
print(top1coin, top1name)

# 로그인
upbit = pyupbit.Upbit(access, secret)
print("오늘은 뭘 사볼까요?")
#시작 메세지 슬랙 전송
post_message(myToken,"#hjs-autoupbit", "지성!" + top1name + "사볼게요!")
time.sleep(2)
total = get_balance("KRW")

#A. 매매할 코인이 없는 경우, 일정 시간 후 준비단계부터 재시작
if len(top1coin) == 0 :
    post_message(myToken,"#hjs-autoupbit", "지성! 지금 없어요...5분 후에 또 볼게요!")
    time.sleep(resettime)

#B. 매매 시작
if len(top1coin) != 0 :
    # 순위 조사
    predict_price(top1coin)
    current_price = get_current_price(top1coin)
    target_price = get_target_price(top1coin, k)
    gap_price = predicted_close_price - current_price
    currencylist = ticker.split('-')[1]
    time.sleep(1) 

    data = pyupbit.get_ohlcv(top1coin, interval) 
    now_rsi = rsi(data, 14).iloc[-1]

    ma1 = get_ma20(top1coin)
    ma = current_price - ma1
    time.sleep(2)

    pcplist.append(predicted_close_price)
    coinname.append(ticker)
    currentprice.append(current_price)
    targetprice.append(target_price)
    gapprice.append(gap_price)
    currencyname.append(currencylist)
    rsilist.append(now_rsi)
    ma20list.append(ma)

    topcoin = {'코인코드': coinname,
                'RSI' : rsilist,
                'ma' : ma20list,
                '목표매수' : targetprice,
                '예상종가': pcplist,
                '차액' : gapprice,
                '현재가격' : currentprice,
                '코인이름' : currencyname}

    df = pd.DataFrame(topcoin)
    top2 = df.sort_values('RSI', ascending=False)
    top2.to_excel(xlxs_dir,
                sheet_name='Sheet2',
                na_rep ='NaN',
                float_format = "%.2f",
                header = True,
                index = True,
                index_label="id",
                startrow=0,
                startcol=0,
                )

    target_price = buythis['목표매수'].head(1).values
    current_price = get_current_price(top1coin[0])
    predicted_close_price = buythis['예상종가'].head(1).values
    coin = get_balance(top1name[0])

    if (current_price < predicted_close_price):  
        #자동 매매 시작
        while True :
            total = get_balance("KRW")
            current_price = get_current_price(top1coin[0])
            buy_average = get_buy_average(top1name[0])
            coin = get_balance(top1name[0])
            time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
            print(current_price,buy_average)

            #매매 알고리즘
            if ((total > 5000) and (current_price < predicted_close_price)):                       #해당 코인의 예상종가가 높으면 매매
                upbit.buy_market_order(top1coin[0], total*0.9995)
                post_message(myToken,"#hjs-autoupbit", "지성!샀어요! 시작해볼게요!")
                time.sleep(30)
                
            elif (current_price > (buy_average * sellrate)):                                       #해당 코인가격이 목표가 도달하면 시장가 매도
                upbit.sell_market_order(top1coin[0], coin)       
                post_message(myToken,"#hjs-autoupbit", "지성!오케이! 하나 더 찾아볼게요!")
                time.sleep(30)
                break
                
    elif (current_price > predicted_close_price) :
        post_message(myToken,"#hjs-autoupbit", "지성! 지금 없어요...10분 후에 다시 볼게요!")
        time.sleep(resettime)
                    
#-------------------------------------------------------------

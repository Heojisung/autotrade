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
from slack_sdk import WebClient
from pandas import DataFrame
from pandas import Series
from fbprophet import Prophet

# 로그인
access = "Your Upbit access key"
secret = "Your Upbit secret key"
myToken = "xoxb-Your slack token"

# 변수 설정
maxx = 71           #RSI 최대 
minn = 66           #RSI 최소
sellrate = 1.05     #목표수익률
interval = "day"    # interval
ticker = pyupbit.get_tickers(fiat="KRW")    #ticker
k = 0.1     # 변동률 k
rate_minus = 0.95   #???

#--------------------------함수 설정--------------------------

# 변동성 돌파 전략으로 매수 목표가 정하기
def get_target_price(ticker, k):   
    df = pyupbit.get_ohlcv(ticker, interval="day", count=2)     
    target_price = df.iloc[0]['close'] + (df.iloc[0]['high'] - df.iloc[0]['low']) * k       
    return target_price
  
# 시작 시간 조회
def get_start_time(tickers, interval):
    df = pyupbit.get_ohlcv(tickers, interval="day", count=1)
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

# 최근 거래체결 날짜 가져오기
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
    df = pyupbit.get_ohlcv(ticker, interval, count=20)
    ma20 = df['close'].rolling(window=20, min_periods=1).mean().iloc[-1]
    print(ma20)
    return ma20

# 슬랙 메시지 전송
def post_message(token, channel, text):         
    response = requests.post("https://slack.com/api/chat.postMessage",
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


# 엑셀만들기
base_dir = "C:/cryptoauto"
file_nm = "topcoin.xlsx"
xlxs_dir = os.path.join(base_dir, file_nm)
#-------------------------------------------------------------


#--------------------------AutoTrade--------------------------
            
# 시작 준비!
while True:

    # 조사 목록
    coinlist = pyupbit.get_tickers(fiat="KRW")  # 코인종류 목록
    # coinlist = ["KRW-UPP","KRW-OMG", "KRW-BTC", "KRW-FCT2"] #테스트 코인
    pcplist = []  # 예상종가 목록
    coinname = [] # "KRW-코인이름" 목록
    currentprice = [] # 현재가격 목록
    targetprice = []  # 목표매수가격 목록
    gapprice = [] # 차액 목록
    percentige = [] # 등락률 목록
    currencyname = [] # "코인이름"목록
    rsilist = []  # RSI 목록
    ma20list = [] #20일선 가격 목록

    for ticker in coinlist:
        predict_price(ticker)
        current_price = get_current_price(ticker)
        target_price = get_target_price(ticker, k)
        gap_price = predicted_close_price - current_price
        percent = gap_price / predicted_close_price * 100
        currencylist = ticker.split('-')[1]
        time.sleep(1) 

        data = pyupbit.get_ohlcv(ticker, interval) 
        now_rsi = rsi(data, 14).iloc[-1]

        ma1 = get_ma20(ticker)
        ma = current_price - ma1
        time.sleep(2)

        pcplist.append(predicted_close_price)
        coinname.append(ticker)
        currentprice.append(current_price)
        targetprice.append(target_price)
        gapprice.append(gap_price)
        percentige.append(percent)
        currencyname.append(currencylist)
        rsilist.append(now_rsi)
        ma20list.append(ma)

        topcoin = {'코인이름': coinname,
                    'RSI' : rsilist,
                    'ma' : ma20list,
                    '비율' : percentige,
                    '목표매수' : targetprice,
                    '예상종가': pcplist,
                    '차액' : gapprice,
                    '현재가격' : currentprice,
                    '이름' : currencyname}

        Top_coin = DataFrame(topcoin)

    df = pd.DataFrame(topcoin)
    top2 = df.sort_values('비율', ascending=False)
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

    #RSI 범위 내 상위 2개 필터링하기
    top1 = pd.read_excel('topcoin.xlsx')                
    buyone = ((top1['RSI'] > minn) & (top1['RSI'] < maxx))
    buythis = top1[buyone]
    top1coin = buythis['코인이름'].head(1).values
    top1name = buythis['이름'].head(1).values
    top2coin = buythis['코인이름'].head(2).values
    top2name = buythis['이름'].head(2).values
    print(top1coin, top1name, "/", top2coin, top2name)

    # 로그인
    upbit = pyupbit.Upbit(access, secret)
    print("오늘은 뭘 사볼까요?")

    #시작 메세지 슬랙 전송
    post_message(myToken,"#hjs-autoupbit", top1name + "과" + top2name + "사볼게요!")
    total = get_balance("KRW")

    #A. 매매할 코인이 2개가 안될 경우, 30분 후 준비부터 재시작
    if (len(top1coin) == 0 or len(top2coin) == 0) :
        post_message(myToken,"#hjs-autoupbit", "지금 없어요...1시간 후에 또 볼게요!")
        time.sleep(1800)

    #B. 매매 코인이 2개가 있는 경우, 알고리즘에 따라 매매 시작
    if (len(top1coin) != 0 and len(top2coin != 0)) :
        target_price1 = buythis['목표매수'].head(1).values
        current_price1 = get_current_price(top1coin[0])
        predicted_close_price1 = buythis['예상종가'].head(1).values
        coin1 = get_balance(top1name[0])
        target_price2 = buythis['목표매수'].head(2).values
        current_price2 = get_current_price(top2coin[0])
        predicted_close_price2 = buythis['예상종가'].head(2).values
        coin2 = get_balance(top2name[0])

        if ((current_price1 < predicted_close_price1) and (current_price2 < predicted_close_price2)):  
            #자동 매매 시작
            i = 0
            while i < 2 :
                total = get_balance("KRW")
                current_price1 = get_current_price(top1coin[0])
                buy_average1 = get_buy_average(top1name[0])   
                current_price2 = get_current_price(top2coin[0])
                buy_average2 = get_buy_average(top2name[0])
                coin1 = get_balance(top1name[0])
                coin2 = get_balance(top2name[0])   
                time.sleep(2)                                                                          #2초에 한번씩 현재가격 갱신
                print(current_price1, buy_average1)
                print(current_price2, buy_average2)

                #매매 알고리즘
                if ((total > 5000) and (current_price1 < predicted_close_price1) and (current_price2 < predicted_close_price2)):                       #해당 코인의 예상종가가 높으면 매매
                    upbit.buy_market_order(top1coin[0], total*0.6)
                    upbit.buy_market_order(top2coin[0], total*0.3995)
                    post_message(myToken,"#hjs-autoupbit", "모두 샀어요! 시작해볼게요!")
                    time.sleep(30)

                if (current_price1 > (buy_average1 * sellrate) and (coin1 != 0)):                                       #해당 코인가격이 목표가 도달하면 시장가 매도
                    upbit.sell_market_order(top1coin[0], coin1)       
                    post_message(myToken,"#hjs-autoupbit", "오케이! 하나 완료!")
                    i += 1
                    time.sleep(30)

                elif (current_price2 > (buy_average2 * sellrate) and (coin2 != 0)):                                       #해당 코인가격이 목표가 도달하면 시장가 매도
                    upbit.sell_market_order(top2coin[0], coin1)       
                    post_message(myToken,"#hjs-autoupbit", "오케이! 하나 완료!")
                    i += 1
                    time.sleep(30)

        #자동 매매 종료 후 다음 시작 준비
        elif ((buy_average1 > 0) and (buy_average2 > 0) and (coin1 == 0) and (coin2 == 0)) :
            post_message(myToken,"#hjs-autoupbit", "모두 팔았어요! 10분 후 다시 시작할게요!")
            time.sleep[600]

#-------------------------------------------------------------

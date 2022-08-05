import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import time, calendar
import requests



# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  


def send_message(msg):
    DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1001556029412229140/Mv4NJ5nXxTZx47vlsAcYP5Ivc6i8a5EZU9bCpnfzqVa7B3r5almLzm4y-XBpbH-6h5Hs"
    """디스코드 메시지 전송"""
    now = datetime.now()
    message = {"content": f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] {str(msg)}"}
    requests.post(DISCORD_WEBHOOK_URL, data=message)
    print(message)

def check_cybos_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        send_message('check_cybos_system() : admin user -> FAILED')
        return False
 
    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        send_message('check_cybos_system() : connect to server -> FAILED')
        return False
 
    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        send_message('check_cybos_system() : init trade -> FAILED')
        return False
    return True

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매도호가, 매수호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매도호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매수호가    
    return item['cur_price'], item['ask'], item['bid']

def get_ohlc(code, qty):
    """인자로 받은 종목의 OHLC 가격 정보를 qty 개수만큼 반환한다."""
    cpOhlc.SetInputValue(0, code)           # 종목코드
    cpOhlc.SetInputValue(1, ord('2'))        # 1:기간, 2:개수
    cpOhlc.SetInputValue(4, qty)             # 요청개수
    cpOhlc.SetInputValue(5, [0, 2, 3, 4, 5]) # 0:날짜, 2~5:OHLC
    cpOhlc.SetInputValue(6, ord('D'))        # D:일단위
    cpOhlc.SetInputValue(9, ord('1'))        # 0:무수정주가, 1:수정주가
    cpOhlc.BlockRequest()
    count = cpOhlc.GetHeaderValue(3)   # 3:수신개수
    columns = ['open', 'high', 'low', 'close']
    index = []
    rows = []
    for i in range(count): 
        index.append(cpOhlc.GetDataValue(0, i)) 
        rows.append([cpOhlc.GetDataValue(1, i), cpOhlc.GetDataValue(2, i),
            cpOhlc.GetDataValue(3, i), cpOhlc.GetDataValue(4, i)]) 
    df = pd.DataFrame(rows, columns=columns, index=index) 
    return df

def get_stock_balance(code):
    """인자로 받은 종목의 종목명과 수량을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpBalance.SetInputValue(0, acc)         # 계좌번호
    cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
    cpBalance.SetInputValue(2, 50)          # 요청 건수(최대 50)
    cpBalance.BlockRequest()     
    if code == 'ALL':
        send_message(f'계좌명:  {cpBalance.GetHeaderValue(0)}')
        send_message(f'결제잔고수량 : {cpBalance.GetHeaderValue(1)}')
        send_message(f'평가금액: {cpBalance.GetHeaderValue(3)}')
        send_message(f'평가손익: {cpBalance.GetHeaderValue(4)}')
        send_message(f'종목수: {cpBalance.GetHeaderValue(7)}')
    stocks = []
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        if code == 'ALL':
            send_message(f'{i+1} {stock_code} {stock_name} : {stock_qty}')
            stocks.append({'code': stock_code, 'name': stock_name, 
                'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        return stocks
    else:
        stock_name = cpCodeMgr.CodeToName(code)
        return stock_name, 0

def get_current_cash():
    """증거금 100% 주문 가능 금액을 반환한다."""
    cpTradeUtil.TradeInit()
    acc = cpTradeUtil.AccountNumber[0]    # 계좌번호
    accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체, 1:주식, 2:선물/옵션
    cpCash.SetInputValue(0, acc)              # 계좌번호
    cpCash.SetInputValue(1, accFlag[0])      # 상품구분 - 주식 상품 중 첫번째
    cpCash.BlockRequest() 
    return cpCash.GetHeaderValue(9) # 증거금 100% 주문 가능 금액

def get_target_price(code):
    """매수 목표가를 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 10)
        if str_today == str(ohlc.iloc[0].name):
            today_open = ohlc.iloc[0].open 
            lastday = ohlc.iloc[1]
        else:
            lastday = ohlc.iloc[0]                                      
            today_open = lastday[3]
        lastday_high = lastday[1]
        lastday_low = lastday[2]
        target_price = today_open + (lastday_high - lastday_low) * 0.3      #전날 변동성 기준 20퍼, default 0.5
        return target_price
    except Exception as ex:
        send_message(f"매수 목표가, 예외 내용: {ex}")
        return None
    
def get_movingaverage(code, window):
    """인자로 받은 종목에 대한 이동평균가격을 반환한다."""
    try:
        time_now = datetime.now()
        str_today = time_now.strftime('%Y%m%d')
        ohlc = get_ohlc(code, 20)
        if str_today == str(ohlc.iloc[0].name):
            lastday = ohlc.iloc[1].name
        else:
            lastday = ohlc.iloc[0].name
        closes = ohlc['close'].sort_index()         
        ma = closes.rolling(window=window).mean()
        return ma.loc[lastday]
    except Exception as ex:
        send_message(f'이평선, {window} 예외 내용: {ex}')
        return None    

def buy(code):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list      # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list: # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            #send_message('code:', code, 'in', bought_list)
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = get_current_price(code) 
        target_price = get_target_price(code)    # 매수 목표가
        ma5_price = get_movingaverage(code, 5)   # 5일 이동평균가
        ma10_price = get_movingaverage(code, 10) # 10일 이동평균가
        buy_qty = 0        # 매수할 수량 초기화
        if ask_price > 0:  # 매도호가가 존재하면   
            buy_qty = buy_amount // ask_price  
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
        #send_message('bought_list:', bought_list, 'len(bought_list):',
        #    len(bought_list), 'target_buy_count:', target_buy_count)     
        #if current_price > target_price and current_price > ma5_price:                                    # 매수 전략 조건
        send_message(f'{stock_name} {code} 매수 수량: {buy_qty} {current_price} 매수 조건 성립!')            
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                
        # 최유리 FOK 매수 주문 설정
        cpOrder.SetInputValue(0, "2")        # 2: 매수
        cpOrder.SetInputValue(1, acc)        # 계좌번호
        cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
        cpOrder.SetInputValue(3, code)       # 종목코드
        cpOrder.SetInputValue(4, buy_qty)    # 매수할 수량
        cpOrder.SetInputValue(7, "0")        # 주문조건 0:기본, 1:IOC, 2:FOK
        cpOrder.SetInputValue(8, "13")       # 주문호가 1:보통, 3:시장가
                                                 # 5:조건부, 12:최유리, 13:최우선 
        # 매수 주문 요청
        ret = cpOrder.BlockRequest() 
        send_message(f'매수, 종목명; 코드; 수량, {stock_name}, {code}, {buy_qty} -> {ret}')
        if ret == 4:
            remain_time = cpStatus.LimitRequestRemainTime
            send_message(f'주의: 연속 주문 제한에 걸림. 대기 시간: {remain_time/1000}')
            time.sleep(remain_time/1000) 
            return False
        time.sleep(2)
        send_message(f'현금주문 가능금액 :  {buy_amount}')
        stock_name, bought_qty = get_stock_balance(code)
        send_message(f'종목명, 보유수량 : {stock_name}, {stock_qty}')
        if bought_qty > 0:
            bought_list.append(code)
            send_message(f"종목 매수. 종목: {stock_name} ; {code}, 매수 수량: {bought_qty}")
    except Exception as ex:
        send_message(f"매수, 종목 코드: {code} 예외 내용: {ex}")

def sell_all():
    """보유한 모든 종목을 최유리 지정가 IOC 조건으로 매도한다."""
    try:
        cpTradeUtil.TradeInit()
        acc = cpTradeUtil.AccountNumber[0]       # 계좌번호
        accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션   
        while True:    
            stocks = get_stock_balance('ALL') 
            total_qty = 0 
            for s in stocks:
                total_qty += s['qty'] 
            if total_qty == 0:
                return True
            for s in stocks:
                if s['qty'] != 0:                  
                    cpOrder.SetInputValue(0, "1")         # 1:매도, 2:매수
                    cpOrder.SetInputValue(1, acc)         # 계좌번호
                    cpOrder.SetInputValue(2, accFlag[0])  # 주식상품 중 첫번째
                    cpOrder.SetInputValue(3, s['code'])   # 종목코드
                    cpOrder.SetInputValue(4, s['qty'])    # 매도수량
                    cpOrder.SetInputValue(7, "1")   # 조건 0:기본, 1:IOC, 2:FOK
                    cpOrder.SetInputValue(8, "12")  # 호가 12:최유리, 13:최우선 
                    # 최유리 IOC 매도 주문 요청
                    ret = cpOrder.BlockRequest()
                    send_message(f"최유리 IOC 매도 {s['code']}, {s['name']}, {s['qty']}, -> cpOrder.BlockRequest() -> returned, {ret}")
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        send_message(f'주의: 연속 주문 제한, 대기시간: {remain_time/1000}')
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        send_message(f"매도 시에 예외 발생: {ex}")

if __name__ == '__main__': 
    try:
        symbol_list = ['A100840', 'A267290', 'A003780', 'A339770', 'A001250', 'A227840']
        bought_list = []     # 매수 완료된 종목 리스트
        target_buy_count = 6 # 매수할 종목 수   target_buy_count x buy_percent <= 1 - 수수료
        buy_percent = 0.16  #몇퍼 살껀지 0.0028 <-  0.01
        send_message('=================================================================')
        send_message(f'증권 서버 접속 상태 : {check_cybos_system()}')  # 크레온 접속 점검
        stocks = get_stock_balance('ALL')      # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())   # 100% 증거금 주문 가능 금액 조회
        buy_amount = total_cash * buy_percent  # 종목별 주문 금액 계산
        send_message(f'100% 증거금 주문 가능 금액 : {total_cash}')
        send_message(f'종목별 주문 비율 : {buy_percent}')
        send_message(f'종목별 주문 금액 : {buy_amount}')
        send_message(f"시작 시간 : {datetime.now().strftime('%m/%d %H:%M:%S')}")
        soldout = False

        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0,microsecond=0)
            today = datetime.today().weekday()
            if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
                if today == 5:
                    send_message('Today is Saturday.' )
                else:
                    send_message('Today is Sunday.')
                sys.exit(0)
            #if t_9 < t_now < t_start and soldout == False:
            #    soldout = True
            #    sell_all()
            if t_start < t_now < t_sell :  # AM 09:05 ~ PM 03:15 : 매수
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy(sym)
                        time.sleep(0.5)
                if t_now.minute == 30 and 0 <= t_now.second <= 5: 
                    get_stock_balance('ALL')
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : 일괄 매도
                #if sell_all() == True:
                send_message('모두 팔려서 자동 종료되었습니다.')
                sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                send_message('장이 마감되어 자동 종료되었습니다.')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        send_message(f'메인 프로그램 예외 발생: {ex}')

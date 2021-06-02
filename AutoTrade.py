import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import time, calendar

import AutoConnect
import GetPrice
import Logger
import PickStock
 
# 크레온 플러스 공통 OBJECT
cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')  

def check_creon_system():
    """크레온 플러스 시스템 연결 상태를 점검한다."""
    # 관리자 권한으로 프로세스 실행 여부
    if not ctypes.windll.shell32.IsUserAnAdmin():
        Logger.printlog('check_creon_system() : admin user -> FAILED')
        return False
 
    # 연결 여부 체크
    if (cpStatus.IsConnect == 0):
        Logger.printlog('check_creon_system() : connect to server -> FAILED')
        Logger.printlog('Auto Restart!')
        AutoConnect.start()

    # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    if (cpTradeUtil.TradeInit(0) != 0):
        Logger.printlog('check_creon_system() : init trade -> FAILED')
        return False
    return True

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
        # Logger.dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
        # Logger.dbgout('결제잔고수량 : ' + str(cpBalance.GetHeaderValue(1)))
        str1 = '평가금액: ' + str(cpBalance.GetHeaderValue(3))
        str2 = '평가손익: ' + str(cpBalance.GetHeaderValue(4))
        str3 = '종목수: ' + str(cpBalance.GetHeaderValue(7))
        account = '주식 잔고 요약\n' + str1 + '\n' + str2 + '\n' + str3
        Logger.dbgout(account)
    stocks = []
    stockList = ''
    for i in range(cpBalance.GetHeaderValue(7)):
        stock_code = cpBalance.GetDataValue(12, i)  # 종목코드
        stock_name = cpBalance.GetDataValue(0, i)   # 종목명
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        stock_qty = cpBalance.GetDataValue(15, i)   # 수량
        stock_pric = cpBalance.GetDataValue(8, i)   # 현재가
        if code == 'ALL':
            stock = '\n' + str(i+1) + ' ' + stock_name + '(' + str(stock_pric) + ')' + ': ' + str(stock_qty)
            stockList += stock
            stocks.append({'code': stock_code, 'name': stock_name, 'qty': stock_qty})
        if stock_code == code:  
            return stock_name, stock_qty
    if code == 'ALL':
        Logger.dbgout('보유 주식 수량' + stockList)
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

def buy_etf(code):
    """인자로 받은 종목을 최유리 지정가 FOK 조건으로 매수한다."""
    try:
        global bought_list      # 함수 내에서 값 변경을 하기 위해 global로 지정
        if code in bought_list: # 매수 완료 종목이면 더 이상 안 사도록 함수 종료
            #Logger.printlog('code:', code, 'in', bought_list)
            return False
        time_now = datetime.now()
        current_price, ask_price, bid_price = GetPrice.get_current_price(code) 
        target_price = GetPrice.get_target_price(code)    # 매수 목표가
        ma5_price = GetPrice.get_movingaverage(code, 5)   # 5일 이동평균가
        ma10_price = GetPrice.get_movingaverage(code, 10) # 10일 이동평균가
        buy_qty = 0        # 매수할 수량 초기화
        if ask_price > 0:  # 매수호가가 존재하면   
            buy_qty = int(buy_amount // ask_price)
        stock_name, stock_qty = get_stock_balance(code)  # 종목명과 보유수량 조회
        #Logger.printlog('bought_list:', bought_list, 'len(bought_list):',
        #    len(bought_list), 'target_buy_count:', target_buy_count)     
        if current_price > target_price and current_price > ma5_price \
            and current_price > ma10_price:  
            Logger.printlog(stock_name + '(' + str(code) + ') ' + str(buy_qty) + '주 (현재가: ' + 
                    str(current_price) + ') meets the buy condition!`')            
            cpTradeUtil.TradeInit()
            acc = cpTradeUtil.AccountNumber[0]      # 계좌번호
            accFlag = cpTradeUtil.GoodsList(acc, 1) # -1:전체,1:주식,2:선물/옵션                
            # 최유리 FOK 매수 주문 설정
            cpOrder.SetInputValue(0, "2")        # 2: 매수
            cpOrder.SetInputValue(1, acc)        # 계좌번호
            cpOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
            cpOrder.SetInputValue(3, code)       # 종목코드
            cpOrder.SetInputValue(4, buy_qty)    # 매수할 수량
            cpOrder.SetInputValue(7, "2")        # 주문조건 0:기본, 1:IOC, 2:FOK
            cpOrder.SetInputValue(8, "12")       # 주문호가 1:보통, 3:시장가
                                                 # 5:조건부, 12:최유리, 13:최우선 
            # 매수 주문 요청
            ret = cpOrder.BlockRequest()
            codeStr = '(' + code + ')'
            Logger.printlog('최유리 FoK 매수 ->', stock_name, codeStr, buy_qty, '->', ret)
            Logger.printlog('거래 결과 ->',cpOrder.GetDibStatus(), cpOrder.GetDibMsg1())
            if ret == 4:
                remain_time = cpStatus.LimitRequestRemainTime
                Logger.printlog('주의: 연속 주문 제한에 걸림. 대기 시간:', remain_time/1000)
                time.sleep(remain_time/1000) 
                return False
            time.sleep(2)
            Logger.printlog('현금주문 가능금액 :', buy_amount, ' / 현재 증거금 : ', get_current_cash())
            stock_name, bought_qty = get_stock_balance(code)
            
            # Logger.printlog('현재 계좌에 존재하는 ' + stock_name + ' 개수: ' + str(stock_qty))
            if bought_qty > 0:
                bought_list.append(code)
                Logger.dbgout(str(stock_name) + "(" + str(code) + "): " + str(bought_qty) + "주 구매완료! (" + str(current_price) + "/1주)")
    except Exception as ex:
        Logger.dbgout("`buy_etf("+ str(code) + ") -> exception! " + str(ex) + "`")

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
                    Logger.printlog('최유리 IOC 매도', s['code'], s['name'], s['qty'], 
                        '-> cpOrder.BlockRequest() -> returned', ret)
                    if ret == 4:
                        remain_time = cpStatus.LimitRequestRemainTime
                        Logger.printlog('주의: 연속 주문 제한, 대기시간:', remain_time/1000)
                time.sleep(1)
            time.sleep(30)
    except Exception as ex:
        Logger.dbgout("sell_all() -> exception! " + str(ex))

if __name__ == '__main__': 
    try:
        today = datetime.today().weekday()
        if today == 5 or today == 6:  # 토요일이나 일요일이면 자동 종료
            sys.exit(0)
        symbol_list = PickStock.stock_info()
        bought_list = []     # 매수 완료된 종목 리스트
        target_buy_count = 10 # 매수할 종목 수
        buy_percent = 0.1
        Logger.printlog('check_creon_system() :', check_creon_system())  # 크레온 접속 점검
        # stocks = get_stock_balance('ALL')      # 보유한 모든 종목 조회
        total_cash = int(get_current_cash())   # 100% 증거금 주문 가능 금액 조회
        buy_amount = int(total_cash * buy_percent)  # 종목별 주문 금액 계산
        Logger.dbgout('100% 증거금 주문 가능 금액 :' + str(int(get_current_cash())))
        Logger.printlog('100% 증거금 주문 가능 금액 :', total_cash)
        Logger.printlog('종목별 주문 비율 :', buy_percent)
        Logger.printlog('종목별 주문 금액 :', buy_amount)
        Logger.printlog('시작 시간 :', datetime.now().strftime('%m/%d %H:%M:%S'))
        soldout = False
        
        while True:
            t_now = datetime.now()
            t_9 = t_now.replace(hour=9, minute=0, second=0, microsecond=0)
            t_start = t_now.replace(hour=9, minute=5, second=0, microsecond=0)
            t_sell = t_now.replace(hour=15, minute=15, second=0, microsecond=0)
            t_exit = t_now.replace(hour=15, minute=20, second=0,microsecond=0)
            if t_9 < t_now < t_start and soldout == False:
                soldout = True
                sell_all()
            if t_start < t_now < t_sell :  # AM 09:05 ~ PM 03:15 : 매수
                for sym in symbol_list:
                    if len(bought_list) < target_buy_count:
                        buy_etf(sym)
                        time.sleep(1)
                if t_now.minute == 30 and 0 <= t_now.second <= 5: 
                    get_stock_balance('ALL')
                    Logger.dbgout('100% 증거금 주문 가능 금액 :' + str(int(get_current_cash())))
                    time.sleep(5)
            if t_sell < t_now < t_exit:  # PM 03:15 ~ PM 03:20 : 일괄 매도
                if sell_all() == True:
                    Logger.dbgout('`sell_all() returned True -> self-destructed!`')
                    sys.exit(0)
            if t_exit < t_now:  # PM 03:20 ~ :프로그램 종료
                Logger.dbgout('`self-destructed!`')
                sys.exit(0)
            time.sleep(3)
    except Exception as ex:
        Logger.dbgout('`main -> exception! ' + str(ex) + '`')
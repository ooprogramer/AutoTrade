import win32com.client
import pandas as pd
from datetime import datetime

import Logger

cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)   # 현재가
    item['ask'] =  cpStock.GetHeaderValue(16)        # 매수호가
    item['bid'] =  cpStock.GetHeaderValue(17)        # 매도호가    
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
        target_price = today_open + (lastday_high - lastday_low) * 0.3  # k값
        return target_price
    except Exception as ex:
        Logger.dbgout("`get_target_price() -> exception! " + str(ex) + "`")
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
        Logger.dbgout('get_movingavrg(' + str(window) + ') -> exception! ' + str(ex))
        return None    
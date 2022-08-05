import talib
from pykrx import stock
import win32com.client

import time
instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = instCpCodeMgr.GetStockListByMarket(1)
t2 = time.strftime('%Y%m%d', time.localtime(time.time()))

import datetime
from dateutil.relativedelta import relativedelta
 

a=datetime.datetime.now()-relativedelta(months=1)
t1 = a.strftime('%Y%m%d')

serched = []
serched2 = []
def crazy(code):
    cd = str(code)
    cd = cd[1:]
    data = stock.get_market_ohlcv_by_date(t1, t2, cd)
    data = data.reset_index()
    data.columns = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
    engulfing = talib.CDLENGULFING(data['Open'], data['High'], data['Low'], data['Close'])
    data['Engulfing'] = engulfing
    engulfing_days = data[data['Engulfing'] == 100] 
    
    if len(engulfing_days) > 0:
        print(code)
        print(engulfing_days['Date'].to_string(index=False))

        if len(engulfing_days['Date'].to_string(index=False)) > 10:
            serched2.append(code)
        serched.append(code)

    return serched, serched2

a = ['A005930', 'A373220', 'A000660', 'A207940', 'A005935', 'A051910', 'A035420', 'A005380', 'A006400', 'A000270', 'A035720', 'A068270', 'A028260', 'A012330', 'A005490', 'A105560', 'A055550', 'A096770', 'A034730', 
'A066570', 'A323410', 'A015760', 'A003550', 'A247540', 'A011200', 'A329180', 'A051900', 'A259960', 'A032830', 'A034020', 'A017670', 'A091990', 'A033780', 'A086790', 'A009150', 'A003670', 'A018260', 'A010950', 'A030200', 'A000810', 'A302440', 'A003490', 'A010130', 'A066970', 'A316140', 'A011070', 'A009830', 'A377300', 'A036570', 'A090430', 'A024110', 'A352820', 'A086280', 'A009540', 'A383220', 'A011170', 'A326030', 'A402340', 'A361610', 'A097950', 'A251270', 'A047810', 'A018880', 'A035250', 'A034220', 'A032640', 'A088980', 'A011790', 'A010140', 'A069500', 'A000720', 'A021240', 'A028300', 'A267250', 'A004020', 'A161390', 'A005830', 'A000100', 'A180640', 'A137310', 'A293490', 'A006800', 'A000060', 'A010620', 'A271560', 'A011780', 'A028050', 'A078930', 'A004990', 'A128940', 'A029780', 'A371460', 'A071050', 'A020150', 'A005387', 'A003410', 'A307950', 'A263750', 'A138040', 'A036460', 'A005940', 'A012450', 'A282330', 'A068760', 'A241560', 'A016360', 'A028670', 'A008560']
for a in a:
    crazy(a)
#for code in codeList:
#    crazy(code)
print("한번: ", serched,"\n", "여러번:", serched2)

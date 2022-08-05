import sys
import win32com.client
import ctypes
 
################################################
# PLUS 공통 OBJECT
g_objCodeMgr = win32com.client.Dispatch('CpUtil.CpCodeMgr')
g_objCpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
g_objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
 
 
################################################
# PLUS 실행 기본 체크 함수
def InitPlusCheck():
    # 프로세스가 관리자 권한으로 실행 여부
    if ctypes.windll.shell32.IsUserAnAdmin():
        print('정상: 관리자권한으로 실행된 프로세스입니다.')
    else:
        print('오류: 일반권한으로 실행됨. 관리자 권한으로 실행해 주세요')
        return False
 
    # 연결 여부 체크
    if (g_objCpStatus.IsConnect == 0):
        print("PLUS가 정상적으로 연결되지 않음. ")
        return False
 
    # # 주문 관련 초기화 - 계좌 관련 코드가 있을 때만 사용
    # if (g_objCpTrade.TradeInit(0) != 0):
    #     print("주문 초기화 실패")
    #     return False
 
    return True
 
 
class CpMarketEye:
    def __init__(self):
        self.objRq = win32com.client.Dispatch("CpSysDib.MarketEye")
        self.RpFiledIndex = 0
 
 
    def Request(self, codes, dataInfo):
        # 0: 종목코드 4: 현재가 20: 상장주식수
        rqField = [0, 4, 20]  # 요청 필드
 
        self.objRq.SetInputValue(0, rqField)  # 요청 필드
        self.objRq.SetInputValue(1, codes)  # 종목코드 or 종목코드 리스트
        self.objRq.BlockRequest()
 
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        print("통신상태", rqStatus, self.objRq.GetDibMsg1())
        if rqStatus != 0:
            return False
 
        cnt = self.objRq.GetHeaderValue(2)
 
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            cur = self.objRq.GetDataValue(1, i)  # 종가
            listedStock = self.objRq.GetDataValue(2, i)  # 상장주식수
 
            maketAmt = listedStock * cur
            if g_objCodeMgr.IsBigListingStock(code) :
                maketAmt *= 1000
#            print(code, maketAmt)
 
            # key(종목코드) = tuple(상장주식수, 시가총액)
            dataInfo[code] = (listedStock, maketAmt)
 
        return True
 
class CMarketTotal():
    def __init__(self):
        self.dataInfo = {}
 
 
    def GetAllMarketTotal(self):
        codeList = g_objCodeMgr.GetStockListByMarket(1)  # 거래소
        codeList2 = g_objCodeMgr.GetStockListByMarket(2)  # 코스닥
        allcodelist = codeList + codeList2
        print('전 종목 코드 %d, 거래소 %d, 코스닥 %d' % (len(allcodelist), len(codeList), len(codeList2)))
 
        objMarket = CpMarketEye()
        rqCodeList = []
        for i, code in enumerate(allcodelist):
            rqCodeList.append(code)
            if len(rqCodeList) == 200:
                objMarket.Request(rqCodeList, self.dataInfo)
                rqCodeList = []
                continue
        # end of for
 
        if len(rqCodeList) > 0:
            objMarket.Request(rqCodeList, self.dataInfo)
 
    def PrintMarketTotal(self):
        k=[]
        # 시가총액 순으로 소팅
        data2 = sorted(self.dataInfo.items(), key=lambda x: x[1][1], reverse=True)
        print(data2)
        print('전종목 시가총액 순 조회 (%d 종목)' % (len(data2)))
        for item in data2:
            name = item[0]
            listed = item[1][0]
            markettot = item[1][1] 
            if (markettot> 3000000000000):
                k.append(name)
            
        print(k)
        print(len(['A005930', 'A373220', 'A000660', 'A207940', 'A005935', 'A051910', 'A035420', 'A005380', 'A006400', 'A000270', 'A035720', 'A068270', 'A028260', 'A012330', 'A005490', 'A105560', 'A055550', 'A096770', 'A034730', 
'A066570', 'A323410', 'A015760', 'A003550', 'A247540', 'A011200', 'A329180', 'A051900', 'A259960', 'A032830', 'A034020', 'A017670', 'A091990', 'A033780', 'A086790', 'A009150', 'A003670', 'A018260', 'A010950', 'A030200', 'A000810', 'A302440', 'A003490', 'A010130', 'A066970', 'A316140', 'A011070', 'A009830', 'A377300', 'A036570', 'A090430', 'A024110', 'A352820', 'A086280', 'A009540', 'A383220', 'A011170', 'A326030', 'A402340', 'A361610', 'A097950', 'A251270', 'A047810', 'A018880', 'A035250', 'A034220', 'A032640', 'A088980', 'A011790', 'A010140', 'A069500', 'A000720', 'A021240', 'A028300', 'A267250', 'A004020', 'A161390', 'A005830', 'A000100', 'A180640', 'A137310', 'A293490', 'A006800', 'A000060', 'A010620', 'A271560', 'A011780', 'A028050', 'A078930', 'A004990', 'A128940', 'A029780', 'A371460', 'A071050', 'A020150', 'A005387', 'A003410', 'A307950', 'A263750', 'A138040', 'A036460', 'A005940', 'A012450', 'A282330', 'A068760', 'A241560', 'A016360', 'A028670', 'A008560']))
                
            #print('%s 상장주식수: %s, 시가총액 %s' %(name, format(listed, ','), format(markettot, ',')))
 
 
if __name__ == "__main__":
    objMarketTotal = CMarketTotal()
    objMarketTotal.GetAllMarketTotal()
    objMarketTotal.PrintMarketTotal()
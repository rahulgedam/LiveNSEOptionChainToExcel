from datetime import datetime, time
from time import sleep
import requests
import xlwings as xw
import json
import math

# Owned
__author__     = "Rahul Gedam"
__copyright__  = "Copyright 2022-2035, Read Live Option Chain in Excel."
__credits__    = ["Rahul Gedam"]
__license__    = "Public"
__version__    = "1.0.0"
__maintainer__ = "Rahul Gedam"
__status__     = "Development"


#Function add different option strike data in respective option strike sheet
def addOptionDataInSheet(wb, currTime,strikePrice,ceLTP,ceOI,ceOIChange,peLTP,peOI,peOIChange,lastRow,instrName,ATM,optExpiryDate):
    wb_sheet = wb.sheets['OptChain'+optExpiryDate]
    wb_sheet.cells(lastRow + 1, 1).value = currTime
    wb_sheet.cells(lastRow + 1, 2).value = ceLTP
    wb_sheet.cells(lastRow + 1, 3).value = ceOI
    wb_sheet.cells(lastRow + 1, 4).value = ceOIChange
    wb_sheet.cells(lastRow + 1, 5).value = strikePrice
    wb_sheet.cells(lastRow + 1, 6).value = peLTP
    wb_sheet.cells(lastRow + 1, 7).value = peOI
    wb_sheet.cells(lastRow + 1, 8).value = peOIChange
    #Color strike Price
    wb_sheet.range(lastRow + 1, 5).color = (255, 255, 0)
    return wb

#Function to calculate over option chain
def createOptionChain(wb, optionData, ATM, lotSize, instrName, optExpDate):
    lastRow = 0
    for optionRow in optionData:
        if optionRow['expiryDate'] == optExpDate:
          try:
           strikePrice = optionRow["CE"]["strikePrice"]
           ceLTP = optionRow["CE"]["lastPrice"]
           ceOI = lotSize*optionRow["CE"]["openInterest"]
           ceOIChange = lotSize*optionRow["CE"]["changeinOpenInterest"]
          except:
              ceLTP = 0
              ceOI = 0
              ceOIChange = 0
              pass
          try:
           strikePrice = optionRow["PE"]["strikePrice"]
           peLTP = optionRow["PE"]["lastPrice"]
           peOI = lotSize*optionRow["PE"]["openInterest"]
           peOIChange = lotSize*optionRow["PE"]["changeinOpenInterest"]
          except:
             peLTP = 0
             peOI = 0
             peOIChange = 0             
             pass     
          lastRow = lastRow + 1
          wb = addOptionDataInSheet(wb, currTime, strikePrice, ceLTP, ceOI, ceOIChange, peLTP, peOI, peOIChange, lastRow, instrName, ATM, optExpDate)
    return wb         

#Function make excel sheet if not present already
def makeOptionChainFile(wb, opt,currTime,ATM,instrName, optExpDate):
    if(instrName == "NIFTY"):
      lotSize = 50
    else:
      lotSize = 25
    wb_sheet = wb.sheets['OptChain'+optExpDate]
    wb_sheet.cells(1, 1).value = "Time"
    wb_sheet.cells(1, 2).value = "CALL LTP"
    wb_sheet.cells(1, 3).value = "CALL OI"
    wb_sheet.cells(1, 4).value = "CALL OI CHANGE"
    wb_sheet.cells(1, 5).value = "STRIKEPRICE"
    wb_sheet.cells(1, 6).value = "PUT LTP"
    wb_sheet.cells(1, 7).value = "PUT OI"    
    wb_sheet.cells(1, 8).value = "PUT OI CHANGE"    

    optionData = opt["records"]["data"]
    wb = createOptionChain(wb, optionData, ATM, lotSize, instrName, optExpDate)
    return currTime, ATM

#Function to add data in OptionChain.xlsx
def putOptionChainData(opt, currTime, ATM, instrName, optExpDate):
    wb = xw.Book('OptionChain'+ instrName + '.' + 'xlsx')
    optionData = opt["records"]["data"]
    if(instrName == "NIFTY"):
      lotSize = 50
    else:
      lotSize = 25    
    wb = createOptionChain(wb,optionData, ATM, lotSize,instrName, optExpDate)
    wb.save()
    return currTime, ATM

def findATM(fut,instrName):
    lastTraded = float((fut["data"][0]["lastPrice"]).replace(',', ''))
    if(instrName == "NIFTY"):
      factor = 50
    else:
      factor = 100
    return int(math.floor(lastTraded / factor)) * factor

def getOptionChain(currTime,isFifteenMin,instrName):
    try:
        #step 1: Read Option chain and extract list of expiries.
        oc_url = "https://www.nseindia.com/api/option-chain-indices?symbol="+ instrName
        oc_headers = {"accept-encoding": "gzip, deflate, br",
           "accept-language": "en-US,en;q=0.9,mr-IN;q=0.8,mr;q=0.7,hi-IN;q=0.6,hi;q=0.5",
           "referer":"https://www.nseindia.com/option-chain?symbol=NIFTY&instrument=-&date=-",
           "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"}        
        session = requests.session()

        opt = session.get(oc_url, headers=oc_headers).json()
        # Pull Json file
        r = requests.get(oc_url, headers=oc_headers)
        content = r.json()
        with open("OptionOI"+ instrName + "." + "json", "w") as json_file:
            json.dump(content, json_file, indent=4)

        optionExpDate, futuresExpDate = readConfigJson()          

        m = 0
        for j in optionExpDate:
          try:
            fut = downloadFuturesJson(instrName, futuresExpDate[m])
            ATM = findATM(fut, instrName)
            print(instrName + " ATM-----> ", ATM) 
            #Build Option Chain Data
            putOptionChainData(opt, currTime, ATM, instrName, optionExpDate[m])
            m = m + 1 
          except Exception as t:
            print(t)

    except Exception as e:
        print(e)


def createExcelSheet(instrName):
    wb = xw.Book()
    wb.save('OptionChain'+ instrName + '.' + 'xlsx')
    return wb

def createSheetsInExcel(optionExpDate, wb):
    count = 0
    for k in optionExpDate:
        try:
          wb.sheets.add(name='OptChain'+optionExpDate[count])
          count = count + 1
        except:
            print("Error in createSheetsInExcel")
            pass
    for sheet in wb.sheets:
        if 'Sheet' in sheet.name:
         sheet.delete()            
    return wb    
#-----------------------------------------------------------------------------------------------------------------
def downloadFuturesJson(instrName, futExpDate):
        fut_url = "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/ajaxFOGetQuoteJSON.jsp?underlying="+instrName+"&instrument=FUTIDX&expiry="+futExpDate
        referer = "https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuoteFO.jsp?underlying="+instrName+"&instrument=OPTIDX&expiry"+futExpDate+"&type=CE&strike=18500.00"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9', 'Accept-Encoding': 'gzip, deflate, br', 'referer': referer}
        cookies = {
            'bm_sv': '808AA9DAA6889DCF6CB5A9AB14BF290F~t420SFeAK8epqchwzNEdgzJGSy4pjs8Mco4O/d7w5qW4TxShiDJYDHoeSAe60uaZMH2n5wjH3v0NnmtPWrKZosEZDfOSHts34ur4GGAPsGGL3LHYVPU/IbcIDhNf6MXF8oqWfpyJ0FXh74VTcL2NhnLG6DlsLun9OrT/+t9vHpY='}
        session = requests.session()

        for cook in cookies:
            session.cookies.set(cook, cookies[cook])
        fut = session.get(fut_url, headers=headers).json()
        # Pull Fut Json file
        r1 = requests.get(fut_url, headers=headers)
        content = r1.json()
        with open("FutureOI"+ instrName + "." + "json", "w") as json_file:
            json.dump(content, json_file, indent=4)
        
        return fut
#-----------------------------------------------------------------------------------------------------------------

def readConfigJson():
    with open('config.json') as user_file:
     parsed_json = json.load(user_file)
    optionExpDate = {}
    futuresExpDate = {}
    n = 0

    for i in parsed_json['data']:
        try:
            optionExpDate[n]  = i['expiryDateOption']
            futuresExpDate[n] = i['expiryDateFuture']
            n = n+1
        except:
            pass

    return optionExpDate, futuresExpDate

def createAndInitFiles(currTime, instrName):
    try:
        #step 1: Read Option chain and extract list of expiries.
        oc_url = "https://www.nseindia.com/api/option-chain-indices?symbol="+ instrName
        oc_headers = {"accept-encoding": "gzip, deflate, br",
           "accept-language": "en-US,en;q=0.9,mr-IN;q=0.8,mr;q=0.7,hi-IN;q=0.6,hi;q=0.5",
           "referer":"https://www.nseindia.com/option-chain?symbol=NIFTY&instrument=-&date=-",
           "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"}        
        session = requests.session()

        opt = session.get(oc_url, headers=oc_headers).json()
        # Pull Json file
        r = requests.get(oc_url, headers=oc_headers)
        content = r.json()
        with open("OptionOI"+ instrName + "." + "json", "w") as json_file:
            json.dump(content, json_file, indent=4)

        optionExpDate, futuresExpDate = readConfigJson()          
        wb = createExcelSheet(instrName)
        createSheetsInExcel(optionExpDate, wb)

        m = 0
        for j in optionExpDate:
         try:
            fut = downloadFuturesJson(instrName, futuresExpDate[m])
            ATM = findATM(fut, instrName)
            print(instrName + " ATM-----> ", ATM) 
            makeOptionChainFile(wb, opt,currTime,ATM,instrName,optionExpDate[m])
            m = m + 1
         except Exception as p:
            print(p)


        return True
    except Exception as e:
        print(e)
        return False


if __name__ == '__main__':
   success = False

   # This loop runs from 9:15 AM to 3:30 PM till Market hours
   while(1):
        t1 = datetime.now()
        currTime = str(t1.hour)+":"+str(t1.minute)
        # From time 9:15 AM to 9:17 AM fetch data and Initialize Files
        #if(time(9,15)<=datetime.now().time()<=time(9,17,30) and not(success)):
        success = createAndInitFiles(currTime, "NIFTY")
        #else:
        #  print("Not 9:15 AM Yet, Market yet to start  ", currTime)   

        isFifMin=1
        #while(time(9,20)<=datetime.now().time()<=time(15,31)):
        while(time(8,00)<=datetime.now().time()<=time(23,59)):
            t1 = datetime.now()
            currTime = str(t1.hour) + ":" + str(t1.minute)
            isFifMin+=1
            isFifteenMin = False
            if isFifMin==3:
                isFifteenMin = True
                isFifMin=0
            getOptionChain(currTime, isFifteenMin, "NIFTY")
            t1 = datetime.now()
            while(t1.minute%5 !=0 or t1.second !=0):
                t1 = datetime.now()
                sleep(1)
        sleep(1)
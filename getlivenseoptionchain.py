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
__email__      = "rahulgedam@gmail.com"
__status__     = "Development"


#Function add different option strike data in respective option strike sheet
def addOptionDataInSheet(wb, currTime, strikePrice, ceLTP,ceChips, ceHawa, ceOI, ceOIChange, peLTP,peChips, peHawa, peOI, peOIChange, lastRow, instrName, ATM, optExpiryDate):
    wb_sheet = wb.sheets['OptChain'+optExpiryDate]
    wb_sheet.cells(lastRow + 1, 1).value = currTime
    wb_sheet.cells(lastRow + 1, 2).value = ceLTP
    wb_sheet.cells(lastRow + 1, 3).value = ceChips
    wb_sheet.cells(lastRow + 1, 4).value = ceHawa
    wb_sheet.cells(lastRow + 1, 5).value = ceOI
    wb_sheet.cells(lastRow + 1, 6).value = ceOIChange
    wb_sheet.cells(lastRow + 1, 7).value = strikePrice
    wb_sheet.cells(lastRow + 1, 8).value = peLTP
    wb_sheet.cells(lastRow + 1, 9).value = peChips
    wb_sheet.cells(lastRow + 1, 10).value = peHawa
    wb_sheet.cells(lastRow + 1, 11).value = peOI
    wb_sheet.cells(lastRow + 1, 12).value = peOIChange
    #Color strike Price
    wb_sheet.range(lastRow + 1, 7).color = (255, 255, 0)
    return wb

#Function to calculate over option chain
def createOptionChain(wb, optionData, ATM, lotSize, instrName, optExpDate, FUT_SPOT):
    lastRow = 0
    for optionRow in optionData:
        if optionRow['expiryDate'] == optExpDate:
          try:
           strikePrice = optionRow["CE"]["strikePrice"]
           ceLTP = optionRow["CE"]["lastPrice"]
           ceChips = FUT_SPOT - strikePrice
           if(ceChips < 0):
             ceChips = 0
           ceHawa = ceLTP - ceChips
           if(ceHawa < 0):
             ceHawa = 0            
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
           peChips = strikePrice - FUT_SPOT
           if(peChips < 0):
             peChips = 0
           peHawa = peLTP - peChips
           if(peHawa < 0):
             peHawa = 0
           peOI = lotSize*optionRow["PE"]["openInterest"]
           peOIChange = lotSize*optionRow["PE"]["changeinOpenInterest"]
          except:
             peLTP = 0
             peOI = 0
             peOIChange = 0             
             pass     
          lastRow = lastRow + 1
          wb = addOptionDataInSheet(wb, currTime, strikePrice, ceLTP,ceChips, ceHawa, ceOI, ceOIChange, peLTP,peChips, peHawa, peOI, peOIChange, lastRow, instrName, ATM, optExpDate)
    return wb         

#Function make excel sheet if not present already
def makeOptionChainFile(wb, opt,currTime,ATM,instrName, optExpDate, FUT_SPOT):
    if(instrName == "NIFTY"):
      lotSize = 50
    else:
      lotSize = 25
    wb_sheet = wb.sheets['OptChain'+optExpDate]
    wb_sheet.cells(1, 1).value = "Time"
    wb_sheet.cells(1, 2).value = "CALL LTP"
    wb_sheet.cells(1, 3).value = "CALL CHIPS"
    wb_sheet.cells(1, 4).value = "CALL HAWA"
    wb_sheet.cells(1, 5).value = "CALL OI"
    wb_sheet.cells(1, 6).value = "CALL OI CHANGE"
    wb_sheet.cells(1, 7).value = "STRIKEPRICE"
    wb_sheet.cells(1, 8).value = "PUT LTP"
    wb_sheet.cells(1, 9).value = "PUT CHIPS"
    wb_sheet.cells(1, 10).value = "PUT HAWA"
    wb_sheet.cells(1, 11).value = "PUT OI"    
    wb_sheet.cells(1, 12).value = "PUT OI CHANGE"   

    wb_sheet.range('A1:L150').api.Borders.LineStyle = 1
    wb_sheet.range('A1:L150').api.Borders.Weight = 2 

    optionData = opt["records"]["data"]
    wb = createOptionChain(wb, optionData, ATM, lotSize, instrName, optExpDate,FUT_SPOT)
    
    return currTime, ATM

#Function to add data in OptionChain.xlsx
def putOptionChainData(opt, currTime, ATM, instrName, optExpDate,FUT_SPOT):
    wb = xw.Book('OptionChain'+ instrName + '.' + 'xlsx')
    optionData = opt["records"]["data"]
    if(instrName == "NIFTY"):
      lotSize = 50
    else:
      lotSize = 25    
    wb = createOptionChain(wb,optionData, ATM, lotSize,instrName, optExpDate,FUT_SPOT)
    wb.save()
    return currTime, ATM

def findATM(lastTraded,instrName):
    lastTraded = float(lastTraded)
    if(instrName == "NIFTY"):
      factor = 50
    else:
      factor = 100
    currATM = int(math.floor(lastTraded / factor)) * factor   
    return currATM, lastTraded

def getOptionChain(currTime,isFifteenMin,instrName, wb):
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
            lastTraded = downloadFuturesJson(instrName, futuresExpDate[m])
            ATM, FUT_SPOT = findATM(lastTraded, instrName)
            print(instrName + " ATM-----> ", ATM, "FUT SPOT----->", FUT_SPOT) 
            makeOptionChainFile(wb, opt,currTime,ATM,instrName,optionExpDate[m], FUT_SPOT)
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
# URL to fetch the session cookies
 homepage_url = "https://www.nseindia.com"

# URL to the API endpoint
 api_url = "https://www.nseindia.com/api/snapshot-derivatives-equity?index=futures"

# Headers for the initial request to the homepage
 headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "en-US,en;q=0.5",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1"
}

# Start a session to persist cookies
 session = requests.Session()

# Make the initial request to the homepage to get cookies
 response = session.get(homepage_url, headers=headers)

# Check if the initial request was successful
 if response.status_code == 200:
    # Add more detailed headers for the API request
    api_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "en-US,en;q=0.5",
        "Referer": "https://www.nseindia.com/",
        "X-Requested-With": "XMLHttpRequest",
        "Connection": "keep-alive"
    }

    # Make the request to the API endpoint with the same session
    response = session.get(api_url, headers=api_headers)

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the JSON data
        data = response.json()
        with open("OptionOIRG.json", "w") as json_file:
            json.dump(data, json_file, indent=4)    
        volume_data = data.get('volume', {}).get('data', [0])
        for item in volume_data:
            if item.get('underlying') == instrName and item.get('expiryDate') == futExpDate:
             lastPrice = item.get('lastPrice')
             break        
        #lastTraded = float((data["data"][0]["lastPrice"]).replace(',', ''))
        # Print the data
        print(lastPrice)
        return lastPrice
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}, Response: {response.text}")
 else:
    print(f"Failed to fetch homepage. Status code: {response.status_code}")
#-----------------------------------------------------------------------------------------------------------------
def readFuturesJson(instrName, futExpDate):
   
 with open('OptionOIRG.json', 'r') as json_file:
    data = json.load(json_file)
 volume_data = data.get('volume', {}).get('data', [])
 for item in volume_data:
    if item.get('underlying') == instrName and item.get('expiryDate') == futExpDate:
        lastPrice = item.get('lastPrice')
    break

 return lastPrice
#-------------------------------------------------------------------------------
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
            lastTraded = downloadFuturesJson(instrName, futuresExpDate[m])
            #lastTraded = readFuturesJson(instrName, futuresExpDate[m])
            ATM, FUT_SPOT = findATM(lastTraded, instrName)
            print(instrName + " ATM-----> ", ATM, "FUT SPOT----->", FUT_SPOT) 
            makeOptionChainFile(wb, opt,currTime,ATM,instrName,optionExpDate[m], FUT_SPOT)
            m = m + 1
         except Exception as p:
            print(p)


        return True, wb
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
        success, wb = createAndInitFiles(currTime, "NIFTY")
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
            getOptionChain(currTime, isFifteenMin, "NIFTY", wb)
            t1 = datetime.now()
            while(t1.minute%5 !=0 or t1.second !=0):
                t1 = datetime.now()
                sleep(1)
        sleep(1)
import os
import datetime
import comtypes.client as cc
import comtypes.gen.SKCOMLib as sk
import math
import time
import pythoncom
import mysql.connector
from Model import db
from Config import user

# testdb = db.conn()
# testdb.autocommit = False
# cursor=testdb.cursor()


skC = cc.CreateObject(sk.SKCenterLib, interface=sk.ISKCenterLib)
skQ = cc.CreateObject(sk.SKQuoteLib, interface=sk.ISKQuoteLib)
skR = cc.CreateObject(sk.SKReplyLib, interface=sk.ISKReplyLib)

# ID = getpass.getpass(prompt='ID=')
# PW = getpass.getpass(prompt='PW=')

ID = user.ID
PWD = user.PWD

# 顯示 event 事件


def pumpwait(t=1):
    for i in range(t):
        time.sleep(1)
        #print('pumpwait: ', datetime.datetime.now())
        pythoncom.PumpWaitingMessages()


class modelSql:
    # def initConn():

    def insert(time, StockNo, StockName, Open, High, Low, Close, Qty):
        # Insert Multiple Records
        sql = "INSERT INTO `future_info`( `time`, `StockNo`, `StockName`, `Open`, `High`, `Low`, `Close`, `Qty`) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)"
        val = (time, StockNo, StockName, Open, High, Low, Close, Qty)
        cursor.execute(sql, val)

    def s_insert(time, StockNo, StockName, Open, High, Low, Close, Qty):
        # Insert Multiple Records
        sql = "INSERT INTO `future_info`( `time`, `StockNo`, `StockName`, `Open`, `High`, `Low`, `Close`, `Qty`) SELECT * FROM (SELECT %s,%s,%s,%s,%s,%s,%s,%s) as tmp WHERE (SELECT Qty FROM future_info ORDER BY id DESC LIMIT 1) < %s OR (SELECT Qty FROM future_info ORDER BY id DESC LIMIT 1) IS NULL"
        val = (time, StockNo, StockName, Open, High, Low, Close, Qty, Qty)
        cursor.execute(sql, val)

# 建立事件類別


class skQ_events:
    def OnConnection(self, nKind, nCode):
        if (nKind == 3001):
            strMsg = "Connected!"
        elif (nKind == 3002):
            strMsg = "DisConnected!"
        elif (nKind == 3003):
            strMsg = "Stocks ready!"
        elif (nKind == 3021):
            strMsg = "Connect Error!"
        print(strMsg)

    def OnNotifyQuote(self, sMarketNo, sStockidx):
        pStock = sk.SKSTOCK()
        skQ.SKQuoteLib_GetStockByIndex(sMarketNo, sStockidx, pStock)

        Open = pStock.nOpen/math.pow(10, pStock.sDecimal)
        High = pStock.nHigh/math.pow(10, pStock.sDecimal)
        Low = pStock.nLow/math.pow(10, pStock.sDecimal)
        Close = pStock.nClose/math.pow(10, pStock.sDecimal)
        Qty = pStock.nTQty
        a = datetime.datetime.now()

        try:
            # modelSql.s_insert(a, pStock.bstrStockNo, pStock.bstrStockName, Open, High, Low, Close, Qty)
            # testdb.commit()
            strMsg = '代碼:', pStock.bstrStockNo, '--名稱:', pStock.bstrStockName, '--開盤價:', pStock.nOpen/math.pow(10, pStock.sDecimal), '--最高:', pStock.nHigh/math.pow(
                10, pStock.sDecimal), '--最低:', pStock.nLow/math.pow(10, pStock.sDecimal), '--成交價:', pStock.nClose/math.pow(10, pStock.sDecimal), '--總量:', pStock.nTQty
            print('OnNotifyQuote: ', a)
            print(strMsg)

        except mysql.connector.Error as error:
            print("Failed to update record to database rollback: {}".format(error))
            # testdb.rollback()

    def OnNotifyTicks(self, sMarketNo, sIndex, nPtr, nDate, nTimehms, nTimemillismicros, nBid, nAsk, nClose, nQty, nSimulate):
        print('OnNotifyTicks: ', datetime.datetime.now())
        print(sMarketNo, sIndex, nPtr, nTimehms, nClose, nQty)

    # def OnNotifyBest5(self, sMarketNo, sStockidx, nBestBid1, nBestBidQty1, nBestBid2, nBestBidQty2, nBestBid3, nBestBidQty3, nBestBid4, nBestBidQty4, nBestBid5, nBestBidQty5, nExtendBid, nExtendBidQty, nBestAsk1, nBestAskQty1, nBestAsk2, nBestAskQty2, nBestAsk3, nBestAskQty3, nBestAsk4, nBestAskQty4, nBestAsk5, nBestAskQty5, nExtendAsk, nExtendAskQty, nSimulate):
    #     print(nBestBidQty1,'\n',nBestBidQty2,'\n',nBestBidQty3,'\n',nBestBidQty4,'\n',nBestBidQty5)


# Event sink, 事件實體
EventQ = skQ_events()

# make connection to event sink
ConnectionQ = cc.GetEvents(skQ, EventQ)

print('Start Login')

# Login
print("Login : ", skC.SKCenterLib_GetReturnCodeMessage(
    skC.SKCenterLib_Login(ID, PWD)))
#print("ConnectByID,", skC.SKCenterLib_GetReturnCodeMessage(skR.SKReplyLib_ConnectByID(ID)))
print("EnterMonitor : ", skC.SKCenterLib_GetReturnCodeMessage(
    skQ.SKQuoteLib_EnterMonitor()))

pumpwait(15)

print("Request Stocks")
skQ.SKQuoteLib_RequestStocks(-1, 'TX00')
# skQ.SKQuoteLib_RequestTicks(-1,'TX00')

while True:
    pumpwait(1)

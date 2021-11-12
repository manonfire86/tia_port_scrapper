# -*- coding: utf-8 -*-
"""
Created on Wed May 16 15:50:55 2018

@author: santana
"""
import win32com.client
import blpapi
import tia
from tia.bbg import LocalTerminal
import pandas as pd
import tia.bbg.datamgr as dm
import os
import numpy as np
import datetime as dt

os.chdir(r'C:\Users\santana\Documents\Bloomberg Scrapes')

CUSIPDB = pd.read_excel('CUSIP Database.xlsx')

mgr = dm.BbgDataManager()

tickers = []
for i in range(len(CUSIPDB)):
    if CUSIPDB['Identifier'][i] == '55406HAA4' or CUSIPDB['Identifier'][i] == '530715AL5' or CUSIPDB['Identifier'][i] == '530715AG6' or CUSIPDB['Identifier'][i] == '87266LAA7':
        tickers.append(LocalTerminal.get_reference_data("/CUSIP/"+CUSIPDB['Identifier'][i],"SECURITY_NAME").as_frame()['SECURITY_NAME'][0])
    else:
        tickers.append(LocalTerminal.get_reference_data(CUSIPDB['Identifier'][i]+" MTGE","SECURITY_NAME").as_frame()['SECURITY_NAME'][0])



MasterDF = pd.concat([CUSIPDB,pd.DataFrame(tickers,columns = ['Ticker'])],axis=1)

security = 'ASC 1996-D2 B1B Mtge'

sids = mgr[security]

start_Date = raw_input('Input start Date (M/DD/YYYY): ')
end_Date = raw_input('Input end Date (M/DD/YYYY): ')


historicalprices = {}
for i in range(len(MasterDF['Ticker'])):
    if MasterDF['Identifier'][i] == '55406HAA4' or MasterDF['Identifier'][i] == '530715AL5' or MasterDF['Identifier'][i] == '530715AG6' or MasterDF['Identifier'][i] == '87266LAA7':
        historicalprices[MasterDF['Ticker'][i]] = mgr["/CUSIP/"+MasterDF['Identifier'][i]].get_historical('PX_Last',start=start_Date,end=end_Date,period='DAILY',PRICING_SOURCE='bval')
    else:
        historicalprices[MasterDF['Ticker'][i]] = mgr[MasterDF['Ticker'][i]+' MTGE'].get_historical('PX_Last',start=start_Date,end=end_Date,period='DAILY',PRICING_SOURCE='bval')

### test1 = (sids.get_historical('PX_Last',start = '5/14/2018', end = '5/16/2018', period = "DAILY", PRICING_SOURCE = 'bval'))

constructdf = pd.DataFrame(index=historicalprices.values()[0].index)
columnheaders = []
for i in range(len(historicalprices)):
    constructdf = pd.merge(constructdf,historicalprices.values()[i],how='left',left_index=True,right_index=True)
    columnheaders.append(historicalprices.keys()[i])

constructdf.columns = columnheaders
finaldf = constructdf.T

finaldf.iloc[:,[0]] = np.where(finaldf.iloc[:,[0]].isnull(),"Paid Off",finaldf.iloc[:,[0]])
finaldf.iloc[:,[1]] = np.where(finaldf.iloc[:,[1]].isnull(),"Paid Off",finaldf.iloc[:,[1]])

filtereddf = finaldf[finaldf[finaldf.columns[1]]!= "Paid Off"]
filtereddf = filtereddf[filtereddf[filtereddf.columns[0]]!= "Paid Off"]
filtereddf[filtereddf.columns[1]] = filtereddf[filtereddf.columns[1]].astype('float')
filtereddf[filtereddf.columns[0]] = filtereddf[filtereddf.columns[0]].astype('float')
filtereddf['DoD Change'] = filtereddf[filtereddf.columns[1]]-filtereddf[filtereddf.columns[0]]
filtereddf['Abs Change'] = abs(filtereddf['DoD Change'])
filtereddf['Abs Change'] = filtereddf['Abs Change'].astype('float')
filtereddf['Abs % Change']= abs(((filtereddf[filtereddf.columns[1]]-filtereddf[filtereddf.columns[0]])/filtereddf[filtereddf.columns[0]])*100)
filtereddf = filtereddf.round(2)
filtereddf.columns.values[0] = filtereddf.columns[0].strftime('%m/%d/%Y')
filtereddf.columns.values[1] = filtereddf.columns[1].strftime('%m/%d/%Y')

top10movers = filtereddf.nlargest(10,'Abs Change')
top10movers = top10movers.sort_values(by=['Abs Change'],ascending = False)
toppercentmovers = filtereddf[filtereddf['Abs % Change']>=1]
toppercentmovers = toppercentmovers.sort_values(by='Abs % Change',ascending = False)

excelwriter = pd.ExcelWriter('Biggest Trade Movers - '+dt.datetime.today().strftime('%m-%d-%Y')+'.xlsx')
toppercentmovers.to_excel(excelwriter,'Top Movers by Percent')
top10movers.to_excel(excelwriter,'Top Ten Movers')
filtereddf.to_excel(excelwriter,'CUSIP Universe Change')
finaldf.to_excel(excelwriter,'Security Data')
workbook = excelwriter.book
worksheetone = excelwriter.sheets['Top Movers by Percent']
worksheettwo = excelwriter.sheets['Top Ten Movers']
worksheetthree = excelwriter.sheets['CUSIP Universe Change']
worksheetfour = excelwriter.sheets['Security Data']
worksheetone.set_column('A:F',18)
worksheettwo.set_column('A:F',18)
worksheetthree.set_column('A:F',18)
worksheetfour.set_column('A:C',18)
excelwriter.save()
excelwriter.close()

olMailItem = 0x0
outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
#print(dir(outlook))
outlook = win32com.client.Dispatch("Outlook.Application")
#print(dir(outlook))
newmail = outlook.CreateItem(olMailItem)
newmail.To = "hsantana@tolisadvisors.com; rportnoy@tolisadvisors.com; spuliafico@tolisadvisors.com; ebanks@tolisadvisors.com; tpangia@tolisadvisors.com; sparker@tolisadvisors.com; jrosato@tolisadvisors.com; bilany@tolisadvisors.com"
newmail.Subject = "Daily Biggest Movers"
newmail.HTMLBody = "Hi Team, <br><br>Please see below for today's biggest price movers by absolute price change and absolute percentage change in price. Attached is the excel report.<br><br> Top Ten Biggest Movers by Price <br><br>" + top10movers.to_html() + "<br><br> Top Movers by Percent <br><br>" + toppercentmovers.to_html() 
newmail.Attachments.Add(Source = r'C:\Users\santana\Documents\Bloomberg Scrapes' +'\Biggest Trade Movers - '+dt.datetime.today().strftime('%m-%d-%Y')+'.xlsx')
newmail.Send()


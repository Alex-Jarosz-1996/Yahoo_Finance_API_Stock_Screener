from YahooFinanceScreenerClass import *
import concurrent.futures
import os
import pandas as pd
import shutil

def getAllTickersFromWatchlist():
    currentPath = str(os.path.abspath(os.getcwd()))
    desiredPath = currentPath.replace('src', 'US')
    df = pd.read_excel(f'{desiredPath}\watchlist.xlsx')
    return df['Desired Stocks'].tolist()

def saveDataToYF_Class(ticker):
    try:
        return Stock(str(ticker))
    except Exception as e:
        print(f'Error with stock ticker: {ticker}')
        print(e)

def allYF_DataClassList():
    allTickers = [*getAllTickersFromWatchlist()]
    with concurrent.futures.ProcessPoolExecutor() as executor:
        return list(executor.map(saveDataToYF_Class, allTickers))

def allYF_DataListFormat():
    dataList = allYF_DataClassList()
    dfLst = [[st.ticker, st.price, st.marketCap, 
              st.numShares, st.yearlyLowPrice, st.yearlyHighPrice, 
              st.fiftyDayMA, st.twoHundredDayMA, st.acquirersMultiple,
              st.currentRatio, st.enterpriseValue, st.eps,
              st.evToEBITDA, st.evToOperatingCashFlow, st.evToRev,
              st.peRatioTrail, st.peRatioForward, st.priceToSales,
              st.priceToBook, st.dividendYield, st.dividendRate,
              st.exDivDate, st.lastDivVal, st.payoutRatio,
              st.bookValue, st.bookValPerShare, st.cash,
              st.cashPerShare, st.cashToMarketCap, st.cashToDebt,
              st.debt, st.debtToMarketCap, st.debtToEquityRatio,
              st.quickRatio, st.returnOnAssets, st.returnOnEquity,
              st.totalAssets, st.ebitda, st.ebitdaMargins, 
              st.ebitdaPerShare, st.earningsGrowth, st.grossMargins,
              st.grossProfit, st.grossProfitPerShare, st.netIncome,
              st.netIncomePerShare, st.operatingMargin, st.profitMargin,
              st.revenue, st.revenueGrowth, st.revenuePerShare,
              st.fcf, st.fcfToMarketCap, st.fcfPerShare,
              st.ocf, st.ocfToRevenueRatio, st.ocfToMarketCap,
              st.ocfPerShare, st.fcfToEV, st.ocfToEV] for st in dataList]
    dfHeaderNames = ['Ticker', 'Price', 'Market_Cap',
                 'Num_Shares', 'Yearly_Low_Price', 'Yearly_High_Price',
                 'Fifty_Day_MA', 'Two_Hundred_Day_MA', 'Acquirers_Multiple',
                 'Current_Ratio', 'EV', 'EPS',
                 'EV_To_EBITDA', 'EV_To_OCF', 'EV_To_Rev', 
                 'PE_Rat_Trail', 'PE_Rat_Forward', 'Price_Sales',
                 'Price_Book', 'Div_Yield', 'Div_Rate',
                 'Ex_Div_Date', 'Last_Div_Value', 'Payout_Ratio',
                 'Book_Value', 'Book_Val_Per_Share', 'Cash',
                 'Cash_Per_Share', 'Cash_To_Market_Cap', 'Cash_To_Debt',
                 'Debt', 'Debt_To_Market_Cap', 'Debt_To_Equity_Ratio',
                 'Quick_Ratio', 'Return_On_Assets', 'Return_On_Equity',
                 'Total_Assets', 'EBITDA', 'EBITDA_Margins', 
                 'EBITDA_Per_Share', 'Earnings_Growth', 'Gross_Margins',
                 'Gross_Profit', 'Gross_Profit_Per_Share', 'Net_Income',
                 'Net_Income_Per_Share', 'Operating_Margin', 'Profit_Margin',
                 'Revenue', 'Revenue_Growth', 'Revenue_Per_Share',
                 'FCF', 'FCF_To_Market_Cap', 'FCF_Per_Share',
                 'OCF', 'OCF_To_Revenue_Ratio', 'OCF_To_Market_Cap',
                 'OCF_Per_Share', 'FCF_To_EV', 'OCF_To_EV']
    print('Created stock data data frame.')
    return pd.DataFrame(dfLst, columns=dfHeaderNames)

def saveDFtoExcel(df, excelFileName):
    removeSingleExcelFile(excelFileName)
    currentPath = str(os.path.abspath(os.getcwd()))
    desiredPath = currentPath.replace('src', 'US')
    df.to_excel(f'{desiredPath}/{excelFileName}.xlsx', index=False)

def removeCache():
    currentPath = str(os.path.abspath(os.getcwd()))
    cachePath = f'{currentPath}\__pycache__'
    shutil.rmtree(cachePath)
    print('Removed Cache.')

def removeSingleExcelFile(excelFileName): 
    currentPath = str(os.path.abspath(os.getcwd()))
    desiredPath = currentPath.replace('src', 'US')
    filePath = f'{desiredPath}\{excelFileName}.xlsx'
    if os.path.exists(filePath):
        os.remove(filePath)
        print(f'Deleted old {excelFileName}.') 

def main():
    excelFileName = 'all_stock_data'
    removeSingleExcelFile(excelFileName)
    df = allYF_DataListFormat()
    saveDFtoExcel(df, excelFileName)
    removeCache()

    

#imports
import yfinance as yf
import os
import pandas as pd
import pywintypes
import xlwings as xw

#create directory to store data
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

# point to excel file and obtain ticker
wb = xw.Book('Stocks.xlsm')
main_sht = wb.sheets('MAIN')
ticker = main_sht.range('C2').value

#GET STOCK DATA
def get_stock_info(ticker):
    try:
        print("--LOADING 0%-- Retrieving Stock info")
        #Create csv file with data
        pd.DataFrame.from_dict(yf.Ticker(ticker).info, orient="index").to_csv(f'{directory}/{ticker}_info.csv', header=False)
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_info.csv', header=None, names=['parameter', 'value'], index_col=0)
        #Point to info sheet and insert data
        sht = wb.sheets('info')
        sht.clear()
        sht.range('A1').value = csv_data
    except:
        print("Something went wrong retrieving stock info")

#GET STOCK HISTORIC DATA
def get_stock_history(ticker):
    try:
        print("--LOADING 15%-- Retrieving Stock historical data")
        #Create csv file with data
        yf.download(ticker, period="max").to_csv(f'{directory}/{ticker}_history.csv', header=True) 
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_history.csv', index_col=0)
        #Point to historic sheet and insert data
        sht = wb.sheets('historic')
        sht.clear()
        sht.range('A1').value = csv_data
    except:
        print("Something went wrong retrieving stock historical data")

#GET STOCK FINANCIALS DATA
def get_stock_financials(ticker):
    try:
        print("--LOADING 30%-- Retrieving Stock financials")
        #Create csv file with data
        pd.DataFrame(data=yf.Ticker(ticker).financials).to_csv(f'{directory}/{ticker}_financials.csv', header=True)
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_financials.csv', index_col=0)
        #Point to finanicals sheet and insert data
        sht = wb.sheets('financials')
        sht.clear()
        sht.range('A1').value = csv_data
    except:
        print("Something went wrong retrieving stock finanicals")

#GET STOCK CASH FLOW DATA
def get_stock_cashflow(ticker):
    try:
        print("--LOADING 45%-- Retrieving Stock cashflow")
        #Create csv file with data
        pd.DataFrame(data=yf.Ticker(ticker).cashflow).to_csv(f'{directory}/{ticker}_cash_flow.csv', header=True)    
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_cash_flow.csv', index_col=0)
        #Point to cash flow sheet and insert data
        sht = wb.sheets('cashflow')
        sht.clear()
        sht.range('A1').value = csv_data  
    except:
        print("Something went wrong retrieving stock cash flow")

#GET STOCK BALANCE SHEET DATA
def get_stock_balance_sheet(ticker):
    try:
        print("--LOADING 60%-- Retrieving Stock balance sheet")
        #Create csv file with data
        pd.DataFrame(data=yf.Ticker(ticker).balance_sheet).to_csv(f'{directory}/{ticker}_balance_sheet.csv', header=True)      
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_balance_sheet.csv', index_col=0)
        #Point to balance sheet and insert data
        sht = wb.sheets('balance')
        sht.clear()
        sht.range('A1').value = csv_data  
    except:
        print("Something went wrong retrieving stock balance sheet")

#GET STOCK EARNINGS DATA
def get_stock_earnings(ticker):
    try:
        print("--LOADING 75%-- Retrieving Stock earnings")
        #Create csv file with data
        pd.DataFrame(data=yf.Ticker(ticker).earnings).to_csv(f'{directory}/{ticker}_earnings.csv', header=True)      
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_earnings.csv', index_col=0)
        #Point to earnings sheet and insert data
        sht = wb.sheets('earnings')
        sht.clear()
        sht.range('A1').value = csv_data  
    except:
        print("Something went wrong retrieving stock earnings")

#GET STOCK DIVIDENDS DATA
def get_stock_dividends(ticker):
    try:
        print("--LOADING 90%-- Retrieving Stock dividends")
        #Create csv file with data
        pd.DataFrame(data=yf.Ticker(ticker).dividends).to_csv(f'{directory}/{ticker}_dividends.csv', header=True)      
        #Read csv file into formatted table
        csv_data = pd.read_csv(f'{directory}/{ticker}_dividends.csv', index_col=0)
        #Point to earnings sheet and insert data
        sht = wb.sheets('dividends')
        sht.clear()
        sht.range('A1').value = csv_data
    except:
        print("Something went wrong retrieving stock dividends")

if __name__ == '__main__':
    get_stock_info(ticker)
    get_stock_history(ticker)
    get_stock_financials(ticker)
    get_stock_cashflow(ticker)
    get_stock_balance_sheet(ticker)
    get_stock_earnings(ticker)
    get_stock_dividends(ticker)
    print("--LOADING 100%-- Stock data loaded")
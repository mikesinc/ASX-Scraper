#imports
import yfinance as yf
import csv
import os
import pandas as pd

#Example list of tickers - Eventually replace with list that will read csv file containing tickers from user's excel doc.
tickers = ["WPL.AX", "CBA.AX", "STO.AX", "CSL.AX"]
# tickers = ["WPL.AX"]

#create directory to store data
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#GET STOCK GENERAL INFO
#Create generator (doesn't store massive array in memory)
def gen_ticker_info():
    for ticker in tickers:
        yield ticker, yf.Ticker(ticker).info
#Create function to store ALL stock (in ticker list) general info in single csv file.
def get_stock_info():
    # Write stock data to csv file
    csv_dict = {}
    csv_dict_columns = list(next(gen_ticker_info())[1].keys())
    for stock in gen_ticker_info():
        csv_dict[stock[0]] = list(stock[1].values())
        print(f'{stock[0]} data extract complete')
    pd.DataFrame.from_dict(data=csv_dict, orient='index', columns=csv_dict_columns).to_csv(f'{directory}/all_stocks_info.csv', header=True)      
    print("Stocks info updated!")

#GET STOCK HISTORIC DATA
def get_stock_history():
    yf.download(tickers=tickers, period="5y", group_by="ticker").to_csv(f'{directory}/all_stocks_history.csv', header=True) 
    print("Stocks history updated!")

#GET STOCK FINANCIALS DATA
#Create function to store SINGLE stock financials info in csv file.
def get_stock_financials(ticker):
     # Write stock data to csv file
    pd.DataFrame(data=yf.Ticker(ticker).financials).to_csv(f'{directory}/stock_financials.csv', header=True)      
    print("Stock financials updated!")

#GET STOCK CASH FLOW DATA
#Create function to store SINGLE stock cash flow info in csv file.
def get_stock_cashflow(ticker):
     # Write stock data to csv file
    pd.DataFrame(data=yf.Ticker(ticker).cashflow).to_csv(f'{directory}/stock_cash_flow.csv', header=True)      
    print("Stock cash flows updated!")

#GET STOCK BALANCE SHEET DATA
#Create function to store SINGLE stock balance sheet info in csv file.
def get_stock_balance_sheet(ticker):
     # Write stock data to csv file
    pd.DataFrame(data=yf.Ticker(ticker).balance_sheet).to_csv(f'{directory}/stock_balance_sheet.csv', header=True)      
    print("Stock balance sheet updated!")

#GET STOCK EARNINGS DATA
#Create function to store SINGLE stock balance sheet info in csv file.
def get_stock_earnings(ticker):
     # Write stock data to csv file
    pd.DataFrame(data=yf.Ticker(ticker).earnings).to_csv(f'{directory}/stock_earnings.csv', header=True)      
    print("Stock earnings updated!")

#GET STOCK RECOMMENDATIONS DATA
#Create function to store SINGLE stock balance sheet info in csv file.
def get_stock_dividends(ticker):
     # Write stock data to csv file
    pd.DataFrame(data=yf.Ticker(ticker).dividends).to_csv(f'{directory}/stock_dividends.csv', header=True)      
    print("Stock dividends updated!")


if __name__ == '__main__':
    get_stock_info()
    get_stock_history()
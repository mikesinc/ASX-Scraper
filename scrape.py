#Import libraries
import pandas as pd
import yfinance as yf
from bs4 import BeautifulSoup
import requests
import os
import time
import pywintypes
import xlwings as xw
from datetime import date
from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import config

#Set directory
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#Excel details
wb = xw.Book('Stocks.xlsm')
main_sht = wb.sheets('MAIN')
ticker = main_sht.range('C2').value

#Store URLs
urls = {'financials': f"https://www.morningstar.com.au/Stocks/CompanyHistoricals/{ticker}", "info": f"https://www.morningstar.com.au/Stocks/NewsAndQuotes/{ticker}", "fair": f"https://www.morningstar.com.au/Stocks/Overview/{ticker}", "balance": f"https://www.morningstar.com.au/Stocks/BalanceSheet/{ticker}"}

#clean strings (remove spaces)
def remove_multiple_spaces(string):
    if type(string)==str:
        return string.replace(",","").replace(" ","").split("\n")
    return string

#Get info
def get_info():
    #Wipe excel sheet
    sht = wb.sheets('stock summary')
    sht.clear()
    #Go to info page
    try:
        r = requests.get(urls['info'])
        data = BeautifulSoup(r.content, features="html5lib")

        info_dict = {}
        for item in data.find('h1', attrs={'class': 'N_QHeader_b'}):
            info_dict['name'] = data.find('label').text
        info_dict['price'] = data.find('span', attrs={'style': 'font-size: 1.5em; color: #333333; float: left;'}).text
        info_dict['time'] = data.find('span', attrs={'class': 'N_QText'}).text
        raw_day_change = remove_multiple_spaces(data.find('div', attrs={'style': 'float: left; font-size: 1.5em;'}).text.replace("|",""))
        info_dict['day change cents'] = ' '.join(raw_day_change).split()[0]
        info_dict['day change %'] = ' '.join(raw_day_change).split()[1]

        raw_table = data.find('table')
        headers = [header.text for header in raw_table.findAll('h3')]
        values = [remove_multiple_spaces(item.text)[0] for item in raw_table.findAll('span')]
        summary_dict = dict(zip(headers,values))
        
        info_dict['open'] = summary_dict['Open Price']
        info_dict['day low'] = summary_dict['Day Range'].split("-")[0]
        info_dict['day high'] = summary_dict['Day Range'].split("-")[1]
        info_dict['52 week low'] = summary_dict['Day Range'].split("-")[0]
        info_dict['52 week high'] = summary_dict['Day Range'].split("-")[1]
        info_dict['market cap'] = summary_dict['Market Cap']
        info_dict['volume'] = summary_dict["Volume - 30 Day Avg"]
        info_dict['sector'] = summary_dict["GICS Sector"]
        info_dict['industry'] = summary_dict["GICS Industry"]
        
        # create dataframe from dictionaries
        pd.DataFrame.from_dict(data=info_dict, orient='index').to_csv(f'{directory}/{ticker}_info.csv', header=False)
        # Read csv files into formatted table
        csv_financials_data = pd.read_csv(f'{directory}/{ticker}_info.csv', header=None, names=['parameter', 'value'], index_col=0)
        # Point to financials sheet and insert data
        sht.range('A1').value = csv_financials_data
    except:
        print("failed loading info data")

#Get historical data
def get_stock_history():
    try:
        sheet_list = ['historical price - max', 'historical price - 5y', 'historical price - 1y', 'historical price - 6mo', 'historical price - 3mo', 'historical price - 1mo', 'historical price - 10d', 'historical price - 1d']
        sheet_dict = {'historical price - max': '1d', 'historical price - 5y': '1d', 'historical price - 1y': '1d', 'historical price - 6mo': '1d', 'historical price - 3mo': '1d', 'historical price - 1mo': '60m', 'historical price - 10d': '30m', 'historical price - 1d': '1m'}
        for sheet in sheet_list:
            #Create csv file with data             
            yf.download(ticker+'.AX', period=sheet.split("- ")[1], interval=sheet_dict[sheet]).to_csv(f'{directory}/{ticker}_{sheet.split("- ")[1]}.csv', header=True) 
    except:
        print("Something went wrong retrieving stock historical data")

#Get financials
def get_financials():
    #Go to financials page
    try:
        driver.get(urls['financials'])
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('loginButton'))
        driver.find_element_by_id('loginButton').click()
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('loginFormNew'))
        driver.find_element_by_xpath("//input[contains(@type, 'email')]").send_keys(config.ms_username)
        driver.find_element_by_xpath("//input[contains(@type, 'password')]").send_keys(config.ms_password)
        driver.find_element_by_xpath("//input[contains(@type, 'password')]").send_keys(Keys.ENTER)
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('pershare'))

        #Beautiful Soup the tables
        html = driver.page_source
        soup = BeautifulSoup(html,'html.parser')
        per_share, = pd.read_html(str(soup.findAll('table')[0]))
        historical_financials, = pd.read_html(str(soup.findAll('table')[1]))
        # cash_flow, = pd.read_html(str(soup.findAll('table')[2]))

        #get other info values
        driver.get(urls['fair'])
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('eqFVval'))
        fair_value = driver.find_element_by_id('eqFVval')
        uncertainty_rating = driver.find_element_by_id('eqURval')
        previous_close = driver.find_element_by_class_name('textB2')
        economic_moat = driver.find_element_by_id('eqEMval')
        sht = wb.sheets('stock summary')

        #Dump fair value
        sht.range('A20').clear()
        sht.range('B20').clear()
        sht.range('A20').value = "Morningstar fair value"
        sht.range('B20').value = fair_value.text

        #Dump uncertainty_rating
        sht.range('A21').clear()
        sht.range('B21').clear()
        sht.range('A21').value = "Morningstar Uncertainty Rating"
        sht.range('B21').value = uncertainty_rating.text

        #Dump previous_close value
        sht.range('A22').clear()
        sht.range('B22').clear()
        sht.range('A22').value = "Previous close"
        sht.range('B22').value = previous_close.text

        #Dump economic_moat value
        sht.range('A23').clear()
        sht.range('B23').clear()
        sht.range('A23').value = "Economic Moat"
        sht.range('B23').value = economic_moat.text

        #Get Balance sheet
        driver.get(urls['balance'])
        WebDriverWait(driver, 30).until(lambda d: d.find_elements_by_class_name('table1 dividendhisttable'))
        alert = driver.switch_to_alert()
        alert.dismiss()
        #Beautiful Soup the table
        html = driver.page_source
        soup = BeautifulSoup(html,'html.parser')
        print(soup)
        cash_flow, = pd.read_html(str(soup.findAll('table')[2]))

        per_share.to_csv(f'{directory}/{ticker}_statistics.csv', header=False)
        historical_financials.to_csv(f'{directory}/{ticker}_financials.csv', header=False)
        cash_flow.to_csv(f'{directory}/{ticker}_cashflow.csv', header=False)
        
        # Read csv files into formatted table
        csv_per_share = pd.read_csv(f'{directory}/{ticker}_statistics.csv', header=None, names=['parameter'] + list(range(date.today().year-10,date.today().year)), index_col=0)
        csv_historical_financials = pd.read_csv(f'{directory}/{ticker}_financials.csv', header=None, names=['parameter'] + list(range(date.today().year-10,date.today().year)), index_col=0)
        csv_cash_flow = pd.read_csv(f'{directory}/{ticker}_cashflow.csv', header=None, names=['parameter'] + list(range(date.today().year-10,date.today().year)), index_col=0)
        # Point to per share statistics sheet and insert data
        #Wipe excel sheet
        sht = wb.sheets('per share statistics')
        sht.clear()
        sht.range('A1').value = csv_per_share
        # Point to historical financials sheet and insert data
        #Wipe excel sheet
        sht = wb.sheets('historical financials')
        sht.clear()
        sht.range('A1').value = csv_historical_financials
        # Point to cash flow sheet and insert data
        #Wipe excel sheet
        sht = wb.sheets('cash flow')
        sht.clear()
        sht.range('A1').value = csv_cash_flow
    except:
        print("Something went wrong retrieving stock finanical data")

if __name__ == '__main__':
    # Initialise Driver
    options = Options()
    # options.headless = True
    options.set_preference('dom.webnotifications.enabled', False)
    # fp = webdriver.FirefoxProfile()
    # fp.set_preference("dom.webnotifications.enabled", False) 
    binary = FirefoxBinary(r'C:\Program Files\Mozilla Firefox\firefox.exe')
    driver = webdriver.Firefox(firefox_binary=binary, options=options)
    # Extract Data
    print("--LOADING 0%-- Retrieving info")
    get_info()
    print("--LOADING 40%-- Retrieving historical prices")
    get_stock_history()
    print("--LOADING 60%-- Retrieving financials")
    get_financials()
    # Stop Driver
    driver.close()
    print("--LOADING 100%-- Stock data loaded")

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
from selenium.webdriver import ActionChains
import config
import random

#Set directory
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#Excel details
wb = xw.Book('Stocks.xlsm')
main_sht = wb.sheets('MAIN')
ticker = main_sht.range('C2').value

#Store URLs
urls = {'financials': f"https://www2.commsec.com.au/", "info": f"https://www.morningstar.com.au/Stocks/NewsAndQuotes/{ticker}"}

#clean strings
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
        info_dict['previous close'] = summary_dict['Prev Close']
        info_dict['day low'] = summary_dict['Day Range'].split("-")[0]
        info_dict['day high'] = summary_dict['Day Range'].split("-")[1]
        info_dict['52 week high'] = summary_dict['52-Week Range'].split("-")[0]
        info_dict['52 week low'] = summary_dict['52-Week Range'].split("-")[1]
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
    try:
        #Go to CommSec Stock Summary
        driver.get(urls['financials'])
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('loginPanel'))
        #Login
        driver.find_element_by_xpath("//input[contains(@id, 'ctl00_cpContent_txtLogin')]").send_keys(config.cs_id)
        driver.find_element_by_xpath("//input[contains(@id, 'ctl00_cpContent_txtLogin')]").send_keys(Keys.TAB)
        driver.find_element_by_xpath("//input[contains(@id, 'ctl00_cpContent_txtPassword')]").send_keys(config.cs_password)
        driver.find_element_by_xpath("//input[contains(@id, 'ctl00_cpContent_txtPassword')]").send_keys(Keys.ENTER)
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_class_name('Level1Wrapper'))
        
        #Scrape Financials Summary page
        driver.get(urls['financials'] + f'quotes/financials?stockCode={ticker}&exchangeCode=ASX#/financials/company')
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_tag_name('table'))
        html = driver.page_source
        soup = BeautifulSoup(html,  features="html5lib")
        historical_statistics, = pd.read_html(str(soup.find('table')))
        historical_statistics.to_csv(f'{directory}/{ticker}_statistics.csv', header=True, index=False)

        #Scrape Historical Financials page
        driver.get(urls['financials'] + f'quotes/financials?stockCode={ticker}&exchangeCode=ASX#/financials/historical')
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_tag_name('table'))
        html = driver.page_source
        soup = BeautifulSoup(html,  features="html5lib")
        historical_financials, = pd.read_html(str(soup.find('table')))
        historical_financials.to_csv(f'{directory}/{ticker}_financials.csv', header=True, index=False)

        #Scrape Balance Sheet page
        driver.get(urls['financials'] + f'quotes/financials?stockCode={ticker}&exchangeCode=ASX#/financials/balance')
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_tag_name('table'))
        html = driver.page_source
        soup = BeautifulSoup(html,  features="html5lib")
        balance_sheet, = pd.read_html(str(soup.find('table')))
        balance_sheet.to_csv(f'{directory}/{ticker}_balance.csv', header=True, index=False)

        #Scrape Fair Value
        driver.get(urls['financials'] + f'quotes/recommendations?stockCode={ticker}&exchangeCode=ASX#/recommendations/premium')
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('recommendations-container'))
        driver.get(urls['financials'] + f'quotes/recommendations?stockCode={ticker}&exchangeCode=ASX#/recommendations/premium')
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_class_name("mstar-premium-overview-contain"))
        fair_value = driver.find_elements_by_tag_name("strong")[2].text
        fair_uncertainty = driver.find_elements_by_tag_name("strong")[3].text
        
        sht = wb.sheets('stock summary')
        sht.range('A17').clear()
        sht.range('B17').clear()
        sht.range('A18').clear()
        sht.range('B18').clear()
        sht.range('A19').clear()
        sht.range('B19').clear()
        sht.range('A17').value = "fair value"
        sht.range('B17').value = fair_value
        sht.range('A18').value = "fair uncertainty"
        sht.range('B18').value = fair_uncertainty

        #Read csv files into excel 
        csv_historical_statistics = pd.read_csv(f'{directory}/{ticker}_statistics.csv', header=0, index_col=0)
        csv_historical_financials = pd.read_csv(f'{directory}/{ticker}_financials.csv', header=0, index_col=0)
        csv_balance_sheet = pd.read_csv(f'{directory}/{ticker}_balance.csv', header=0, index_col=0)
        # Point to statistics sheet and insert data
        sht = wb.sheets('historical statistics')
        sht.clear()
        sht.range('A1').value = csv_historical_statistics
        # Point to financials sheet and insert data
        sht = wb.sheets('historical financials')
        sht.clear()
        sht.range('A1').value = csv_historical_financials
        # Point to financials sheet and insert data
        sht = wb.sheets('balance sheet')
        sht.clear()
        sht.range('A1').value = csv_balance_sheet

    except:
        print("Something went wrong retrieving financial data")

if __name__ == '__main__':
    # Initialise Driver
    options = Options()
    options.headless = True
    # geckodriver = os.getcwd()
    # driver = webdriver.Firefox(executable_path=geckodriver, options=options)
    binary = FirefoxBinary(r'C:\Program Files\Mozilla Firefox\firefox.exe')
    driver = webdriver.Firefox(firefox_binary=binary, options=options)
    # except:
    #     print("Failed to intialise driver")
    
    # Scrape Functions
    print("--LOADING 0%-- Retrieving info")
    get_info()
    print("--LOADING 40%-- Retrieving historical prices")
    get_stock_history()
    print("--LOADING 60%-- Retrieving financials")
    get_financials()
    # Stop Driver
    driver.close()
    print("--LOADING 100%-- Stock data loaded")
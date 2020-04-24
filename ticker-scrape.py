#Import libraries
import pandas as pd
import yfinance as yf
from bs4 import BeautifulSoup
import requests
import os
import pywintypes
import xlwings as xw
import glob
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys

#Set directory
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#Excel details
try:
    wb = xw.Book('ASXScraper.xlsm')
    main_sht = wb.sheets(str(sys.argv[1]))
    ticker = main_sht.range('C4').value
except:
    print('Could not find the worksheet!')
    print('Press any key to close...')
    input()
    quit()

#clean strings
def remove_multiple_spaces(string):
    if type(string)==str:
        return string.replace(",","").replace(" ","").split("\n")
    return string

#Get historical data
def get_stock_history():
    sheet_list = ['historical price - max', 'historical price - 5y', 'historical price - 1y', 'historical price - 6mo', 'historical price - 3mo', 'historical price - 1mo', 'historical price - 5d', 'historical price - 1d']
    sheet_dict = {'historical price - max': '1d', 'historical price - 5y': '1d', 'historical price - 1y': '1d', 'historical price - 6mo': '1d', 'historical price - 3mo': '1d', 'historical price - 1mo': '1d', 'historical price - 5d': '60m', 'historical price - 1d': '1m'}    
    try:      
        #Dump trend data for the periods as csv files to be called from VBA
        for sheet in sheet_list:   
            yf.download(ticker+'.AX', period=sheet.split("- ")[1], interval=sheet_dict[sheet]).to_csv(f'{directory}/trend data/{ticker}_{sheet.split("- ")[1]}.csv', header=True)
    except:
        print(f"Something went wrong retrieving {ticker} historical data")
        print('Press any key to close...')
        input()
        quit()

#Get financials
def get_info():
    try:
        driver.get(f"https://www.morningstar.com.au/Stocks/NewsAndQuotes/{ticker}")
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.TAG_NAME, "table")))
    except:
        print(f"Something went wrong loading the {ticker} morning star page!")
        driver.quit()
        print('Press any key to close...')
        input()
        quit()
    try:
        html = driver.page_source
        soup = BeautifulSoup(html,'html.parser')
        info_set = {}
        for item in soup.findAll('td'):
            if item.find('span') and item.find('h3'):
                info_set[item.find('h3').text] = item.find('span').text
        info_set['name'] = driver.find_element_by_xpath("//h1[contains(@class, 'N_QHeader_b')]/label").text
        info_set['value'] = driver.find_element_by_xpath("//div[contains(@class, 'N_QPriceLeft')]/div[1]/span/span[2]").text
        info_set['day change cents'] = driver.find_element_by_xpath("//div[contains(@class, 'N_QPriceLeft')]/div[2]/div[2]/span[1]").text
        info_set['day change percent'] = driver.find_element_by_xpath("//div[contains(@class, 'N_QPriceLeft')]/div[2]/div[2]/span[3]").text
        info_set['as of text'] = driver.find_element_by_class_name("N_QText").text
    except:
        print(f"Something went wrong scraping {ticker} info!")
        driver.quit()
        print('Press any key to close...')
        input()
        quit()
    try:
        main_sht.range("C5").value = info_set['value']
        main_sht.range("B2").value = info_set['name']
        main_sht.range("B27").value = info_set['name']
        main_sht.range("B3").value = info_set['GICS Sector']
        main_sht.range("H2").value = info_set['Market Cap'].replace("\xa0", "").replace(",", "").replace("M", "000000").replace("B", "000000000")
        main_sht.range("L2").value = info_set['52-Week Range'].split("-")[0]
        main_sht.range("L3").value = info_set['52-Week Range'].split("-")[1]
        main_sht.range("F2").value = info_set['Prev Close']
        main_sht.range("F3").value = info_set['Open Price']
        main_sht.range("J2").value = info_set['Day Range'].split("-")[1]
        main_sht.range("J3").value = info_set['Day Range'].split("-")[0]
        main_sht.range("H3").value = info_set['Volume - 30 Day Avg']
        main_sht.range("C7").value = info_set['as of text']
        main_sht.range("B182").value = info_set['day change cents']
        main_sht.range("B183").value = info_set['day change percent']
    except:
        print(f"Something went wrong importing {ticker} info into excel..")
        driver.quit()
        print('Press any key to close...')
        input()
        quit()

if __name__ == '__main__':
    # initialise driver
    print("....LOADING: Starting Web Driver...")
    options = Options()
    options.headless = True
    geckodriver = os.getcwd() + '\\geckodriver.exe'
    driver = webdriver.Firefox(executable_path=geckodriver, options=options)
    extension_dir = 'C:\\Users\\Michael\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\0mhu8le1.default-release\\extensions\\'
    extensions = [
        'https-everywhere@eff.org.xpi',
        'uBlock0@raymondhill.net.xpi',
        ]
    for extension in extensions:
        driver.install_addon(extension_dir + extension, temporary=True)
    # Scrape stock daily info
    print("....LOADING: Retrieving stock info data")
    get_info()
    #stop driver
    driver.quit()
    print("....LOADING: Retrieving historical stock price trend data")
    get_stock_history()
    print("....LOADING DONE!")
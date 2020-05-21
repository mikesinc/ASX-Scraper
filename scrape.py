import csv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import bcrypt
import config
import time
import os
from bs4 import BeautifulSoup

#Set directory
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

def get_dividends(password):
    try:
        driver.get(f"https://www.morningstar.com.au/Stocks/CompanyHistoricals/{tickers[0]}")
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('loginButton'))
        driver.find_element_by_id('loginButton').click()
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('loginFormNew'))
        driver.find_element_by_xpath("//input[contains(@type, 'email')]").send_keys(config.ms_username)
        time.sleep(2)
        driver.find_element_by_xpath("//input[contains(@type, 'password')]").send_keys(password)
            
        recaptcha = driver.find_element_by_xpath("//div[contains(@id, 'reCaptchaContainer')]/div/div/iframe")
        driver.switch_to.frame(recaptcha)
        iframe_checkbox = driver.find_element_by_xpath("//span[contains(@id, 'recaptcha-anchor')]")
        iframe_checkbox.click()
        driver.switch_to.default_content()
        recaptcha_images = driver.find_elements(By.TAG_NAME, "iframe")
        driver.switch_to.frame(recaptcha_images[3])
        time.sleep(1)
        buster_button = driver.find_element_by_xpath("//button[contains(@id, 'solver-button')]")
        buster_button.click()
        time.sleep(10) 
        driver.switch_to.default_content()
        driver.find_element_by_xpath("//input[contains(@id, 'LoginSubmit')]").click()
        time.sleep(5)
        print("hacker boi, im in") 
    except: 
        print('failed login attempt')  
        driver.quit()
        quit()  

    for ticker in tickers:
        try:
            driver.get(f"https://www.morningstar.com.au/Stocks/CompanyHistoricals/{ticker}")
            WebDriverWait(driver, 5).until(lambda d: d.find_element_by_id('pershare'))

            years = ['']

            for year in driver.find_elements_by_xpath("//table[contains(@id, 'pershare')]/tbody/tr[1]/td"):
                if len(year.text) > 0:
                    years.append("20"+year.text.split("/")[1])

            html = driver.page_source
            soup = BeautifulSoup(html, features="html5lib")
            PS_df, = pd.read_html(str(soup.findAll('table')[0]), header=None, skiprows=1)
            HF_df, = pd.read_html(str(soup.findAll('table')[1]), header=None, skiprows=1)
            cash_df, = pd.read_html(str(soup.findAll('table')[2]), header=None, skiprows=1)

            PS_df.columns = years
            HF_df.columns = years
            cash_df.columns = years
            
            PS_df[''] = ticker + " " + PS_df[''].astype(str)
            HF_df[''] = ticker + " " + HF_df[''].astype(str)
            cash_df[''] = ticker + " " + cash_df[''].astype(str)

            ticker_df = pd.concat([
                PS_df,
                HF_df,
                cash_df
                ], 
                sort=False)

            ticker_details.append(ticker_df)

            print(f'{ticker} data scraped')
        except:
            print(f"failed loading {ticker} page")
    
    pd.concat(ticker_details, sort=True).to_csv(f'{directory}/stock_patch.csv', header=True, columns=['', '2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019'], index=False)

    database_csv = pd.read_csv(f'{directory}/stock_data_refined.csv')
    patch_csv = pd.read_csv(f'{directory}/stock_patch.csv')

    pd.concat([patch_csv, database_csv], sort=False).to_csv(f'{directory}/stock_full_database.csv', header=True, index=False)

if __name__ == '__main__':
    ms_password = input("password: ")
    if bcrypt.checkpw(ms_password.encode('utf8'), config.ms_password):
        # ticker list
        tickers = []
        for value in pd.read_csv(f'{os.getcwd()}/ASXListedCompanies.csv', usecols=[1], header=-1).values:
            tickers.append(value[0])

        ticker_details = []

        options = Options()
        # options.headless = True
        geckodriver = os.getcwd() + '\\geckodriver.exe'
        driver = webdriver.Firefox(executable_path=geckodriver, options=options)
        extension_dir = os.getcwd() + '\\driver_extensions\\'
        extensions = [
            'https-everywhere@eff.org.xpi',
            'uBlock0@raymondhill.net.xpi',
            '{e58d3966-3d76-4cd9-8552-1582fbc800c1}.xpi'
            ]
        for extension in extensions:
            driver.install_addon(extension_dir + extension, temporary=True)
        # Scrape history and fair value
        get_dividends(ms_password)
        # Stop Driver
        driver.quit()
    else:
        print("invalid password")
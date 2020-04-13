#Import libraries
import pandas as pd
import os
import time
import pywintypes
import xlwings as xw
from datetime import date
from selenium import webdriver
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait

#Set directory
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#Excel details
wb = xw.Book('Stocks.xlsm')
main_sht = wb.sheets('MAIN')
ticker = main_sht.range('C2').value.split(".")[0]

#Store URLs
urls = {'financials': f"https://financials.morningstar.com/income-statement/is.html?t={ticker}&region=aus&culture=en-US&platform=sal", 'balance': f"https://financials.morningstar.com/balance-sheet/bs.html?t={ticker}&region=aus&culture=en-US&platform=sal"}

#clean strings (remove spaces)
def remove_multiple_spaces(string):
    if type(string)==str:
        return string.replace(",","").replace(" ","").split("\n")
    return string

def get_financials():
    #Go to financials page
    try:
        driver.get(urls['financials'])
        WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('Li1'))
        time.sleep(0.5)
    except:
        print("failed to load page")

    #Set period 10 years
    clicker = driver.find_element_by_id('Li1')
    clicker.click()
    dropdown_list = clicker.find_elements_by_tag_name('ul')
    year_selector = dropdown_list[0].find_elements_by_tag_name('li')[1]
    year_selector.click()
    WebDriverWait(driver, 30).until(lambda d: d.find_element_by_id('Y_10'))
    time.sleep(0.5)

    #Create financials dictionary
    try:
        financials_dict = {}
        financials_dict['revenue'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i1')]").text)
        financials_dict['cost of revenue'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i6')]").text)
        financials_dict['gross profit'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i10')]").text)
        financials_dict['sales, general and administrative'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i12')]").text)
        financials_dict['other operating expenses'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i29')]").text)
        financials_dict['total operating expenses'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttg3')]").text)
        financials_dict['operating income'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i30')]").text)
        financials_dict['interest expense'] = remove_multiple_spaces(driver.find_elements_by_xpath("//div[contains(@id, 'data_i51')]")[1].text)
        financials_dict['other income (expense)'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i52')]").text)
        financials_dict['income before taxes'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i60')]").text)
        financials_dict['provision for income taxes'] = remove_multiple_spaces(driver.find_elements_by_xpath("//div[contains(@id, 'data_i61')]")[1].text)
        financials_dict['other income'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i69')]").text)
        financials_dict['net income from continuing operation'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i70')]").text)
        financials_dict['other'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i74')]").text)
        financials_dict['net income'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i80')]").text)
        financials_dict['net income available to common shareholders'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i82')]").text)
        financials_dict['eps basic'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i83')]").text)
        financials_dict['eps diluted'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i84')]").text)
        financials_dict['weight av share outstanding basic'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i85')]").text)
        financials_dict['weight av share outstanding diluted'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i86')]").text)
        financials_dict['ebitda'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i90')]").text)

        #create dataframe from dictionaries
        pd.DataFrame.from_dict(data=financials_dict, orient='index').to_csv(f'{directory}/{ticker}_financials.csv', header=False)
        # Read csv files into formatted table
        csv_financials_data = pd.read_csv(f'{directory}/{ticker}_financials.csv', header=None, names=['parameter'] + list(range(date.today().year-10,date.today().year)), index_col=0)
        # Point to financials sheet and insert data
        sht = wb.sheets('financials')
        sht.clear()
        sht.range('A1').value = csv_financials_data
    except:
        print("failed loading financial data")

def get_balance_sheet():
    #Go to balance sheet page
    try:
        driver.get(urls['balance'])
        WebDriverWait(driver, 30).until(lambda d: d.find_elements_by_class_name('rf_crow'))
        time.sleep(0.5)
    except:
        print("failed to load page")

    #Create balance sheet dictionary
    try:
        balance_dict = {}
        balance_dict['cash and cash equivalents'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i1')]").text)
        balance_dict['short-term investments'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i2')]").text)
        balance_dict['total cash'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttgg1')]").text)
        balance_dict['receivables'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i3')]").text)
        balance_dict['inventories'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i4')]").text)
        balance_dict['deferred income taxes'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i5')]").text)
        balance_dict['prepaid expenses'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i6')]").text)
        balance_dict['other current assets'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i8')]").text)
        balance_dict['total current assets'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttg1')]").text)
        balance_dict['gross property, plant and equipment'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i9')]").text)
        balance_dict['accumulation depreciation'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i10')]").text)
        balance_dict['net property, plant and equipment'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttgg2')]").text)
        balance_dict['equity and other investments'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i11')]").text)
        balance_dict['non-current deferred income taxes'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i14')]").text)
        balance_dict['other long-term assets'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i17')]").text)
        balance_dict['total non-current assets'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttg2')]").text)
        balance_dict['total assets'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_tts1')]").text)
        balance_dict['short-term debt'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i41')]").text)
        balance_dict['capital leases'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i42')]").text)
        balance_dict['accounts payable'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i43')]").text)
        balance_dict['liabilities deferred income taxes'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i44')]").text)
        balance_dict['deferred revenues'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i47')]").text)
        balance_dict['other current liabilities'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i49')]").text)
        balance_dict['total current liabilities'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttgg5')]").text)
        balance_dict['long-term debt'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i50')]").text)
        balance_dict['non-current capital leases'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i51')]").text)
        balance_dict['non-current deferred taxes liabilities'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i52')]").text)
        balance_dict['non-current deferred revenues'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i54')]").text)
        balance_dict['pensions and other benefits'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i55')]").text)
        balance_dict['minority interest'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i56')]").text)
        balance_dict['other long-term liabilities'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i58')]").text)
        balance_dict['total non-current liabilities'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttgg6')]").text)
        balance_dict['total liabilities'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttg5')]").text)
        balance_dict['common stock'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i82')]").text)
        balance_dict['other equity'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i83')]").text)
        balance_dict['retained earnings'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i85')]").text)
        balance_dict['accumulated other comprehensive incomes'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_i89')]").text)
        balance_dict['total stockholders equity'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_ttg8')]").text)
        balance_dict['total liabilties and stockholders equity'] = remove_multiple_spaces(driver.find_element_by_xpath("//div[contains(@id, 'data_tts2')]").text)

        #create dataframe from dictionaries
        pd.DataFrame.from_dict(data=balance_dict, orient='index').to_csv(f'{directory}/{ticker}_balance_sheet.csv', header=False)
        # Read csv files into formatted table
        csv_balance_data = pd.read_csv(f'{directory}/{ticker}_balance_sheet.csv', header=None, names=['parameter'] + list(range(date.today().year-10,date.today().year)), index_col=0)
        # Point to balance sheet and insert data
        sht = wb.sheets('balance')
        sht.clear()
        sht.range('A1').value = csv_balance_data
    except:
        print("failed loading balance sheet data")

if __name__ == '__main__':
    # Initialise Driver
    options = Options()
    options.headless = True
    binary = FirefoxBinary(r'C:\Program Files\Mozilla Firefox\firefox.exe')
    driver = webdriver.Firefox(firefox_binary=binary, options=options)
    # Extract Data
    print("--LOADING 0%-- Retrieving financials")
    get_financials()
    print("--LOADING 50%-- Retrieving balance sheet")
    get_balance_sheet()
    # Stop Driver
    driver.close()
    print("--LOADING 100%-- Stock data loaded")

#Import libraries
import pywintypes
import xlwings as xw
import pandas as pd
import os

#Set directory
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#Criteria details
try:
    wb = xw.Book('ASXScraper.xlsm')
    criterion = {
        "Market cap": {
            'max': wb.sheets("HOME").range("B26").value,
            'min': wb.sheets("HOME").range("B27").value,
            'screened': {}
            },
        "Dividends (¢)": {
            'max': wb.sheets("HOME").range("D26").value,
            'min': wb.sheets("HOME").range("D27").value,
            'screened': {}
            },
        "Dividend Yield (%)": {
            'max': wb.sheets("HOME").range("D30").value,
            'min': wb.sheets("HOME").range("D31").value,
            'screened': {}
            },
        "eps basic": {
            'max': wb.sheets("HOME").range("H26").value,
            'min': wb.sheets("HOME").range("H27").value,
            'screened': {}
            },
        "Average annual P/E ratio (%)": {
            'max': wb.sheets("HOME").range("J26").value,
            'min': wb.sheets("HOME").range("J27").value,
            'screened': {}
            },
        "Shares outstanding": {
            'max': wb.sheets("HOME").range("B30").value,
            'min': wb.sheets("HOME").range("B31").value,
            'screened': {}
            },
        "total cash": {
            'max': wb.sheets("HOME").range("F30").value,
            'min': wb.sheets("HOME").range("F31").value,
            'screened': {}
            }, 
        "Net profit margin (%)": {
            'max': wb.sheets("HOME").range("J30").value,
            'min': wb.sheets("HOME").range("J31").value,
            'screened': {}
            }, 
        "ebit": {
            'max': wb.sheets("HOME").range("B34").value,
            'min': wb.sheets("HOME").range("B35").value,
            'screened': {}
            }, 
        "ebitda": {
            'max': wb.sheets("HOME").range("D34").value,
            'min': wb.sheets("HOME").range("D35").value,
            'screened': {}
            }, 
        "Revenue": {
            'max': wb.sheets("HOME").range("F34").value,
            'min': wb.sheets("HOME").range("F35").value,
            'screened': {}
            }, 
        "net income": {
            'max': wb.sheets("HOME").range("H34").value,
            'min': wb.sheets("HOME").range("H35").value,
            'screened': {}
            }, 
        "long-term debt": {
            'max': wb.sheets("HOME").range("J34").value,
            'min': wb.sheets("HOME").range("J35").value,
            'screened': {}
            }  
        }
    functional_criterion = {
        "EV": {
            'max': wb.sheets("HOME").range("H30").value,
            'min': wb.sheets("HOME").range("H31").value,
            'screened': {}
            }, 
        "CY": {
            'max': wb.sheets("HOME").range("F26").value,
            'min': wb.sheets("HOME").range("F27").value,
            'screened': {}
            }   
    }
except:
    print('Could not find the worksheet!')
    print('Press any key to close...')
    input()
    quit()

#Get CSV database
database = pd.read_csv(f"{directory}/database.csv", low_memory=False)

#Screening function
def check(value, Max, Min, Dict, ticker):
    #Check if criteria entered
    try:
        if Max or Min:
            #Check if has min AND max criteria
            if Max and Min:
            #Add tickers within criteria range
                if float(value) >= Min and float(value) <= Max:
                    Dict[ticker] = float(value)
            #Check if has max criteria
            elif Max:
                #Add tickers that are below max criteria
                if float(value) <= Max:
                    Dict[ticker] = float(value)
            #Check if has min criteria
            elif Min:
                #Add tickers that are above min criteria
                if float(value) >= Min:
                    Dict[ticker] = float(value)
            #If there are no bounds set, add the ticker
        else:
            Dict[ticker] = value
    except:
        print("failed screening")

def initial_screen():
    #Go through database and screen tickers based on MOST RECENT year
    for row in database.get_values():
        if not row[len(row)-1] == "—":
            for Property in criterion.keys():
                if Property in row[0]:
                    check(row[len(row)-1], criterion[Property]['max'], criterion[Property]['min'], criterion[Property]['screened'], row[0].split(" ")[0])         

def functional_screen():
    #Define empty property object to store desired properties
    prop_dict = {}
    #Collects already-screened property values from inital screening dictionaries
    try:
        for Property in ['Market cap', 'long-term debt', 'total cash', 'ebitda']:
            prop_dict[Property] = criterion[Property]['screened']
    except:
        print("something went wrong creating property dictionary")
                
    #Loop through tickers (that are already narrowed-down by initial screening if performed) and
    #screen further based on EV and CY values
    for ticker in screened_tickers:
        #Reset values
        CY_value = None
        EV_value = None
        market_cap = None
        longterm_debt = None
        ebitda = None
        total_cash = None

        #Extract values for ticker
        try:
            market_cap = prop_dict["Market cap"][ticker]
            longterm_debt = prop_dict["long-term debt"][ticker]
            ebitda = prop_dict["ebitda"][ticker]
            total_cash = prop_dict["total cash"][ticker]
        except:
            print("something went wrong retreving function inputs")

        #If values exist, calculate EV and CY
        try:
            if market_cap and longterm_debt and ebitda:
                if not float(ebitda) == 0 and not (float(longterm_debt) + float(market_cap) == 0):
                    CY_value = (float(ebitda) * 100) / (float(longterm_debt) + float(market_cap))
                    if total_cash:
                        EV_value = (float(longterm_debt) + float(market_cap) - float(total_cash)) / float(ebitda)
        except:
            print("something went wrong calculating EV and / or CY values")
         
        #Check if values for CY and EV meet criteria and add them to output list, if CY and EV cannot be calculated, the ticker is not added to the list.
        try:
            if CY_value and EV_value:
                check(CY_value, functional_criterion['CY']['max'], functional_criterion['CY']['min'], functional_criterion['CY']['screened'], ticker)
                check(EV_value, functional_criterion['EV']['max'], functional_criterion['EV']['min'], functional_criterion['EV']['screened'], ticker)
        except:
            print("something went wrong screening based on EV and CY criteria!")

def copy_to_excel(tickers):
    try:
        #Clear cells
        wb.sheets("HOME").range("B41:B2000").api.Delete()
        #From row 41, insert ticker list into excel
        row = 41
        for ticker in tickers:
            wb.sheets("HOME").range("B" + str(row)).value = ticker
            row += 1
    except:
        print("someting went wrong importing to excel!")

if __name__ == '__main__':
    #Run initial screen based on simple database data (immediately avaiable)
    initial_screen()
    #initial screened list
    try:
        screened_list = []
        for Property in criterion.keys():
            screened_list.append(criterion[Property]['screened'])
        screened_tickers = sorted(list(set(screened_list[0]).intersection(*screened_list)))
    except:
        print("something went wrong creating initial screened ticker list")

    #If EV or CY criteria is specified, run the extra screening function (performs functions on database values)
    if functional_criterion['CY']['max'] or functional_criterion['CY']['min'] or functional_criterion['EV']['max'] or functional_criterion['EV']['min']:
        functional_screen()
        #final screened list
        try:
            for Property in functional_criterion.keys():
                if len(functional_criterion[Property]['screened']) > 0:
                    screened_list.append(functional_criterion[Property]['screened'])
            screened_tickers = sorted(list(set(screened_list[0]).intersection(*screened_list)))
        except:
            "something went wrong creating the final screened ticker list"
    
    #Dump to excel
    copy_to_excel(screened_tickers)
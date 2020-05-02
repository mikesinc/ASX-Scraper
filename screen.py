#Import libraries
import pywintypes
import xlwings as xw
import pandas as pd
import os
import sys
from statistics import mean

#Set directory
directory = os.path.abspath(os.path.join(os.path.dirname( __file__ ), '..', 'data'))

#Criteria details
try:
    wb = xw.Book('ASXScraper.xlsm')
    criterion = {
        "Market cap": {
            'max': wb.sheets("HOME").range("B23").value,
            'min': wb.sheets("HOME").range("B24").value,
            'screened': {}
            },
        "Dividends (¢)": {
            'max': wb.sheets("HOME").range("D23").value,
            'min': wb.sheets("HOME").range("D24").value,
            'screened': {}
            },
        "Dividend Yield (%)": {
            'max': wb.sheets("HOME").range("D27").value,
            'min': wb.sheets("HOME").range("D28").value,
            'screened': {}
            },
        "eps basic": {
            'max': wb.sheets("HOME").range("H23").value,
            'min': wb.sheets("HOME").range("H24").value,
            'screened': {}
            },
        "Average annual P/E ratio (%)": {
            'max': wb.sheets("HOME").range("J23").value,
            'min': wb.sheets("HOME").range("J24").value,
            'screened': {}
            },
        "Shares outstanding": {
            'max': wb.sheets("HOME").range("B27").value,
            'min': wb.sheets("HOME").range("B28").value,
            'screened': {}
            },
        "total cash": {
            'max': wb.sheets("HOME").range("F27").value,
            'min': wb.sheets("HOME").range("F28").value,
            'screened': {}
            }, 
        "Net profit margin (%)": {
            'max': wb.sheets("HOME").range("J27").value,
            'min': wb.sheets("HOME").range("J28").value,
            'screened': {}
            }, 
        "ebit": {
            'max': wb.sheets("HOME").range("B31").value,
            'min': wb.sheets("HOME").range("B32").value,
            'screened': {}
            }, 
        "ebitda": {
            'max': wb.sheets("HOME").range("D31").value,
            'min': wb.sheets("HOME").range("D32").value,
            'screened': {}
            }, 
        "net income": {
            'max': wb.sheets("HOME").range("H31").value,
            'min': wb.sheets("HOME").range("H32").value,
            'screened': {}
            }, 
        "long-term debt": {
            'max': wb.sheets("HOME").range("F31").value,
            'min': wb.sheets("HOME").range("F32").value,
            'screened': {}
            }
        }
    functional_criterion = {
        "EV": {
            'max': wb.sheets("HOME").range("H27").value,
            'min': wb.sheets("HOME").range("H28").value,
            'screened': {}
            }, 
        "CY": {
            'max': wb.sheets("HOME").range("F23").value,
            'min': wb.sheets("HOME").range("F24").value,
            'screened': {}
            }   
    }
except:
    print('Could not find the worksheet!')
    print('Press enter to close...')
    input()
    sys.exit()

#Get CSV database
database = pd.read_csv(f"{directory}/database.csv")

#Screening function
def check(value, Max, Min, Dict, ticker):
    #Check if criteria entered
    try:
        if Max or Min:
            #Check if has min AND max criteria
            if not value == "—":
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
        print("failed checking criteria")

def basic_screen(searchtype):
    #Go through database and screen tickers
    ticker_props = {}
    for ticker in tickers:
        ticker_props[ticker] = []
    for row in database.to_numpy():
        ticker = row[0].split(" ")[0]
        for Property in criterion.keys():
            if Property in row[0] and not "from" in row[0] and not "available" in row[0]: #ensure "from" and "available" aren't in the word to remove extra net income results
                ticker_props[ticker].append(Property) 
                if searchtype == '2': #if search type is average of all years since 2010 (search type == 2), get average value
                    values = []
                    for value in row[range(1, len(row))]:
                        if not value == "—":
                            values.append(float(value))
                    if len(values):
                        value = mean(values)
                else:
                    value = row[len(row)-1]
                check(value, criterion[Property]['max'], criterion[Property]['min'], criterion[Property]['screened'], row[0].split(" ")[0])      
    
    #If the ticker did not have the property listed in the database, do a check anyway on a default "—" value, so that it is still added to the results
    #given there were no criteria entered for that specific property.
    for ticker in ticker_props:
        for Property in criterion.keys():
            if not Property in ticker_props[ticker]:
                check("—", criterion[Property]['max'], criterion[Property]['min'], criterion[Property]['screened'], ticker)      

def screen(searchtype):
    try: #Collect formula inputs by running initial screening for the required inputs
        basic_screen(str(sys.argv[1]))
        for Property in criterion.keys():
            screened_lists.append(criterion[Property]['screened'])
        screened_tickers = sorted(list(set(screened_lists[0]).intersection(*screened_lists)))
    except:
        print("something went wrong with basic screen")

    try:
        market_cap_values = criterion['Market cap']['screened']
        longterm_debt_values = criterion['long-term debt']['screened']
        total_cash_values = criterion['total cash']['screened']
        ebitda_values = criterion['ebitda']['screened']
    except:
        print("something went wrong collecting formula inputs.")
                
    #Loop through tickers (that are already narrowed-down by initial screening if performed) and
    #screen further based on EV and CY values
    for ticker in screened_tickers: #Reset values
        CY_value = None
        EV_value = None
        market_cap = None
        longterm_debt = None
        ebitda = None
        total_cash = None

        try: #Extract values for ticker
            if ticker in market_cap_values and not market_cap_values[ticker] == "—":
                market_cap = float(market_cap_values[ticker])
            if ticker in longterm_debt_values and not longterm_debt_values[ticker] == "—":
                longterm_debt = float(longterm_debt_values[ticker])
            if ticker in total_cash_values and not total_cash_values[ticker] == "—":
                total_cash = float(total_cash_values[ticker])
            if ticker in ebitda_values and not ebitda_values[ticker] == "—":
                ebitda = float(ebitda_values[ticker])
        except:
            print("something went wrong retreving function inputs.")
  
        try: #If values exist, calculate EV and CY
            if market_cap and longterm_debt and ebitda:
                if not ebitda == 0 and not (longterm_debt + market_cap == 0):
                    CY_value = (ebitda * 100) / (longterm_debt + market_cap)
                    check(CY_value, functional_criterion['CY']['max'], functional_criterion['CY']['min'], functional_criterion['CY']['screened'], ticker)
                    if total_cash:
                        EV_value = (longterm_debt + market_cap - total_cash) / ebitda
                        check(EV_value, functional_criterion['EV']['max'], functional_criterion['EV']['min'], functional_criterion['EV']['screened'], ticker)
                    else:
                        check("—", functional_criterion['EV']['max'], functional_criterion['EV']['min'], functional_criterion['EV']['screened'], ticker)
                        check("—", functional_criterion['CY']['max'], functional_criterion['CY']['min'], functional_criterion['CY']['screened'], ticker)
            else: 
                #If they do not exist (i.e. cannot perform the calculations), do a search on default "—", so that it is still added to the results 
                #given there were no criteria entered for that specific property.)
                check("—", functional_criterion['CY']['max'], functional_criterion['CY']['min'], functional_criterion['CY']['screened'], ticker)
                check("—", functional_criterion['EV']['max'], functional_criterion['EV']['min'], functional_criterion['EV']['screened'], ticker)
        except:
            print("something went wrong calculating EV and / or CY values")

    try:
        for Property in functional_criterion.keys():
            screened_lists.append(functional_criterion[Property]['screened'])
        return sorted(list(set(screened_lists[0]).intersection(*screened_lists)))
    except:
        print("something went wrong with functional screen.")

def sector_screen(tickers):
    sector_filtered_tickers = tickers #Default to list of all tickers after main screening
    if not wb.sheets("HOME").range("J31").value == "All": #If sector criteria defined, further screen for this
        sector_filtered_tickers = [] #Redefine as empty list to be filled with tickers matching sector
        sector_criteria = wb.sheets("HOME").range("J31").value
        for ticker in tickers:
            if listings[ticker][1] == sector_criteria:
                sector_filtered_tickers.append(ticker)
    screened_lists.append(sector_filtered_tickers)
    return sorted(list(set(screened_lists[0]).intersection(*screened_lists)))

def copy_to_excel(tickers):
    try:
     #Clear cells
        wb.sheets("HOME").range("A42:S2500").clear_contents()
        wb.sheets("HOME").range("B40").value = f"{len(tickers)} listings matched your critera"
        
        rows = []
        for ticker in tickers:
            row = [ticker, listings[ticker][0], "", listings[ticker][1]]
            for Property in functional_criterion.keys():
                row.append(functional_criterion[Property]['screened'][ticker])
            for Property in criterion.keys():
                row.append(criterion[Property]['screened'][ticker])
            rows.append(row)
        
        data = pd.DataFrame(rows)
        wb.sheets("HOME").range("A42").value = data 
    except:
        print("someting went wrong importing to excel!")

if __name__ == '__main__':
    try:
        #Get ticker list fromm database (not all ASX listed tickers have data in database)
        tickers = []
        for row in database.to_numpy():
            tickers.append(row[0].split(" ")[0])
        tickers = sorted(list(set(tickers)))
            #Go through full ASX listing and extract name and sector of those listings which have data in database
        listings = {}
        for name, ticker, sector in pd.read_csv(f'{directory}/ASXListedCompanies.csv', usecols=[0, 1, 2], header=None).values:
            if ticker in tickers:
                listings[ticker] = name, sector
        print("ticker list unfiltered: ", len(tickers), len(listings.keys()))
    except:
        print("error generating ticker listings")

    try:
        screened_lists = [] #List that stores lists of tickers that match criteria for each property
        screened_tickers = [] #List of tickers that match all criteria (intersect of lists in screen_lists)

        #Run the screening, parameter is SEARCH TYPE (recent year or average all years)
        screened_tickers = screen(str(sys.argv[1]))

        #Screen for sector
        screened_tickers = sector_screen(screened_tickers)

        #Export to excel
        copy_to_excel(screened_tickers)
        print("screen complete!")
        
    except:
        print("something went wrong with the screen function")
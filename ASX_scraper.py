from alpha_vantage.timeseries import TimeSeries
import csv
import os
import config

#Example list of tickers
myTickers = ["WPL", "CBA", "CSL", "STO"]

#create directory to store data
directory = os.getcwd() + "/data"
if not os.path.exists(directory):
    os.makedirs(directory)

#Loop to create CSV files for each Stock in ticker list
ts = TimeSeries(key=config.api_key, output_format='csv')
for ticker in myTickers:
    #outputsize is "compact" - giving last 100 entires, can set to full for ALL data
    data_csvreader, meta = ts.get_daily(symbol=f'ASX:{ticker}', outputsize="compact")
    with open(f'{directory}/{ticker}.csv', 'w') as write_csvfile:
        writer = csv.writer(write_csvfile, dialect='excel')
        for row in data_csvreader:
            writer.writerow(row)
    print(f"{ticker} data extract complete")
print("Stocks updated!")
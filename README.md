# ASX Scraper & Analysis Tool
Python web scraper for ASX stock analysis. Built into excel (VBA code not included here).
Users can use this tool to perform their own predictions / projection of stock performance based on the scraped data of stocks.

## scrape.py 
Script built to web-scrape historical financial data for over 2000 ASX listings from Morningstar and save into a (local) database (.csv format for ease of import into excel).

## ticker-scrape.py
Script built to retrieve daily data for desired stock listings on ASX. Built into excel such that several stocks can be added and monitored, to make informed desicions about investments.

## screen.py
Screening tool script built to sort/filter over 2000 ASX listings to match criteria specified by the user in the excel document. Users can search based on recent year or average since data collection begun.

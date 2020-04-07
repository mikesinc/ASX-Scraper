import yh_scraper

try:
    search_ticker = input('Please enter a ticker (include ".AX" suffix for ASX listed companies): ')
    yh_scraper.get_stock_financials(search_ticker)
    yh_scraper.get_stock_cashflow(search_ticker)
    yh_scraper.get_stock_balance_sheet(search_ticker)
    yh_scraper.get_stock_earnings(search_ticker)
    yh_scraper.get_stock_dividends(search_ticker)
except:
    print("Please enter a valid ticker")
from enum import Enum
import time
import pandas as pd
from yahoofinancials import YahooFinancials
import xlwings as xw

class Column(Enum):
    long_name = 1
    ticker = 2
    current_price = 5
    currency = 6
    conversion_rate = 7
    open_price = 8
    daily_low = 9
    daily_high = 10
    yearly_low = 11
    yearly_high = 12
    fifty_day_moving_avg = 13
    twohundred_day_moving_avg = 14
    payout_ratio = 19
    exdividend_date = 20
    yield_rel = 21
    dividend_rate = 22

def timestamp():
    t = time.localtime()
    timestamp = time.strftime("%b-%d-%Y_%H:%M:%S", t)
    return timestamp

def clear_content_in_excel(sheet, start_row, last_row):
    """Clear the old contents in Excel."""
    if last_row > start_row:
        print(f"Clear Contents from row {start_row} to {last_row}")
        for data in Column:
            if data != Column.ticker:
                sheet.range((start_row, data.value), (last_row, data.value)).options(expand="down").clear_contents()

def get_conversion_rate(ticker_currency, target_currency):
    """Calculate the conversion rate between ticker currency & target currency."""
    if target_currency == "TICKER CURRENCY":
        print(f"Display values in {ticker_currency}")
        conversion_rate = 1
    else:
        conversion_rate = YahooFinancials(f"{ticker_currency}{target_currency}=X").get_current_price()
        print(f"Conversion Rate from {ticker_currency} to {target_currency}: {conversion_rate}")
    return conversion_rate

def pull_stock_data(sheet, tickers, target_currency):
    """Pull financial data for a list of tickers."""
    if not tickers:
        return pd.DataFrame()

    df = pd.DataFrame()
    for ticker in tickers:
        print(f"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print(f"Pulling financial data for: {ticker} ...")
        data = YahooFinancials(ticker)
        open_price = data.get_open_price()

        if open_price is None:
            print(f"Ticker: {ticker} not found on Yahoo Finance. Please check")
            df = pd.concat([df, pd.DataFrame([pd.Series(dtype=str)])], ignore_index=True)
        else:
            try:
                long_name = data.get_stock_quote_type_data().get(ticker, {}).get("longName")
                yield_rel = data.get_summary_data().get(ticker, {}).get("yield")

                ticker_currency = data.get_currency()
                conversion_rate = get_conversion_rate(ticker_currency, target_currency)

                new_row = {
                    "ticker": ticker,
                    "currency": ticker_currency,
                    "long_name": long_name,
                    "conversion_rate": conversion_rate,
                    "yield_rel": yield_rel,
                    "exdividend_date": data.get_exdividend_date(),
                    "payout_ratio": data.get_payout_ratio(),
                    "open_price": open_price,
                    "current_price": data.get_current_price(),
                    "daily_low": data.get_daily_low(),
                    "daily_high": data.get_daily_high(),
                    "yearly_low": data.get_yearly_low(),
                    "yearly_high": data.get_yearly_high(),
                    "fifty_day_moving_avg": data.get_50day_moving_avg(),
                    "twohundred_day_moving_avg": data.get_200day_moving_avg(),
                    "dividend_rate": data.get_dividend_rate(),
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                print(f"Successfully pulled financial data for: {ticker}")

            except Exception as e:
                print(f"Error pulling data for {ticker}: {str(e)}")
                df = pd.concat([df, pd.DataFrame([pd.Series(dtype=str)])], ignore_index=True)
    return df

def write_data_to_excel(sheet, df, start_row):
    if not df.empty:
        print(f"~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
        print(f"Writing data to Excel...")
        options = dict(index=False, header=False)
        for data in Column:
            if data != Column.ticker:
                sheet.range(start_row, data.value).options(**options).value = df[data.name]

def main():
    print(
        """
    ==============================
    Dividend & Portfolio Overview
    ==============================
    """
    )

    print(f"Please wait. The program is running ...")
    wb = xw.Book.caller()
    sheet = wb.sheets("Portfolio")
    target_currency = sheet.range("TARGET_CURRENCY").value
    start_row = sheet.range("TICKER").row + 1
    last_row = sheet.range(sheet.cells.last_cell.row, Column.ticker.value).end("up").row
    sheet.range("TIMESTAMP").value = timestamp()
    tickers = sheet.range(start_row, Column.ticker.value).options(expand="down", numbers=str).value

    clear_content_in_excel(sheet, start_row, last_row)
    df = pull_stock_data(sheet, tickers, target_currency)
    write_data_to_excel(sheet, df, start_row)

    print(f"Program ran successfully!")

if __name__ == "__main__":
    main()
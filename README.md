# Quantitative Momentum Strategy: S&P 500 Stock Analyzer

This Python script implements a quantitative momentum strategy to analyze S&P 500 stocks. It calculates the one-year price return for each stock, determines the number of shares to buy based on a given portfolio size, and outputs the results to a formatted Excel file.

## Features
- **Historical Data Retrieval**: Fetches one year of historical stock price data using the Yahoo Finance API.
- **One-Year Price Return Calculation**: Computes the percentage change in stock prices over the past year.
- **Equal Weighting Strategy**: Distributes a portfolio size equally among selected stocks.
- **Excel Export**: Saves the recommended trades, including stock tickers, prices, one-year price returns, and the number of shares to buy, to a formatted Excel file.

## Requirements
- Python 3.x
- pandas
- numpy
- yfinance
- scipy
- xlsxwriter

You can install the required packages using pip:

```bash
pip install pandas numpy yfinance scipy xlsxwriter
```
# How to Use
Prepare the Stock List: Ensure you have a CSV file containing a column named 'Ticker' with the list of S&P 500 stock tickers. Update the filepath variable in the script with the path to your CSV file.

Run the Script: Execute the script using Python. It will prompt you to enter the value of your portfolio.

View Results: The script will generate an Excel file with columns for Ticker, Price, One-Year Price Return, and Number of Shares to Buy based on the momentum strategy.

Check Output: The Excel file will be saved at the specified output_filepath with custom formatting.

# Notes
Ensure the CSV file has the correct format with a 'Ticker' column.
The script includes basic error handling for unavailable or delisted stocks.
Modify the file paths in the script according to your local file structure.
The strategy assumes equal weighting across selected stocks based on portfolio size.

# File Paths
- Input File: /path/to/your/sp_500_stocks.csv
- Output File: /path/to/save/recommended_trades.xlsx

import pandas as pd
import numpy as np
import yfinance as yf
import math
from scipy import stats

# Filepath for the list of S&P 500 stocks
filepath = '/Users/ahamedrashid/Desktop/Projects/Trading Algorithm Project/Strategy Test/Follow Supplement/sp_500_stocks.csv'
list_stocks = pd.read_csv(filepath)['Ticker']  # Ensure you're reading the 'Ticker' column

# Initialize the DataFrame with appropriate columns
new_columns = ['Ticker', 'Price', 'One-Year Price Return', 'Number of Shares to Buy']
rows = []

# Iterate over each stock symbol in the list
for symbol in list_stocks:
    stock = yf.Ticker(symbol)
    
    # Download historical data for the past year
    stock_data = yf.download(symbol, start='2023-08-18', end='2024-08-18')
    
    # Check if stock_data is empty (handle delisted or unavailable tickers)
    if stock_data.empty:
        print(f"No data found for {symbol}, skipping.")
        continue
    
    # Get the starting and ending prices
    starting_price = stock_data['Adj Close'].iloc[0]
    ending_price = stock_data['Adj Close'].iloc[-1]
    
    # Calculate percentage change over the year
    percentage_change = ((ending_price - starting_price) / starting_price) * 100
    
    # Get the latest stock price
    price = stock.info.get('previousClose')
    
    if price is None:
        print(f"Failed to retrieve the price for {symbol}, skipping.")
        continue
    
    rows.append([symbol, price, percentage_change, 0])  # Append the data to rows

# Convert the rows into a DataFrame
new_dataframe = pd.DataFrame(rows, columns=new_columns)

# Calculate the position size
portfolio_size = float(input("Enter the value of your portfolio: "))
position_size = portfolio_size / len(new_dataframe)

# Calculate the number of shares to buy
new_dataframe['Number of Shares to Buy'] = new_dataframe.apply(
    lambda row: math.floor(position_size / row['Price']), axis=1
)

# Define the output path
output_filepath = '/Users/ahamedrashid/Desktop/Projects/Trading Algorithm Project/Strategy Test/Quantitative Momentum Strategy/recommended_trades.xlsx'

# Save the DataFrame to an Excel file with formatting
with pd.ExcelWriter(output_filepath, engine='xlsxwriter') as writer:
    new_dataframe.to_excel(writer, sheet_name='Recommended Trades', index=False)

    # Excel Formatting
    workbook = writer.book
    worksheet = writer.sheets['Recommended Trades']
    
    background_color = '#0a0a23'
    font_color = '#ffffff'

    string_format = workbook.add_format({
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })

    dollar_format = workbook.add_format({
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })
    percent_template = writer.book.add_format({
                'num_format':'0.0%',
                'font_color': font_color,
                'bg_color': background_color,
                'border': 1
            })
    integer_format = workbook.add_format({
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })

    # Column Formatting
    column_formats = {
        'A': ['Ticker', string_format],
        'B': ['Price', dollar_format],
        'C': ['One-Year Price Return', percent_template],
        'D': ['Number of Shares to Buy', integer_format]
    }

    for column, (header, fmt) in column_formats.items():
        worksheet.set_column(f'{column}:{column}', 20, fmt)
        worksheet.write(f'{column}1', header, string_format)

print(f"Excel file saved to {output_filepath}")

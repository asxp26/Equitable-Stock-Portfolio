import pandas as pd
import requests
import xlsxwriter
import math

# Import IEX Cloud API token from a confidential file
from starters.confidential import IEX_CLOUD_API_TOKEN

# Load S&P 500 tickers from the sp_500_stocks CSV file
stocks = pd.read_csv('starters/sp_500_stocks.csv')

# Define columns for the final output dataframe
my_columns = [ 'Ticker', 'Stock Price', 'Market Cap', 'Number of Shares to Buy' ]

# Function to divide a list into chunks of a specified size
def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

# Split the list of stocks into chunks of 100 (IEX Cloud API limitation)
symbol_groups = list(chunks(stocks['Ticker'], 100))
all_symbol_strings = []

# Create a list of strings containing the stock symbols
for group in symbol_groups:
    symbol_string = ','.join(group)
    all_symbol_strings.append(symbol_string)

# Create an empty dataframe with specified columns
final_dataframe = pd.DataFrame(columns = my_columns)

# Retrieve data for each stock using batch API calls
for symbol_string in all_symbol_strings:
    batch_api_call_url = f'https://api.iex.cloud/v1/stock/market/batch?symbols={symbol_string}&types=quote&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    
    # Extract relevant data for each stock and append to the dataframe
    for symbol in data:
        new_data = pd.Series(
            [
                symbol,
                data[symbol]['quote']['latestPrice'],
                data[symbol]['quote']['marketCap'],
                'N/A'
            ],
            index = my_columns
        )

        final_dataframe = pd.concat([final_dataframe, new_data.to_frame().T], ignore_index=True)

# Get user input for their portfolio value
while True:
    portfolio_size = input('\nEnter portfolio value: ')

    try:
        val = float(portfolio_size)
        break 
    except ValueError:
        print("\nThat's not a number! \nPlease try again...")

print()

# Calculate the position size for each stock
position_size = val/len(final_dataframe.index)

# Calculate the number of shares to buy for each stock
for i in range(0, len(final_dataframe.index)):
    final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Stock Price'])

# Create an Excel writer object and write the final dataframe to an Excel file
writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer, 'Recommended Trades', index = False)

# Define formatting for Excel cells
background_color = '#16558F'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color' : font_color,
        'bg_color' : background_color,
        'border' : 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format' : '$0.00',
        'font_color' : font_color,
        'bg_color' : background_color,
        'border' : 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format' : '0',
        'font_color' : font_color,
        'bg_color' : background_color,
        'border' : 1
    }
)

# Define formatting for each column in the Excel sheet
column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C': ['Market Cap', dollar_format], 
    'D': ['Number of Shares to Buy', integer_format]
}

# Apply formatting to each column in the Excel sheet
for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 18, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], column_formats[column][1])

# Save and close the Excel file
writer.close()
import numpy as np
import pandas as pd
import xlsxwriter as xl
import requests
import math

#An Authentic API Key must be used for app to function. Build Secrets.py and insert IEX Cloud API Token. 
from secrets import IEX_CLOUD_API_TOKEN

clean = pd.read_csv('stocks.csv')

my_columns = ['Ticker', 'Price', 'Market Capitalisation', 'Number of Shares to Buy']


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


symbol_groups = list(chunks(clean['Symbol'], 100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

final_df2 = pd.DataFrame(columns=my_columns)


for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch/?types=quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'

    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):

        final_df2 = final_df2.append(pd.Series(
            [
                symbol,
                data[symbol]['quote']['latestPrice'],
                data[symbol]['quote']['marketCap'],
                'N/A'],
            index=my_columns),
            ignore_index=True
        )

portfolio_size = float('1000000')
# input('Enter the value of your portfolio: ')

#try:
#     val = float(portfolio_size)
#     print(val)
# except ValueError:
#     print('Value Error')
#     print('That is not a number')
#     size = input('Enter the value of your portfolio: ')


position_size = float(portfolio_size) / len(final_df2.index)
for i in range(0, len(final_df2['Ticker'])-1):
    final_df2.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / final_df2['Price'][i])

print(final_df2)

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
final_df2.to_excel(writer, sheet_name='Recommended Trades', index = False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

dollar_format = writer.book.add_format(
    {
        'num_format':'$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

integer_format = writer.book.add_format(
    {
        'num_format':'0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)


column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Price', dollar_format],
    'C': ['Market Capitalization', dollar_format],
    'D': ['Number of Shares to Buy', integer_format]
}

for column in column_formats.keys():
    writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
    writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()

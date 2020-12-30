import numpy as np
import pandas as pd
import requests
import xlsxwriter
import math

stocks = pd.read_csv('newstocks.csv')
my_columns = ['Ticker','Stocks price','Market Cap','Shares to buy']

def chunks(lst, n):
    for i in range(0,len(lst),n):
        yield lst[i:i+n]

symbol_groups = list(chunks(stocks['Ticker'],100))
symbol_strings = []

for i in range(0, len(symbol_groups)):
  symbol_strings.append(','.join(symbol_groups[i]))


final_dataframe = pd.DataFrame(columns = my_columns)
for symbol_string in symbol_strings:
    batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={symbol_string}&types=quote&token=Tpk_059b97af715d417d9f49f50b51b1c448'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        final_dataframe=final_dataframe.append(
            pd.Series([
            symbol,
            data[symbol]['quote']['latestPrice'],
            data[symbol]['quote']['marketCap'],
            'n/a'
           ],index=my_columns),ignore_index = True
        )

portfolio_size=input('Enter the value of your portfolio:')
try:
    val = float(portfolio_size)
    print(val)
except ValueError:
    print('enter valid number')
    portfolio_size=input('Enter the value of your portfolio:')
    val = float(portfolio_size)

position_size=val/len(final_dataframe.index)
for i in range(0,len(final_dataframe.index)):
    final_dataframe.loc[i,'Shares to buy']=math.floor(position_size/final_dataframe.loc[i,'Stocks price'])

writer = pd.ExcelWriter('recommended trades.xlsx', engine = 'xlsxwriter')
final_dataframe.to_excel(writer,'recommended trades',index = False)
background_color = '#569e04'
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
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)


column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Stock Price', dollar_format],
    'C':['Market Cap',dollar_format],
    'D':['Number of Shares to buy',integer_format]
}

for column in column_formats.keys():
    writer.sheets['recommended trades'].set_column(f'{column}:{column}',18, column_formats[column][1])
    writer.sheets['recommended trades'].write(f'{column}1',column_formats[column][0],column_formats[column][1])
writer.save()

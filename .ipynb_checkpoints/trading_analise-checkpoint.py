import yfinance as yf
import datetime as dt
from pandas_datareader import data as pdr

end = dt.datetime.now()
start = end - dt.timedelta(5000)

stockList = ['GOOG','AAPL']

df = pdr.get_data_yahoo(stockList, start, end)
df.head()
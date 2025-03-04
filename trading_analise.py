import pandas as pd
import datetime as dt
import yfinance as yf
import xlsxwriter
import os
import signal 
import subprocess
from openpyxl import load_workbook
from packaging.version import Version

def terminate_excel():
    process = subprocess.Popen(['tasklist'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    stdout, stderr = process.communicate()

    process_name = 'EXCEL.EXE'
    lines = stdout.splitlines()
    for line in lines:
        if process_name in line:
            parts = line.split()
            pid = int(parts[1])  # The PID is the second element in the split line
            os.kill(pid, signal.CTRL_BREAK_EVENT)
            return
    print(f"{process_name} is not running.")

terminate_excel()

print('Welcome to Trading 101. Please choose what stock you would like to analyze, \nstarting date and ending date of period. I hope it will help :)\n')

stock_list = input("Insert stocks in which you are interested in: ").split()
start = input("Insert start-date (YYYY-MM-DD): ")
end = input("Insert end-date (YYYY-MM-DD): ")

if len(stock_list) == 4:
    position = ['AA2', 'AA17', 'AA32', 'AA47', 'AA62', 'AA77']
if len(stock_list) == 3:
    position = ['U2', 'U17', 'U32', 'U47', 'U62', 'U77']
if len(stock_list) == 2:
    position = ['O2', 'O17', 'O32', 'O47', 'O62', 'O77']
if len(stock_list) == 1:
    position = ['I2', 'I17', 'I32', 'I47', 'I62', 'I77']

colors = ['#4299f5', '#5d42f5', '#ce42f5', '#f5425d', '#000000', '#FFFFFF']

# Fetch stock data using yfinance
stock_data = {}
for stock in stock_list:
    try:
        stock_data[stock] = yf.download(stock, start=start, end=end)
    except Exception as e:
        print(f"Failed to download data for {stock}: {e}")

# Combine all stock data into a single DataFrame
df = pd.concat(stock_data.values(), keys=stock_data.keys(), axis=1)
output_path = 'D:/git/cryptoApp/ports/crypto_report.xlsx'
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    worksheet.set_column('A:A', 30)

    y = 1

    for i, stock in enumerate(stock_data.keys()):
        chart = workbook.add_chart({'type': 'line'})

        chart.add_series({
            'name': stock,
            'categories': ['Sheet1', 1, 0, len(df), 0],
            'values': ['Sheet1', 1, y, len(df), y],
            'line': {'color': colors[i]},
        })

        chart.set_title({'name': stock})
        chart.set_x_axis({'name': 'Date'})
        chart.set_y_axis({'name': 'USD$'})

        worksheet.insert_chart(position[i], chart)
        y += 1

os.startfile(output_path)
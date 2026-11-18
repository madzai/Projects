import pandas as pd

# filename of the csv file of single column Symbol
source_file = "mt5_symbols.csv"
# filename of the text file to be uploaded to TradingView
output_file = "watchlist.txt"

# Read the CSV file
data = pd.read_csv(source_file)

# Convert specific column into a list
symbols = data['Symbol'].tolist()

# Combine the list of symbols into string (comma-separated, without space)
mt5_us = ""

for symbol in symbols:
  mt5_us += symbol + ','

# remove last char (comma)
mt5_us = mt5_us[:-1]

with open(output_file, "a") as f:
  f.write(mt5_us)
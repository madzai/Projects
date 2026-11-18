import pandas as pd

source_file = "mt5_symbols.csv"

# Reading the CSV file
data = pd.read_csv(source_file)

# Converting specific columns into lists
symbols = data['Symbol'].tolist()

mt5_us = ""

for symbol in symbols:
  mt5_us += symbol + ','

# remove last char (comma)
mt5_us = mt5_us[:-1]

print(mt5_us)
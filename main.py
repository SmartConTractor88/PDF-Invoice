import pandas as pd # for reading .xslx
#import openpyxl
import glob

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)

for file in filepaths:
    df = pd.read_excel(file, sheet_name="Sheet 1")
    print(df)

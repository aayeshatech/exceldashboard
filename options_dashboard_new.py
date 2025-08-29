import xlwings as xw
import pandas as pd

wb = xw.Book("Live_Option_Chain_Terminal.xlsm")   # must be OPEN in Excel
sht = wb.sheets["OC_1"]                           # pick Nifty sheet
df = sht.range("A1").expand().options(pd.DataFrame, header=1, index=False).value

print(df.head())

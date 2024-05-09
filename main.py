import pandas as pd
import openpyxl
sheetName= pd.ExcelFile("meibo.xlsx")
df = pd.read_excel("部員名簿.xlsx")
print(df)


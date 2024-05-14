import pandas as pd
import openpyxl

# サンプルデータフレームの作成
df1 = pd.DataFrame({
    'Column1': [1, 2, 3],
    'Column2': ['A', 'B', 'C']
})

df2 = pd.DataFrame({
    'Column1': [4, 5, 6],
    'Column2': ['D', 'E', 'F']
})

df3 = pd.DataFrame({
    'Column1': [7, 8, 9],
    'Column2': ['G', 'H', 'I']
})

# データフレームをディクショナリにまとめる（キーはシート名）
dataframes = {
    'Sheet1': df1,
    'Sheet2': df2,
    'Sheet3': df3
}

# Excelファイルに書き込む
with pd.ExcelWriter('multiple_sheets.xlsx', engine='openpyxl') as writer:
    for sheet_name, df in dataframes.items():
        df.to_excel(writer, sheet_name="Sheet1", index=False)

print("データフレームがExcelファイルに書き込まれました。")

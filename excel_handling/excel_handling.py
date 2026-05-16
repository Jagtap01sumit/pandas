import pandas as pd

file_path = "..\downloads\Monthly_Portfolio_30042026.xlsx"

# This is the MOST commonly used function.
# It:
# opens Excel file
# reads data
# returns DataFrame directly
df = pd.read_excel(file_path)

# This prints the entire DataFrame stored in df.
# print(df)


# df.head() is used to display the first 5 rows of a DataFrame in pandas.
# print(df.head())


# df.head() is used to display the first 3 rows of a DataFrame in pandas.
# print(df.head(3))


# df.tail() is used to display the Last 5 rows
print(df.tail())


# get all the sheets name from the workbook
sheet_data = pd.read_excel(file_path, sheet_name=None)
print(sheet_data.keys())
# dict_keys(['qActive', 'qMPU', 'qMCN', 'qMBS', 'qSCF', 'qMLC', 'qMBC', 'qESG', 'qFF', 'qMCF', 'qMTK', 'qMCO', 'qLF', 'qFLEXI', 'qMES', 'qMHC', 'qIF', 'qMON', 'qMQM', 'qMAB', 'qMDA', 'qMGG', 'qMAF', 'qL&MF', 'qMMF', 'qTP', 'qAF', 'qMMO', 'qVF'])


# if file has multiple sheets in workbook so we can access particular sheets
df = pd.read_excel(file_path, sheet_name="qSCF")
# print(df);

#OR

# sheet_data = pd.read_excel(file_path, sheet_name=None)
# print(sheet_data["qSCF"].head())


excel = pd.ExcelFile(file_path)
# print(excel.sheet_names)



# This only opens workbook structure.
# It does NOT directly load sheet data.
# Useful when:
# file has multiple sheets
# you want sheet names first
# you want to process sheets dynamically

excel = pd.ExcelFile(file_path)
print(excel.sheet_names)
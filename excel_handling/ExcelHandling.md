# Pandas Excel Handling Notes

## Introduction

This guide explains how to use pandas for:
- Reading Excel files
- Reading multiple sheets
- Viewing data
- Accessing workbook information
- Understanding DataFrames

---

# Import Pandas

```python
import pandas as pd
```

`pd` is an alias (short name) for pandas.

---

# Define File Path

```python
file_path = "downloads\\Monthly_Portfolio_30042026.xlsx"
```

This stores the Excel file location inside a variable.

---

# Reading Excel File

## Code

```python
df = pd.read_excel(file_path)
```

---

## Explanation

### `pd.read_excel()`

This is the MOST commonly used pandas function.

It:
1. Opens the Excel file
2. Reads sheet data
3. Converts it into a DataFrame

---

# What is DataFrame?

A DataFrame is a table structure in pandas.

Example:

| Name | Age |
|---|---|
| Sumit | 22 |
| Rahul | 23 |

This entire table is called a DataFrame.

---

# Print Entire DataFrame

## Code

```python
print(df)
```

---

# View First Rows Using `head()`

## First 5 Rows

```python
print(df.head())
```

---

## Explanation

`head()` shows:
- first 5 rows by default

Useful for:
- checking data quickly
- debugging
- previewing large datasets

---

# View First 3 Rows

## Code

```python
print(df.head(3))
```

---

## Explanation

Displays only:
- first 3 rows

---

# View Last Rows Using `tail()`

## Code

```python
print(df.tail())
```

---

## Explanation

`tail()` shows:
- last 5 rows by default

Useful for:
- checking ending records
- verifying appended data

---

# Read All Sheets from Workbook

## Code

```python
sheet_data = pd.read_excel(file_path, sheet_name=None)

print(sheet_data.keys())
```

---

# Explanation

## `sheet_name=None`

This tells pandas:

```plaintext
Read ALL sheets from workbook
```

---

# Returned Data Type

Pandas returns a:

```plaintext
Dictionary
```

Where:

| Key | Value |
|---|---|
| Sheet Name | DataFrame |

---

# Example Output

```python
dict_keys([
'qActive',
'qMPU',
'qMCN',
'qMBS',
'qSCF'
])
```

These are workbook sheet names.

---

# Access Particular Sheet

## Method 1

```python
df = pd.read_excel(file_path, sheet_name="qSCF")

print(df)
```

---

# Explanation

Reads only:

```plaintext
qSCF sheet
```

from workbook.

---

# Alternative Method

## Method 2

```python
sheet_data = pd.read_excel(file_path, sheet_name=None)

print(sheet_data["qSCF"].head())
```

---

# Explanation

- Reads all sheets
- Accesses specific sheet using dictionary key

---

# Using `pd.ExcelFile()`

## Code

```python
excel = pd.ExcelFile(file_path)

print(excel.sheet_names)
```

---

# Explanation

`ExcelFile()`:
- opens workbook structure
- does NOT directly load sheet data

Useful when:
- workbook has many sheets
- you want sheet names first
- dynamic sheet processing is required

---

# Output Example

```python
[
'qActive',
'qMPU',
'qMCN',
'qMBS',
'qSCF'
]
```

---

# Difference Between `read_excel()` and `ExcelFile()`

| Feature | `pd.read_excel()` | `pd.ExcelFile()` |
|---|---|---|
| Reads data directly | Yes | No |
| Returns DataFrame | Yes | No |
| Returns workbook object | No | Yes |
| Best for single sheet | Yes | Yes |
| Best for multiple sheets | Moderate | Excellent |
| Can get sheet names | Indirectly | Directly |
| Memory efficient | Moderate | Better |

---

# Workflow Understanding

## `read_excel()`

```plaintext
Excel File → DataFrame
```

---

## `ExcelFile()`

```plaintext
Excel File → Workbook Object → Sheet Selection → DataFrame
```

---

# Important Notes

## `head()`

| Function | Meaning |
|---|---|
| `head()` | First 5 rows |
| `head(3)` | First 3 rows |

---

## `tail()`

| Function | Meaning |
|---|---|
| `tail()` | Last 5 rows |

---

# Real-World Use Cases

These concepts are heavily used in:
- Mutual fund analysis
- Bank statement processing
- Stock dashboards
- Automation pipelines
- Data cleaning systems
- Excel report generators

---

# Complete Code

```python
import pandas as pd

file_path = "downloads\\Monthly_Portfolio_30042026.xlsx"

# Read Excel file
df = pd.read_excel(file_path)

# Print complete DataFrame
# print(df)

# First 5 rows
# print(df.head())

# First 3 rows
# print(df.head(3))

# Last 5 rows
print(df.tail())

# Get all sheet names using read_excel
sheet_data = pd.read_excel(file_path, sheet_name=None)

print(sheet_data.keys())

# Read specific sheet
df = pd.read_excel(file_path, sheet_name="qSCF")

# print(df)

# Alternative way
# print(sheet_data["qSCF"].head())

# Open workbook object
excel = pd.ExcelFile(file_path)

# Print all sheet names
print(excel.sheet_names)
```

---

# Summary Table

| Topic | Function |
|---|---|
| Read Excel | `pd.read_excel()` |
| Open Workbook | `pd.ExcelFile()` |
| First Rows | `head()` |
| Last Rows | `tail()` |
| All Sheets | `sheet_name=None` |
| Specific Sheet | `sheet_name="Sheet1"` |
| Sheet Names | `excel.sheet_names` |

---

# Important Beginner Notes

## Common Mistake

Do NOT name your Python file:

```plaintext
pandas.py
```

because it conflicts with pandas library.

Correct examples:
- `main.py`
- `excel_learning.py`

---

# Understanding Variables

| Variable | Meaning |
|---|---|
| `df` | DataFrame |
| `sheet_data` | Dictionary containing sheets |
| `excel` | Workbook object |

---

# Best Practices

| Situation | Recommended Function |
|---|---|
| Single sheet reading | `pd.read_excel()` |
| Need sheet names | `pd.ExcelFile()` |
| Read all sheets | `sheet_name=None` |
| Large workbook processing | `ExcelFile()` |

---

# Next Topics To Learn

- Selecting columns
- Filtering rows
- Handling null values
- Sorting data
- GroupBy
- Merging Excel files
- Pivot tables
- Exporting to Excel
- Data cleaning
- Financial analysis using pandas
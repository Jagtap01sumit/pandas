# Pandas Excel Handling Notes

This guide explains how to use pandas for:
- Reading Excel files
- Reading multiple sheets
- Viewing data
- Accessing workbook information
- Understanding DataFrames

---

#### Import Pandas

```python
import pandas as pd
```

`pd` is an alias (short name) for pandas.

---

#### Define File Path

```python
file_path = "downloads\\Monthly_Portfolio_30042026.xlsx"
```

This stores the Excel file location inside a variable.

---

#### Reading Excel File

```python
df = pd.read_excel(file_path)
```
This is the MOST commonly used pandas function.

It:
1. Opens the Excel file
2. Reads sheet data
3. Converts it into a DataFrame

---

#### What is DataFrame?

A DataFrame is a table structure in pandas.

Example:

| Name | Age |
|---|---|
| Sumit | 22 |
| Rahul | 23 |

This entire table is called a DataFrame.

#### Print Entire DataFrame

```python
print(df)
```

---

#### View First Rows Using `head()`

- First 5 Rows

```python
print(df.head())
```

`head()` shows:
- first 5 rows by default

Useful for:
- checking data quickly
- debugging
- previewing large datasets

---

- View First 3 Rows

```python
print(df.head(3))
```
Displays only:
- first 3 rows

---

#### View Last Rows Using `tail()`

```python
print(df.tail())
```

`tail()` shows:
- last 5 rows by default

Useful for:
- checking ending records
- verifying appended data

---

#### Read All Sheets from Workbook

## Code

```python
sheet_data = pd.read_excel(file_path, sheet_name=None)

print(sheet_data.keys())
```
#### `sheet_name=None`

This tells pandas:

```plaintext
Read ALL sheets from workbook
```

---

#### Returned Data Type

Pandas returns a:

```plaintext
Dictionary
```

Where:

| Key | Value |
|---|---|
| Sheet Name | DataFrame |

---

- Example Output

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

#### Access Particular Sheet

## Method 1

```python
df = pd.read_excel(file_path, sheet_name="qSCF")

print(df)
```

Reads only:

```plaintext
qSCF sheet
```

from workbook.

---

#### Alternative Method

#### Method 2

```python
sheet_data = pd.read_excel(file_path, sheet_name=None)

print(sheet_data["qSCF"].head())
```

- Reads all sheets
- Accesses specific sheet using dictionary key

---

#### Using `pd.ExcelFile()`

```python
excel = pd.ExcelFile(file_path)

print(excel.sheet_names)
```

`ExcelFile()`:
- opens workbook structure
- does NOT directly load sheet data

Useful when:
- workbook has many sheets
- you want sheet names first
- dynamic sheet processing is required

```python
[
'qActive',
'qMPU',
'qMCN',
'qMBS',
'qSCF'
]
```
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

#### Workflow Understanding

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

# Best Practices

| Situation | Recommended Function |
|---|---|
| Single sheet reading | `pd.read_excel()` |
| Need sheet names | `pd.ExcelFile()` |
| Read all sheets | `sheet_name=None` |
| Large workbook processing | `ExcelFile()` |

---

# SharePoint Excel

A Python class for interacting with Sharepoint Excel via `ExcelWorksheet` objects. 

Supports retrieving, adding, deleting, and modifying rows in a SharePoint Excel file, as well as locking and unlocking sheet protection.

---

## Import

```python
from src.graph_api_functions.excel import ExcelWorksheet
```

## Setup
Instantiate with `.setup()`. Four arguments are always required:
 
| Argument | Description | Example |
|---|---|---|
| `hostname` | SharePoint host name | `phila.sharepoint.com` |
| `site_name` | SharePoint site name | `ps360-metrics` |
| `document_library` | Name of the document library | `Documents` |
| `workbook_path` | Path of the Excel workbook | `"Philly Stat - OIT/OIT_data.xlsx"` |


Two additional arguments are optional, depending on which functions you plan to run:
 
| Argument | Note | Example | Required for |
|---|---|---|---|
| `worksheet_name` | Name of the worksheet | `Metrics` | `toggle_protection` |
| `table_name` | Name of the table | `Table1` | `add_rows`, `delete_row_by_pk`, `update_row_by_pk`, `list_rows`, `list_columns` |


```python
excel = ExcelWorksheet.setup(
    hostname=hostname,
    site_name=site_name,
    document_library=document_library,
    workbook_path=workbook_path,
    worksheet_name=worksheet_name,
    table_name=table_name
)
```
 
---

## Viewing Data
 
### List Columns

Returns a list of column names for a specific excel table. Column names are returned in the order they appear in the table. 

```python
columns = excel.list_columns()
```

### List Rows
 
Returns the rows of an Excel table as a `List[dict]`. Each dictionary's keys are column names, and the values are the corresponding cell values.

```python
import pandas as pd 

rows = excel.list_rows() 

# Can easily convert to a dataframe view if needed
df = pd.DataFrame(rows)
df.head()
```

---

## Modifying Data
 
### Add Rows
 
Appends one or more rows to the bottom of the Excel worksheet.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `data` | `List[List]` | Values to insert, in column order. Multiple rows can be loaded at once. |
 
**Raises**
 
- `MalformedRowError` — the number of values to insert doesn't match the number of columns in the table.

For example, given three columns `col1`, `col2`, `col3`, and `data=[[6, 12, 36]]`, the table becomes:
 
| col1 | col2 | col3 |
|---|---|---|
| ... | ... | ... |
| 6 | 12 | 36 |
 
```python
data = [[6, 12, 36], [1, 2, 3]]
excel.add_rows(data)
```

### Delete Row by Primary Key Value
 
Deletes a row based on a primary key value.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `pk_col` | `str` | Name of the primary key column |
| `pk_val` | `str \| int` | Value to match on |
| `password` | `str`, optional | Sheet protection password |
 
**Raises**
 
- `KeyError` — the key isn't found as a column in the table.
- `PrimaryKeyValueNotFound` — the value to look for doesn't exist.
- `ValueError` — the operation is attempted on an empty table.
```python
pk_col = "primary_key"
pk_val = "n"
 
resp = excel.delete_row_by_pk(pk_col=pk_col, pk_val=pk_val)
```

### Delete Row at an Index
 
Deletes a row at a specified index from the sheet.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `index` | `int` | 0-based index of the row to delete (e.g. `index=0` deletes the first row of data) |
| `password` | `str`, optional | Sheet protection password |
 
```python
excel.delete_row_at_index(index=3)
```

### Update Row by Primary Key Value
 
Updates a row based on a primary key value.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `pk_col` | `str` | Name of the primary key column |
| `pk_val` | `str \| int` | Value to match on |
| `value` | `list` | Nested list of update data, e.g. `value=[[col1_update, None, col3_update]]` |
| `password` | `str`, optional | Sheet protection password |
 
**Raises**
 
- `KeyError` — the key isn't found as a column in the table.
- `PrimaryKeyValueNotFound` — the value to look for doesn't exist.
- `ValueError` — the operation is attempted on an empty table.
```python
data = [[None, None, 1738]]
 
pk_col = "primary_key"
pk_val = "v"
 
resp = excel.update_row_by_pk(pk_col=pk_col, pk_val=pk_val, value=data)
```

### Update Row at an Index
 
Updates an existing row in an Excel table with the values in `value`. The order of values must mirror the order of columns in the table. To leave a particular cell unchanged, pass `None` in its position.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `index` | `int` | 0-based index of the row to update (e.g. `index=12` updates the 13th row of data) |
| `value` | `list` | Nested list of update data, e.g. `value=[[col1_update, None, col3_update]]` |
| `password` | `str`, optional | Sheet protection password |
 
```python
# Table has 3 columns: col1, col2, col3
update = [[None, 1738, None]]  # Update only col2 value
 
resp = excel.update_row_at_index(index=11, value=data)
```
 
---
 
## Miscellaneous
 
### Toggle Protection
 
Toggles protection for a specific worksheet on or off.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `password` | `str` | Sheet protection password |
| `protect` | `bool` | `True` to protect, `False` to unprotect |
 
```python
# Unprotect
password = "#######"
excel.toggle_protection(password=password, protect=False)
 
# Protect
excel.toggle_protection(password=password, protect=True)
```
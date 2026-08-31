# SharePoint Excel

Class for interacting with Sharepoint Excel objects. Allows one to perform operations such as adding/viewing rows and columns of data and locking/unlocking sheet protection. 
## Import

```python
from src.graph_api_functions.excel import ExcelWorksheet
```

## Instantiate with .setup()
There are 4 required arguments needed to setup an _ExcelWorksheet_ object. 

| Argument | Note |Example|
| ---         | ---   |---
| **hostname**    | Sharepoint Host Name   | _phila.sharepoint.com_|
| **site_name**   | Sharepoint site name  | _ps360-metrics_|
| **document_library**   | The name of the document library  | _Documents_|
| **workbook_path**   | The path of the excel workbook  | _"Philly Stat - OIT/OIT_data.xlsx"_|


There are 2 optional parameters to set, depending on what kind of functions you want to  run. 

| Argument | Note |Example| Function to Run |
| ---         | ---   |--- | ---|
| **worksheet_name**    | The name of the workshee   | _Metrics_|_toggle_protection_|
| **table_name**   | The name of the table  | _Table1_| _add_rows_, _list_rows_, and _list_columns_|


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

## Functions

### List Columns 
Returns a list of column names for a specific excel table. Column names are returned in the order they appear in the table. 

```python
columns = excel.list_columns()
```

### List Rows 
Returns the rows of an excel table, represented as a List[dict]. The keys to the dictionary represent the column name, and the value is the value of the column. 

```python
import pandas as pd 

rows = excel.list_rows() 

# Can easily convert to a dataframe view if needed
df = pd.DataFrame(rows)
df.head()
```

### Add Rows
The function appends a row(s) to the bottom of the excel worksheet.

Arguments:

- **data** (List[List]): Data you want to insert. _data_ must appear in the order to be inserted. Can load multiple rows at once. 

Throws: 

`MalformedRowError` if the number of values to insert is not the same as the number of columns in the table.

For example, if you have three columns, col1, col2, col3, and set `data=[[6,12,36]]`, the table will look like: 

| col1 | col2 |col3|
| --- | --- |---
| ...    | ...   | ...|
| 6    | 12   | 36|


```python
data = [[6,12,36], [1,2,3]]
excel.add_rows(data)
```

### Delete Row 
The function deletes a row at a specified index from the sheet. 

Arguments:

- **index** (int): Index you want to delete from, 0-based. 
  - e.g. To delete the first row of data, set `index=0`
- **password** (str, optional): The sheet protection password 

```python
excel.delete_row(index=3)
```

### Toggle Protection 
Toggles protection for a specific worksheet on or off. 

Arguments:
- **password** str: The sheet protection password. 
- **protect** bool: Boolean to protect (True) or unprotect (False) 

```python
# Unprotect 
password = '#######'
excel.toggle_protection(password=password, protect=False)
...
# Protect
excel.toggle_protection(password=password, protect=True)
```


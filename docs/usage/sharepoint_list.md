# SharePoint List

## Import
```python
from graph_api_functions.list import SharePointList
```

## Instantiate with .setup()
```python
sp_list = SharePointList.setup(site_name="list_site_name", list_name="list_name")
```

## List Columns
```python
columns = sp_list.list_columns()

for column in columns: # Returns an iterator of SharePointListColumn objects
    print(column)
```
## List Rows
```python

rows = sp_list.list_rows() # Returns an iterator of dicts

for row in rows:
    print(row)
```

## Add Row
```python
row_to_add = {"name": "Billy Penn", 
              "address": "1234 Market St, Philadelphia, PA", 
              "age": 250}


response = sp_list.add_row(row_to_add) # Returns a JSON response
```
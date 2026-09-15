# SharePoint List

## Import
```python
from espy.list import SharePointList
```

## Instantiate with .setup()
```python
sp_list = SharePointList.setup(site_name="list_site_name", list_name="list_name")
```

## List Columns
Lists all user-added columns in the list. Does not include hidden programmatic columns.

```python
columns = sp_list.list_columns()

for column in columns: # Returns an iterator of SharePointListColumn objects
    print(column)
```
## List Rows
Returns all rows in the list as an iterator of dictionaries.

```python

rows = sp_list.list_rows() # Returns an iterator of dicts

for row in rows:
    print(row)
```
## Get Row
Gets one row from the list, specified by primary key.

```python
row = sp_list.get_row("user_id", "1234") # Returns a dict
```

## Add Row
Add a row to the list.

```python
row_to_add = {
              "user_id": 1,
              "name": "Billy Penn", 
              "address": "1234 Market St, Philadelphia, PA", 
              "age": 250}


response = sp_list.add_row(row_to_add) # Returns a JSON response
```

## Edit Row
Edit a row in the list, use primary key to identify the row.

```python
row_to_edit = {
              "user_id": 1,
              "name": "William Penn", 
              "address": "1234 Market St, Philadelphia, PA", 
              "age": 250}

response = sp_list.edit_row("user_id", 1, row_to_edit)
```

## Delete Row
Delete a row in a list, use primary key to identify the row.

!!! warning 
    Use very carefully! Deleted data cannot be recovered!

```python
response = sp_list.delete_row("user_id", 1) # Returns a JSON http response
```

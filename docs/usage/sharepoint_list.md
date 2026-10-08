# SharePoint List

## Import
```python
from espy.list import SharePointList
```

## Instantiate with .setup()
All arguments are keyword-only.

```python
sp_list = SharePointList.setup(
    site_name="list_site_name", list_name="list_name"
)
```

Optional arguments:

| Argument | Default | Description |
|---|---|---|
| `hostname` | `None` | SharePoint host name, e.g. `example.sharepoint.com`. If omitted, read from `SHAREPOINT_HOSTNAME`. |
| `creds` | `None` | `{"tenant_id", "client_id", "client_secret"}`. If omitted, read from `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`. |

## Summary Methods
Methods which provide information about the entire list.

### List Columns
Lists all user-added columns in the list. Does not include hidden programmatic columns.

```python
columns = sp_list.list_columns()

for column in columns:  # Returns an iterator of SharePointListColumn objects
    print(column)
```
### List Rows
Returns all rows in the list as an iterator of dictionaries.

```python
rows = sp_list.list_rows()  # Returns an iterator of dicts

for row in rows:
    print(row)
```

## Single-Row Methods
Methods which perform an operation against a single row.

### Get Row
Gets one row from the list, specified by primary key.

```python
row = sp_list.get_row("user_id", "1234")  # Returns a dict
```

### Add Row
Add a row to the list.

```python
row_to_add = {
    "user_id": 1,
    "name": "Billy Penn",
    "address": "1234 Market St, Philadelphia, PA",
    "age": 250,
}


response = sp_list.add_row(row_to_add)  # Returns a JSON response
```

### Edit Row
Edit a row in the list, use primary key to identify the row.

```python
row_to_edit = {
    "user_id": "1",
    "name": "William Penn",
    "address": "1234 Market St, Philadelphia, PA",
    "age": 250,
}

response = sp_list.edit_row("user_id", "1", row_to_edit)
```

### Delete Row
Delete a row in a list, use primary key to identify the row.

!!! warning 
    Use very carefully! Deleted data cannot be recovered!

```python
response = sp_list.delete_row("user_id", "1")  # Returns a JSON http response
```

## Batch Methods
Methods that perform operations against multiple rows. Operations are
performed in batches of 20, and the result of each operation is returned
in a list of `BatchResult` objects, which can be parsed by the consumer.

!!! warning
    Batch results may be a mixture of successes and failures. Parse the response
    data to examine and handle failures.

### Get Rows
Gets multiple rows from the list, specified by primary key.

```python
key_col = "user_id"
values = ["24", "25", "26"]
response = sp_list.get_rows(key_col=key_col, values=values) # Returns a list of BatchResponse objects
```

### Add Rows
Add multiple rows to the list.

```python
rows_to_add = [{
    "user_id": "1",
    "name": "Billy Penn",
    "address": "1234 Market St, Philadelphia, PA",
    "age": 381,
},
{
    "user_id": "2",
    "name": "Betsy Ross",
    "address": "239 Arch St, Philadelphia, PA",
    "age": 274,
}]
response = sp_list.add_rows(rows_to_add)  # Returns a list of BatchResponse objects
```

### Edit Rows
Edits multiple rows in the list, use primary key to identify the rows, and values to determine which rows should be edited.

```python
rows_to_edit = [{
    "user_id": "1",
    "name": "Billy Penn",
    "address": "1234 Market St, Philadelphia, PA",
    "age": 381,
},
{
    "user_id": "2",
    "name": "Betsy Ross",
    "address": "239 Arch St, Philadelphia, PA",
    "age": 274,
}]

key_col = "user_id"
values = ["1", "2"]

response = sp_list.edit_rows(key_col=key_col, values=values, data=data_to_edit) # Returns a list of BatchResponse objects
```

### Delete Rows
Delete multiple rows in a list, use primary key to identify the rows.

!!! warning 
    Use very carefully! Deleted data cannot be recovered!

```python
key_col = "user_id"
values = ["24", "25", "26"]  # Returns a JSON http response
response = sp_list.delete_rows(key_col=key_col, values=values)
```
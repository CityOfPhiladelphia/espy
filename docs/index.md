# Graph API Functions

Graph API functions is a project to make using the Microsoft Graph API
easier.

## Core Features
- **Clear Models**: Each type of file in SharePoint (Excel, List) has its own class.
- **Easy Authentication**: Handles the authentication workflow. Just provide
site name and file name.

## Getting Started

To install this library to your project, run:

```bash
uv add git+https://github.com/CityOfPhiladelphia/graph_api_functions.git
```

### Usage Example
```python
from graph_api_functions.list import SharePointList

# Instantiate a SharePoint List Object
sp_list = SharePointList.setup(site_name="list_site_name", list_name="list_name")

# Create a row mapping data to column names in the list
row_to_add = {"name": "Billy Penn", 
              "address": "1234 Market St, Philadelphia, PA", 
              "age": 250}


response = sp_list.add_row(row_to_add)
```

## Table of Contents
- [Usage](./usage/)
- [Development](./development/)
- [Reference](./reference/)
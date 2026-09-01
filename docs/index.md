# EsPy

EsPy is a project to make using the Microsoft Graph API
easier.

!!! note

    This library necessitates a `Graph App` being created and registered in Azure (Entra ID). Further this app needs to be provisioned to target sharepoint sites with read/write permissions. At the moment, the library is defaulting to using CityGeo's `Graph App`, which is provisioned for a handful of sharepoint sites. Future development will allow users to pass credentials for their own `Graph App`, but at the moment this will only work with CityGeo's `Graph App`.


## Core Features
- **Clear Models**: Each type of file in SharePoint (Excel, List) has its own class.
- **Easy Authentication**: Handles the authentication workflow. Just provide
site name and file name.

## Getting Started

To install this library to your project, run:

Include the tag for the latest version.

```bash
uv add "git+https://github.com/CityOfPhiladelphia/espy.git" --tag v1.0.0
```

### Upgrading EsPy
```
uv lock --upgrade-package espy && uv sync
```

### Usage Example
```python
from espy.list import SharePointList

# Instantiate a SharePoint List Object
sp_list = SharePointList.setup(site_name="list_site_name", list_name="list_name")

# Create a row mapping data to column names in the list
row_to_add = {"name": "Billy Penn", 
              "address": "1234 Market St, Philadelphia, PA", 
              "age": 250}


response = sp_list.add_row(row_to_add)
```
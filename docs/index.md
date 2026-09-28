# EsPy

EsPy is a project to make using the Microsoft Graph API
easier.

!!! note

    This library requires a `Graph App` registered in Azure (Entra ID) and provisioned with read/write permissions on the target SharePoint sites. See [Authentication](#authentication) for how to provide its credentials.


## Core Features
- **Clear Models**: Each type of file in SharePoint (Excel, List) has its own class.
- **Easy Authentication**: Handles the authentication workflow. Just provide
site name and file name.

## Getting Started

To install this library to your project, run:

Include the tag for the latest version.

```bash
uv add "git+https://github.com/CityOfPhiladelphia/espy.git" --tag v1.1.0
```

### Upgrading EsPy
```
uv lock --upgrade-package espy && uv sync
```

### Authentication
EsPy authenticates with an Azure (Entra ID) app registration that has read/write
access to your SharePoint sites. Provide its credentials in one of two ways.

**Environment variables (recommended).** If no `creds` are passed, EsPy reads:

```bash
export AZURE_TENANT_ID="..."
export AZURE_CLIENT_ID="..."
export AZURE_CLIENT_SECRET="..."
export SHAREPOINT_HOSTNAME="example.sharepoint.com"
```

**Passing `creds` directly** to any `setup()` or `GraphAPIClient.authenticate()` call:

```python
creds = {
    "tenant_id": "...",
    "client_id": "...",       # a.k.a. Application ID
    "client_secret": "...",   # a.k.a. Secret Value
}
sp_list = SharePointList.setup(site_name="list_site_name", list_name="list_name", creds=creds)
```

The SharePoint host is read from `SHAREPOINT_HOSTNAME` unless you pass `hostname=` to `setup()`.

### Usage Example
```python
from espy.list import SharePointList

# Instantiate a SharePoint List Object
# (credentials read from AZURE_* environment variables)
sp_list = SharePointList.setup(
    site_name="list_site_name", list_name="list_name"
)

# Create a row mapping data to column names in the list
row_to_add = {
    "name": "Billy Penn",
    "address": "1234 Market St, Philadelphia, PA",
    "age": 250,
}


response = sp_list.add_row(row_to_add)
```
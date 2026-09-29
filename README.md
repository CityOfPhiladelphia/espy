# EsPy

## Documentation
Full documentation **[can be found here.](https://probable-adventure-nyjj26e.pages.github.io/)**

Documentation is created using Zensical.

In order to serve the zensical documentation locally, run

```bash
uv run zensical serve
```

The docs will be served at ```http://localhost:3000/```.

## Getting Started
To install this library to your project, run:

```bash
uv add "git+https://github.com/CityOfPhiladelphia/espy.git" --tag v1.1.0
```

### Check for Updates
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

## Development
### To install packages needed to run:
```bash
uv sync
```
### To install packages needed to develop this package:
```bash
uv sync --dev
```

### Adding dependencies
Before you add a dependency, ask: do I need this to **only to develop** the tool,
or is this needed to run the tool?

If you only need the package for development (for example, linters like ruff, pytest, ipython, etc...), then please add the --dev flag when adding the dependency:

```bash
uv add --dev {{package name}}
```

If the dependency is truly needed for functionality, run:
```bash
uv add {{package name}}
```

### When adding a dependency:
1. Consider whether or not the dependency is necessary before adding it at all. Python's library is quite large, and may have tools that perform the functionality you're looking for.
2. Research the dependency. Is it well maintained? How many stars does the GitHub repo have? When was the last time it was updated? Is there a better, more supported dependency out there that is less likely to be abandoned in the future?
3. Add dependencies through `uv`, not through `pip`.

# SharePoint Client
 
A Python class representing a Graph API client. This class is used internally by the `ExcelWorksheet` and `SharepointList` classes to make requests on their behalf.
 
When used alone, it allows you to upload and download files to/from SharePoint.
 
---


## Import
 
```python
from src.espy.client import GraphAPIClient
```
 
## Setup
 
Instantiate with `.authenticate()`. At the moment, the `GraphAPIClient` is instantiated with CityGeo's Graph App in Azure. Future development will allow a user to pass their own credentials for authentication.
 
```python
client = GraphAPIClient.authenticate()
```
 
---

## Functions

### Upload a Local File
 
Uploads a file to a SharePoint Documents folder.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `site_name` | `str` | Name of the SharePoint site to upload the file to |
| `local_path` | `str` | Exact local path where the file is located |
| `dest_path` | `str` | Path, relative to the Documents folder, where the file should be saved |
 
**Returns**
 
| Type | Description |
|---|---|
| `dict` | The response JSON |
 
For example:
 
- `dest_path="Folder Name"` creates a file at `Documents/Folder Name/file.xlsx`
- `dest_path=""` saves it at `Documents/file.xlsx`
```python
client.upload_local_file(site_name=site_name,
                          local_path=local_path,
                          dest_path=dest_path)
```
### Download a File
 
Gets the raw bytes of the Excel workbook specified by `file_path`.
 
**Arguments**
 
| Name | Type | Description |
|---|---|---|
| `hostname` | `str` | Sharepoint host name |
| `site_name` | `str` |Sharepoint site name  |
| `document_library` | `str` | The sharepoint document library.  |
| `file_path` | `str` | Path of the file to retrieve, relative to the instantiated document library |

**Returns**
 
| Type | Description |
|---|---|
| `bytes` | Raw bytes representing the Excel file |
 
```python
import pandas as pd
from io import BytesIO
 
workbook_path = "etl_tools_test_workbook.xlsx"
raw_bytes = client.get_content(file_path=workbook_path)
 
# Can convert to a dataframe easily
buf = BytesIO(raw_bytes)
df = pd.read_excel(buf)
df.head()
```
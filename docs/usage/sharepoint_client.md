# SharePoint Client

Class representing a GraphAPI client. This class is a part of the ExcelWorksheet and SharepointList classes and is responsible for making requests for them. 

When used alone, allows you to upload/download files to sharepoint. 
## Import

```python
from src.graph_api_functions.client import GraphAPIClient
```

## Instantiate with .authenticate()

```python
client = GraphAPIClient.authenticate()
```

## Functions

### Upload a Local File 
Upload a file to a SharePoint Documents folder.

Arguments:
 - site_name (str) : The name of the sharepoint site you want to upload a file to 
 - local_path (str): the exact local path where your file is 
 - dest_path (str): the path, relative to the Documents folder, where you want to save the file

Returns:
- dict: The response json 

For example: `dest_path="Folder Name"` will create a file at _Documents/Folder Name/file.xlsx_. `dest_path=""` will save it at _Documents/file.xlsx_

```python
client.upload_local_file(site_name=site_name, 
                         local_path=local_path, 
                         dest_path=dest_path)
```

### Download a File  
Get the raw bytes of the excel workbook specified by _file_path_

Arguments: 
- file_path (str): The path of the file you want, relative to the instantiated document library/ 

Returns: 
- bytes: Raw bytes representing the excel file. 

```python
import pandas as pd 
from io import BytesIO

workbook_path = 'etl_tools_test_workbook.xlsx'
raw_bytes = client.get_content(file_path=workbook_path)

# Can convert to dataframe easily 
buf = BytesIO(raw_bytes)
df = pd.read_excel(buf)
df.head()
```

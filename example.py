from src.graph_api_functions.excel import ExcelWorksheet
from src.graph_api_functions.constants import GRAPH_URL, HOST_NAME
from io import BytesIO
import pandas as pd 

site_name        = "ps360-metrics-share"
document_library = "Documents"
workbook_path    = 'etl_tools_test_workbook.xlsx'
table_name       = 'testing'
worksheet_name   = 'Dataset'

excel = ExcelWorksheet.setup(
    hostname=HOST_NAME,
    site_name=site_name,
    document_library=document_library,
    workbook_path=workbook_path,
    worksheet_name=worksheet_name,
    table_name=table_name
)

data = [[100,99,98]]

excel.append_row(data)

raw_bytes = excel.get_content(workbook_path)

buf = BytesIO(raw_bytes)

df = pd.read_excel(buf)
print(df.head())

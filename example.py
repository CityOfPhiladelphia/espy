from src.graph_api_functions.excel import ExcelWorksheet
from src.graph_api_functions.client import GraphAPIClient
from src.graph_api_functions.constants import GRAPH_URL, HOST_NAME
from io import BytesIO
import pandas as pd 

# site_name        = "ps360-metrics-share"
# document_library = "Documents"
# workbook_path    = 'etl_tools_test_workbook.xlsx'
# table_name       = 'testing'
# worksheet_name   = 'Dataset'

# excel = ExcelWorksheet.setup(
#     hostname=HOST_NAME,
#     site_name=site_name,
#     document_library=document_library,
#     workbook_path=workbook_path,
#     worksheet_name=worksheet_name,
#     table_name=table_name
# )

# data = [[100,99,98]]

# excel.append_row(data)

# raw_bytes = excel.get_content(workbook_path)
# print(type(raw_bytes))
# # buf = BytesIO(raw_bytes)

# # df = pd.read_excel(buf)
# # print(df.head())

'''Uploading a file example'''
site_name        = "ps360-metrics-share"
local_path       = "/home/ubuntu/Repos/testing/lol.xlsx"
dest_path        = "Testing" 
# want to save at Documents/Testing/lol.xlsx 

client = GraphAPIClient.authenticate()

client.upload_local_file(site_name=site_name, local_path=local_path, dest_path=dest_path)
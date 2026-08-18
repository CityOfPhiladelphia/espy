# excel.py
from graph_api_functions.client import GraphAPIClient
from graph_api_functions.constants import GRAPH_URL
from graph_api_functions.models import SharePointSiteInfo, SharePointURL
from graph_api_functions.urls import _URLResolver
from typing import Any

class ExcelWorksheet:

    def __init__(self, 
                client: GraphAPIClient, 
                hostname: str,
                site_path: str,
                document_library: str,
                workbook_path: str, 
                worksheet_name: str,
                ):
        
        self.client = client
        self.hostname = hostname
        self.site_path = site_path
        self.document_library = document_library
        self.workbook_path = workbook_path
        self.worksheet_name = worksheet_name

        self._resolver = _URLResolver(self.client)
        self._site_id = self._resolver.get_site_id(self.hostname, self.site_path)
        self._drive_id = self._resolver.get_drive_id(
            self._site_id, self.document_library
            )
        self._workbook_id = self._resolver.get_workbook_id(self._drive_id, 
            self.workbook_path
            )
        self._worksheet_url = self._build_worksheet_url()

    def _build_worksheet_url(self) -> SharePointURL:
        worksheet_url = (
            f"{GRAPH_URL}/drives/{self._drive_id}"
            f"/items/{self._workbook_id}"
            f"/workbook"
            f"/worksheets/{self.worksheet_name}"
        )

        return SharePointURL(url=worksheet_url)
        

    def get_row(self, row_id: str) -> dict [str, Any]:
            ...
    
    def list_rows(self) -> list[dict[str, Any]]:
        ...

    def add_row(self, data: dict[str, Any]) -> dict[str, Any]:
        ...

    def edit_row(self, row_id: str, 
                    data: dict[str, Any]) -> dict[str, Any]:
        ...

    def delete_row(self, row_id: str) -> bool:
        ...
    
    def upsert_row(self, key_col: str,
                    data: dict[str, Any]) -> dict[str, Any]:
        ...

    def toggle_lock(self, password: str) -> bool:
        ...


# class ExcelWorkBook:
    
#     def __init__(self, client: GraphAPIClient, site_id: str, file_path: str):
#         ...

#     def list_tables(self) -> list[str]:
#         ...

#     def get_table(self, table_name: str) -> ExcelWorksheet:
#         ...

#     def create_table(self, file_path: str, table_name: str) -> ExcelWorksheet:
#         ...



        #  def connect(self):
        #         '''
        #         Function to establish the necessary id's for a sharepoint resource. 
        #         '''
        #         if self.site_id is not None: 
        #             return 
        
        #         self.site_id = self.get_site_id()
        
        #         if self.list_name:
        #             self.list_id = self.get_list_id()
        
        #         if self.document_library:
        #             self.drive_id = self.get_drive_id(self.site_id)
                
        #         if self.workbook_path is not None and self.table_name is not None and self.worksheet_name is not None:
        #             self.item_id = self.get_workbook_id(self.drive_id)
        
        #     def protect_worksheet(self, password: str):
        #         '''
        #         Turns on sheet protection for the specified worksheet 
        
        #         Arguments:
        #         password - str that represents the sheet protection password 
        #         '''
        #         if password is None:
        #             print("No password specified, not doing anything")
        #             return 
                
        #         if self.worksheet_name is None: 
        #             print("No worksheet has been specified. Doing nothing...")
        #             return 
                
        #         print("Re-Protecting sheet...")
        #         url = (
        #             f"{GRAPH_URL}/drives/{self.drive_id}"
        #             f"/items/{self.item_id}"
        #             f"/workbook"
        #             f"/worksheets/{self.worksheet_name}"
        #             f"/protection/protect"
        #         )
        
        #         body = {
        #             "password": password
        #         }
        
        #         self._post(url, body)
        
        #         print("Protection re-enabled!!!")
            
        #     def unprotect_worksheet(self, password):
        #         '''
        #         Function that turns off protection for the worksheet. 
        
        #         Arguments:
        #         password - str that represents the sheet protection password 
        #         '''
        #         print("Unprotecting sheet...")
        
        #         if password is None:
        #             print("No password specified, not doing anything")
        #             return 
                
        #         if self.worksheet_name is None: 
        #             print("No worksheet has been specified. Doing nothing...")
        #             return 
        
        #         url = (
        #             f"{GRAPH_URL}/drives/{self.drive_id}"
        #             f"/items/{self.item_id}"
        #             f"/workbook"
        #             f"/worksheets/{self.worksheet_name}"
        #             f"/protection/unprotect"
        #         )
        
        #         body = {
        #             "password": password
        #         }
        
        #         self._post(url, body)
        
        #         print("Unprotect successful!")
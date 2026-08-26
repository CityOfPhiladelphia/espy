from collections.abc import Iterator
from enum import StrEnum
from typing import Any

from graph_api_functions.client import GraphAPIClient

# TODO: Make host name a variable, not a constant. Edit in optional config file?
from graph_api_functions.constants import GRAPH_URL, HOST_NAME
from graph_api_functions.models.graph_api_models import (
    SHARE_POINT_LIST_EXCLUDED_COLUMNS,
)
from graph_api_functions.models.models import (
    GraphAPIResponse,
    HTTPMethod,
    SharePointListColumn,
    SharePointListRow,
)
from graph_api_functions.urls import build_url


class ExcelEndpoints(StrEnum):
    """An endpoint registry for all API operations made by the ExcelWorksheet
    class.

    Args:
        StrEnum (StrEnum): Inherits from the StrEnum class in the enum
        library.
    """
    WORKBOOK_ID = "{graph_url}/drives/{drive_id}/root:/{workbook_path}"

    PROTECT = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/worksheets/{worksheet_name}/protection/protect"
    )
    
    UNPROTECT = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/worksheets/{worksheet_name}/protection/unprotect"
    )

    ADD_ROW = (
        "{graph_url}/drives/{drive_id}"
        "/items/{workbook_id}"
        "/workbook/tables/{table_name}"
        "/rows/add"
    )

    GET_CONTENT = (
        "{graph_url}/sites/{site_id}"
        "/drives/{drive_id}/root:/{file_path}:/content"
    )



class ExcelWorksheet:
    def __init__(
        self,
        client: GraphAPIClient,
        site_id: str,
        drive_id: str, 
        workbook_id: str|None,
        worksheet_name: str|None,
        table_name: str|None,
    ):

        self.client         = client
        self.site_id        = site_id
        self.drive_id       = drive_id
        self.workbook_id    = workbook_id
        self.worksheet_name = worksheet_name
        self.table_name     = table_name

    @classmethod
    def setup(
        cls, 
        hostname:str, 
        site_name:str, 
        document_library:str,
        workbook_path:str|None = None,
        worksheet_name:str|None=None,
        table_name:str|None=None
        ):

        client = GraphAPIClient.authenticate()

        site_id = client.get_site_id(hostname, site_name)
        drive_id = client.get_drive_id(site_id, document_library)

        if workbook_path:
            workbook_id_url = build_url(
                ExcelEndpoints.WORKBOOK_ID,
                graph_url=GRAPH_URL,
                drive_id=drive_id, 
                workbook_path=workbook_path
            )
            workbook_id = client.make_request(HTTPMethod.GET, workbook_id_url)["id"]

            return cls(client, site_id, drive_id, workbook_id, worksheet_name, table_name)
        else:
            return cls(client, site_id, drive_id)

    def append_row(self, row:list, password=None):
        # TODO: Consider append_rows for batching data 
        self.toggle_protection(password, protect=False)

        add_row_url = build_url(
            ExcelEndpoints.ADD_ROW,
            graph_url=GRAPH_URL,
            drive_id=self.drive_id,
            workbook_id=self.workbook_id,
            table_name=self.table_name
        )

        json = {
            "values": row
        }

        self.client.make_request(HTTPMethod.POST, add_row_url, json=json, timeout=60)
        print("Data appended sucessfully!")

        self.toggle_protection(password, protect=True)


    def toggle_protection(self, password:str, protect:bool):
        if not password:
            print("No password specified, not doing anything")
            return 
        
        if not self.worksheet_name:
            print("No worksheet has been specified. Doing nothing...")
            return 

        if not protect:
            url = build_url(
                        ExcelEndpoints.UNPROTECT,
                        graph_url=GRAPH_URL,
                        drive_id=self.drive_id,
                        workbook_id=self.workbook_id,
                        worksheet_name=self.worksheet_name
                    )
        else:
            url = build_url(
                        ExcelEndpoints.PROTECT,
                        graph_url=GRAPH_URL,
                        drive_id=self.drive_id,
                        workbook_id=self.workbook_id,
                        worksheet_name=self.worksheet_name
                    )

        json = { "password": password }

        self.client.make_request(HTTPMethod.POST, url, json=json)

    def get_content(self, file_path:str):

         content_url = build_url(ExcelEndpoints.GET_CONTENT,
                                 graph_url=GRAPH_URL,
                                 site_id=self.site_id,
                                 drive_id=self.drive_id,
                                 file_path=file_path
         )

         request = self.client.make_request(HTTPMethod.GET,
                                            content_url, 
                                            timeout=60, 
                                            follow_redirects=True)

         return request.content 
        


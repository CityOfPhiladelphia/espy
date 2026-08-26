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
    ):

        self.client         = client
        self.site_id        = site_id
        self.drive_id       = drive_id
        self.workbook_id    = workbook_id


    @classmethod
    def setup(
        cls, 
        hostname:str, 
        site_path:str, 
        document_library:str,
        workbook_path:str|None = None ):

        client = GraphAPIClient.authenticate()

        site_id = client.get_site_id(hostname, site_path)
        drive_id = client.get_drive_id(site_id, document_library)

        if workbook_path:
            workbook_id_url = build_url(
                ExcelEndpoints.WORKBOOK_ID,
                graph_url=GRAPH_URL,
                drive_id=drive_id, 
                workbook_path=workbook_path
            )
            workbook_id = client.make_request(HTTPMethod.GET, workbook_id_url)["id"]

            return cls(client, site_id, drive_id, workbook_id)
        else:
            return cls(client, site_id, drive_id)

        

# list.py
from graph_api_functions.client import GraphAPIClient, _URLResolver
from graph_api_functions.constants import HOST_NAME
from typing import Any

class SharePointList:

    def __init__(self, client: GraphAPIClient, site_id: str, list_id: str):
         self.client = client
         self.site_id = site_id
         self.list_id = list_id

    @classmethod
    def get_list(cls, site_path: str, list_name: str):
         client = GraphAPIClient.authenticate()
         site_id = _URLResolver.get_site_id(HOST_NAME, site_path)
         list_id = _URLResolver.get_list_id(site_id, list_name)

         return cls(client, site_id, list_id)

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
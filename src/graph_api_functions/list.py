# list.py
from graph_api_functions.client import GraphAPIClient
from typing import Any

class SharePointList:

    def __init__(self, client: GraphAPIClient, site_id: str, list_id: str):
        ...

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
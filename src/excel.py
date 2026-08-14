# list.py
from connection import GraphAPIClient
from typing import Any

class ExcelTable:

    def __init__(self, GraphAPIClient, site_id: str, 
                 file_path: str, table_name: str):
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

    def toggle_lock(self) -> bool:
         ...

class ExcelWorkBook:
    def __init__(self, client: GraphAPIClient, site_id: str, file_path: str):
        ...

    def list_tables(self) -> list[str]:
        ...

    def get_table(self, table_name: str) -> ExcelTable:
        ...

    def create_table(self, file_path: str, table_name: str) -> ExcelTable:
        ...
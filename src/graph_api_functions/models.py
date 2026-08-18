from enum import StrEnum
from typing import Any, Protocol
from pydantic import BaseModel, HttpUrl, DirectoryPath

## Enums
class HTTPMethod(StrEnum):
    GET = "GET"
    POST = "POST"
    PUT = "PUT"
    PATCH = "PATCH"

## Errors
class UnsupportedMethodError(BaseException):
    pass

## API Input Classes
class SharePointSiteInfo(BaseModel):
    hostname: HttpUrl
    site_path: DirectoryPath


class TabularStorage(Protocol):
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

## API Response Classes

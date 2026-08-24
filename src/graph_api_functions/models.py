from enum import StrEnum
from typing import Any, Protocol

from pydantic import (
    BaseModel,
    ConfigDict,
    DirectoryPath,
    Field,
    HttpUrl
)

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
    def get_row(self, row_id: str) -> dict[str, Any]: ...

    def list_rows(self) -> list[dict[str, Any]]: ...

    def add_row(self, data: dict[str, Any]) -> dict[str, Any]: ...

    def edit_row(self, row_id: str, data: dict[str, Any]) -> dict[str, Any]: ...

    def delete_row(self, row_id: str) -> bool: ...

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...


## API Response Classes
class GraphAPIResponse[T](BaseModel):
    # OData metadata context link
    odata_context: str | None = Field(None, alias="@odata.context")
    next_link: str | None = Field(None, alias="@odata.nextLink")
    value: list[T]
    model_config = ConfigDict(populate_by_name=True)

class SharePointListRow(BaseModel):
    fields: dict = Field(default_factory=dict)

class SharePointListColumn(BaseModel):
    column_group: str = Field(alias="columnGroup")
    description: str | None = Field(None)
    display_name: str = Field(alias="displayName")
    enforce_unique_values: bool = Field(alias="enforceUniqueValues")
    hidden: bool
    id: str
    indexed: bool
    name: str
    readOnly: bool
    required: bool
    text: dict

from enum import StrEnum
from typing import Any, Protocol

from pydantic import BaseModel, ConfigDict, DirectoryPath, Field, HttpUrl

import graph_api_functions.models.graph_api_models as cols


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


### Share Point List Data
class SharePointListRow(BaseModel):
    fields: dict = Field(default_factory=dict)


class SharePointListColumn(cols.SharePointListColumnType):
    description: str | None = Field(None)
    display_name: str
    enforce_unique_values: bool
    hidden: bool
    id: str
    indexed: bool
    name: str
    read_only: bool
    required: bool

    # Share Point List Columns may be one of fourteen types, modeled in
    # list_column_definitions.py
    # When parsing this, you can use the "exclude_unset" option
    # in pydantic's model dump to exclude fields that are none

    boolean: cols.BooleanColumn | None = Field(None)
    calculated: cols.CalculatedColumn | None = Field(None)
    choice: cols.ChoiceColumn | None = Field(None)
    content_approval_status: cols.ContentApprovalStatusColumn | None = Field(
        None
    )
    currency: cols.CurrencyColumn | None = Field(None)
    date_time: cols.DateTimeColumn | None = Field(None)
    geolocation: cols.GeolocationColumn | None = Field(None)
    hyperlink_or_picture: cols.HyperlinkOrPictureColumn | None = Field(None)
    lookup: cols.LookupColumn | None = Field(None)
    number: cols.NumberColumn | None = Field(None)
    person_or_group: cols.PersonOrGroupColumn | None = Field(None)
    term: cols.TermColumn | None = Field(None)
    text: cols.TextColumn | None = Field(None)
    thumbnail: cols.ThumbnailColumn | None = Field(None)

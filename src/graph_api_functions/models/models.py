from enum import StrEnum
from typing import Any, Protocol

from pydantic import BaseModel, ConfigDict, Field

import graph_api_functions.models.graph_api_models as cols


## Enums
class HTTPMethod(StrEnum):
    """HTTP Method Enum. Used to validate HTTP request types
    before making requests.

    Args:
        StrEnum (StrEnum): Inherits the StrEnum parent class
        from the enum library
    """
    GET = "GET"
    POST = "POST"
    PUT = "PUT"
    PATCH = "PATCH"


## Errors
class UnsupportedMethodError(BaseException):
    """Unsupported Method Error. Raised when
    a method against an API is not permitted.

    Args:
        BaseException (BaseException): Inherited from the Base Exception class.
    """
    pass


## API Input Classes
class TabularStorage(Protocol):
    """Protocol defining a contract for what methods any class representing
    a Tabular Storage object (e.g., the SharePointList class) must fulfill. 

    Args:
        Protocol (Protocol): Inherits from the Protocol class in the typing
        module.
    """
    def get_row(self, row_id: str) -> dict[str, Any]: ...

    def list_rows(self) -> list[dict[str, Any]]: ...

    def add_rows(self, data: dict[str, Any]) -> dict[str, Any]: ...

    def edit_row(self, row_id: str, data: dict[str, Any]) -> dict[str, Any]: ...

    def delete_row(self, row_id: str) -> bool: ...

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...


## API Response Classes
class GraphAPIResponse[T](BaseModel):
    """A model representing the data envelope given back by the Microsoft
    Graph API. Data is wrapped in an envelope containing pagination data
    as well as other things. Returned data is in the value field,
    modelled as a list of generic types. Generic type is passed to
    GraphAPIResponse for Pydantic parsing of the records in value.
    """
    # OData metadata context link
    odata_context: str | None = Field(None, alias="@odata.context")
    next_link: str | None = Field(None, alias="@odata.nextLink")
    value: list[T] | None = Field(None)
    fields: dict = Field(default_factory=dict)
    model_config = ConfigDict(populate_by_name=True)


### Share Point List Data
class SharePointListRow(BaseModel):
    """A model representing a row from a SharePoint List.

    Args:
        BaseModel (BaseModel): Inherits from Pydantic's BaseModel class.
    """
    fields: dict = Field(default_factory=dict)


class SharePointListColumn(cols.SharePointListColumnType):
    """A model representing a column of a SharePoint list. Contains
    information about the column and validation rules assigned to the column.

    Args:
        cols (SharePointListColumnType): Inherits from the \
        SharePointListColumnType class, allowing for conversion between \
        camel and snake case field names.
    """
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

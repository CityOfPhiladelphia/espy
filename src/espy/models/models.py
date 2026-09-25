# Models.py

from datetime import datetime
from enum import StrEnum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, model_validator
from pydantic.alias_generators import to_camel

from espy.operations import BatchOperation


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
    DELETE = "DELETE"


class ColumnKind(StrEnum):
    BOOLEAN = "boolean"
    CALCULATED = "calculated"
    CHOICE = "choice"
    CONTENT_APPROVAL_STATUS = "contentApprovalStatus"
    CURRENCY = "currency"
    DATETIME = "dateTime"
    GEOLOCATION = "geolocation"
    HYPERLINK_OR_PICTURE = "hyperlinkOrPicture"
    LOOKUP = "lookup"
    NUMBER = "number"
    PERSON_OR_GROUP = "personOrGroup"
    TERM = "term"
    TEXT = "text"
    THUMBNAIL = "thumbnail"
    UNTYPED = "untyped"


## Mappings
# Column kinds the Graph API does not allow writing to
READ_ONLY_COLUMN_KINDS = frozenset(
    {
        ColumnKind.CALCULATED,
        ColumnKind.GEOLOCATION,
        ColumnKind.HYPERLINK_OR_PICTURE,
        ColumnKind.THUMBNAIL,
        ColumnKind.UNTYPED,
    }
)

# Python types accepted when writing to each writable column kind
ACCEPTABLE_PYTHON_TYPES: dict[ColumnKind, tuple[type, ...]] = {
    ColumnKind.BOOLEAN: (bool,),
    ColumnKind.CHOICE: (
        str,
        list,
    ),
    ColumnKind.CONTENT_APPROVAL_STATUS: (int,),
    ColumnKind.CURRENCY: (float,),
    ColumnKind.DATETIME: (
        datetime,
        str,
    ),
    ColumnKind.LOOKUP: (
        int,
        str,
    ),
    ColumnKind.NUMBER: (
        int,
        float,
    ),
    ColumnKind.PERSON_OR_GROUP: (
        int,
        str,
    ),
    ColumnKind.TERM: (str,),
    ColumnKind.TEXT: (str,),
}


## Errors
class UnsupportedMethodError(Exception):
    """Unsupported Method Error. Raised when
    a method against an API is not permitted.
    """


## API Response Classes
### Validation for data returned by the API
class GraphCollection[T](BaseModel):
    """A model representing the data envelope given back by the Microsoft
    Graph API. Data is wrapped in an envelope containing pagination data
    as well as other things. Returned data is in the value field,
    modelled as a list of generic types. Generic type is passed to
    GraphAPIResponse for Pydantic parsing of the records in value.
    """

    next_link: str | None = Field(None, alias="@odata.nextLink")
    value: list[T] = Field(default_factory=list)
    model_config = ConfigDict(populate_by_name=True)


#### Batch response data, internal models:
class GraphBatchSubResponse(BaseModel):
    id: str
    status: int
    headers: dict
    body: dict[str, Any] | None = None

class GraphBatchResponse(BaseModel):
    responses: list[GraphBatchSubResponse]


#### Batch response data, user-facing models:
class BatchResult(BaseModel):
    operation: BatchOperation
    value: Any = None
    error: BatchError | None = None

class BatchError(BaseModel):
    status: int
    message: str
    code: str | None = None


### Share Point List Data
class SharePointListRow(BaseModel):
    """A model representing a row from a SharePoint List.

    Args:
        BaseModel (BaseModel): Inherits from Pydantic's BaseModel class.
    """

    fields: dict = Field(default_factory=dict)


class SharePointListColumn(BaseModel):
    """A model representing a column of a SharePoint list. Contains
    information about the column and validation rules assigned to the column.
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
    type: ColumnKind

    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True)

    # Share Point List Columns may be one of fourteen types, modeled in
    # graph_api_models.py
    # When parsing this, you can use the "exclude_unset" option
    # in pydantic's model dump to exclude fields that are none

    @model_validator(mode="before")
    @classmethod
    def transform_api_key_to_type(cls, data: dict) -> dict:
        if not isinstance(data, dict):
            raise TypeError("API response not in valid JSON format.")

        # Get the data type based on the name of the field in the
        # incoming data
        # Some column types are not easily identifiable from the
        # response returned from the SharePoint API, including
        # location, which contains hidden sub columns
        # and hyperlink. These fall back to UNTYPED.
        column_kind = next(
            (kind for kind in ColumnKind if kind in data), ColumnKind.UNTYPED
        )

        # Copy rather than mutate the caller's dict
        return {**data, "type": column_kind}

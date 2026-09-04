# Models.py 

from collections.abc import Iterator
from dataclasses import dataclass

from enum import StrEnum 
from datetime import datetime
from typing import Any, Protocol

from pydantic import (
    BaseModel, ConfigDict, 
    Field, model_validator
    )

from pydantic.alias_generators import to_camel

## Enums
class HTTPMethod(StrEnum):
    """HTTP Method Enum. Used to validate HTTP request types
    before making requests.

    Args:
        StrEnum (StrEnum): Inherits the StrEnum parent class
        from the enum library
    """
    GET    = "GET"
    POST   = "POST"
    PUT    = "PUT"
    PATCH  = "PATCH"
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
@dataclass(frozen=True)
class UnsupportedForWrite:
    """Sentinel type indicating a column cannot be written to"""
    pass

COLUMN_KIND_PYTHON_TYPES = {
    ColumnKind.BOOLEAN: (bool,),
    ColumnKind.CALCULATED: (UnsupportedForWrite,),
    ColumnKind.CHOICE: (str, list), # This is one where the choice column is actually helpful, as it might pull in valid values for us, but we can do this later
    ColumnKind.CONTENT_APPROVAL_STATUS: (int,),
    ColumnKind.CURRENCY: (float,),
    ColumnKind.DATETIME: (datetime, str),
    ColumnKind.GEOLOCATION: (UnsupportedForWrite,),
    ColumnKind.HYPERLINK_OR_PICTURE: (UnsupportedForWrite,),
    ColumnKind.LOOKUP: (int, str),
    ColumnKind.NUMBER: (int, float),
    ColumnKind.PERSON_OR_GROUP: (int, str),
    ColumnKind.TERM: (str,),
    ColumnKind.TEXT: (str,),
    ColumnKind.THUMBNAIL: (UnsupportedForWrite,),
    ColumnKind.UNTYPED: (UnsupportedForWrite, )
}

## Errors
class UnsupportedMethodError(BaseException):
    """Unsupported Method Error. Raised when
    a method against an API is not permitted.

    Args:
        BaseException (BaseException): Inherited from the Base Exception class.
    """
    pass

class InvalidIncomingRowError(BaseException):
    pass

class UnsupportedColumnForWriteError(BaseException):
    """
    Raised when a column type does not support write operations.

    Args:
    BaseException (BaseException): Inherited from the Base Exception class.
    """

## API Input Classes

### Validation for Data Coming from SharePoint API

class TabularStorage(Protocol):
    """Protocol defining a contract for what methods any class representing
    a Tabular Storage object (e.g., the SharePointList class) must fulfill. 

    Args:
        Protocol (Protocol): Inherits from the Protocol class in the typing
        module.
    """
    def get_row(self, row_id: str) -> dict[str, Any]: ...

    def list_rows(self) -> list[dict[str, Any]] | Iterator : ...

    def list_columns(self) -> list: ...
        
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
    # list_column_definitions.py
    # When parsing this, you can use the "exclude_unset" option
    # in pydantic's model dump to exclude fields that are none

    @model_validator(mode="before")
    @classmethod
    def transform_api_key_to_type(cls, data: dict) -> dict:
        if not isinstance(data, dict):
            raise ValueError("API response not in valid JSON format.")


        # Get the data type based on the name of the field in the
        # incoming data 
        for data_type in ColumnKind:
            if data_type in data:
                data["type"] = data_type
                return data

        # Some column types are not easily identifiable from the
        # response returned from the SharePoint API, including
        # location, which contains hidden sub columns
        # and hyperlink
        data["type"] = ColumnKind.UNTYPED
        return data

### Validation for data coming from the user
class IncomingField(BaseModel):
    """Checks whether or not the type for an incoming field is a valid
    type, given the type on the Microsoft List"""
    field_name: str
    field_value: Any
    list_column_kind: ColumnKind

    @model_validator(mode='after')
    def check_incoming_data_type_is_valid(self):
        valid_data_types = COLUMN_KIND_PYTHON_TYPES[self.list_column_kind]

        if UnsupportedForWrite in valid_data_types:
            raise ValueError(f"""Field name {self.field_name} has column type
            that does not support write operations with the API: {
                self.list_column_kind.value}
            .""")

        if not isinstance(self.field_value, valid_data_types):
            raise ValueError(f"""Field name {self.field_name} has the wrong
                    data type. Data type must be one of 
                    {",".join(t.__name__ for t in valid_data_types)}""")

        return self
# graph_api_models.py

## NOTE: These models are not currently used by the validation flow
## but may be in the future, if we need more granular error handling
## based on data types and constraints.

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field
from pydantic.alias_generators import to_camel


# ========= Custom errors ===========
class GraphAPILimitationError(Exception):
    """Raised when the graph API does not provide this functionality."""

class MalformedRowError(Exception):
    """Raised when a row of data to insert is malformed."""

class PrimaryKeyValueNotFound(Exception):
    """Raised when searching for a value under the primary key, but it is not found"""

# ========= SharePoint List Column Data Types ===========

class SharePointListColumnType(BaseModel):
    """A parent class for Share Point List Columns. Converts attributes
    named in Share Point's camel case to pythonic snake case.

    Args:
        BaseModel (Pydantic BaseModel): A Pydantic BaseModel class
    """
    # Microsoft returns field names in camel case, so we need to
    # convert. This will allow accessing the field either by
    # snake case or camel case.

    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True)
   
class BooleanColumn(SharePointListColumnType):
    """Boolean column type. Currently empty, per MS documentation.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """


class CalculatedColumn(SharePointListColumnType):
    """Calculated column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    format: str | None = Field(None)
    formula: str | None = Field(None)
    output_type: str | None = Field(None)


class ChoiceColumn(SharePointListColumnType):
    """Choice column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    allow_text_entry: bool | None = Field(None)
    choices: list[str] | None = Field(None)
    display_as: str | None = Field(None)


class ContentApprovalStatusColumn(SharePointListColumnType):
    """Content Approval Status column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """


class CurrencyColumn(SharePointListColumnType):
    """Currency column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    locale: str | None = Field(None)


class DateTimeColumn(SharePointListColumnType):
    """Datetime column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    display_as: str | None = Field(None)
    format: str | None = Field(None)


class LookupColumn(SharePointListColumnType):
    """Lookup column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    allow_multiple_values: bool | None = Field(None)
    allow_unlimited_length: bool | None = Field(None)
    column_name: str | None = Field(None)
    list_id: str | None = Field(None)
    primary_lookup_column_id: str | None = Field(None)


class NumberColumn(SharePointListColumnType):
    """Number column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    decimal_places: (
        Literal["automatic", "none", "one", "two", "three", "four", "five"]
        | None
    ) = Field(None)
    display_as: str | None = Field(None)
    maximum: float | None = Field(None)
    minimum: float | None = Field(None)


class PersonOrGroupColumn(SharePointListColumnType):
    """Person or Group column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    allow_multiple_selection: bool | None = Field(None)
    choose_from_type: str | None = Field(None)
    display_as: str | None = Field(None)


class TermColumn(SharePointListColumnType):
    """Term column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    allow_multiple_values: bool | None = Field(None)
    show_fully_qualified_name: bool | None = Field(None)


class TextColumn(SharePointListColumnType):
    """Text column type.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    allow_multiple_lines: bool | None = Field(None)
    append_changes_to_existing_text: bool | None = Field(None)
    lines_for_editing: int | None = Field(None)
    max_length: int | None = Field(None)
    text_type: str | None = Field(None)


class ThumbnailColumn(SharePointListColumnType):
    """Thumbnail column type. Currently empty, per MS documentation.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """


class GeolocationColumn(SharePointListColumnType):
    """Geolocation column type. Currently empty, per MS documentation.
    The GraphAPI currently does not allow data of this type to be added
    or edited with the Graph API.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """


class HyperlinkOrPictureColumn(SharePointListColumnType):
    """Hyperlink or Picture column type. Currently empty, per MS documentation.
    The GraphAPI currently does not allow data of this type to be added
    or edited with the Graph API.

    Args:
        SharePointListColumnType (Pydantic Model): Inherits the \
        SharePointListColumnType class.
    """
    is_picture: bool | None = Field(None)
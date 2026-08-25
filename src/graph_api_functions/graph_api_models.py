from pydantic import (
    BaseModel,
    ConfigDict,
    Field
)

from pydantic.alias_generators import to_camel

from typing import Literal

SHARE_POINT_LIST_EXCLUDED_COLUMNS = [
    'Title', 
    'Color Tag',
    'Compliance Asset Id',
    'ID',
    'Content Type',
    'Modified',
    'Created',
    'Modified By',
    '_UIVersionString',
    'Attachments',
    'Edit',
    'Type',
    'Item Child Count',
    'Folder Child Count',
    'Label setting',
    'Retention label',
    'Retention label Applied',
    'Label applied by',
    'Item is a Record',
    'App Created By',
    'App Modified By'

]

# ========= Custom errors ===========

class GraphAPILimitationError(Exception):
    """Raised when the graph API does not provide this functionality."""
    pass

# ========= Microsoft Column Classes ===========

class SharePointListColumnType(BaseModel):
    # Microsoft returns field names in camel case, so we need to
    # convert. This will allow accessing the field either by
    # snake case or camel case.

    model_config = ConfigDict(
        alias_generator=to_camel,
        populate_by_name=True)

class BooleanColumn(SharePointListColumnType):
    pass

class CalculatedColumn(SharePointListColumnType):
    format: str | None = Field(None)
    formula: str | None = Field(None)
    output_type: str | None = Field(None)

class ChoiceColumn(SharePointListColumnType):
    allow_text_entry: bool | None = Field(None)
    choices: list[str] | None = Field(None)
    display_as: str | None = Field(None)

class ContentApprovalStatusColumn(SharePointListColumnType):
    pass

class CurrencyColumn(SharePointListColumnType):
    locale: str | None = Field(None)

class DateTimeColumn(SharePointListColumnType):
    display_as: str | None = Field(None)
    format: str | None = Field(None)

class LookupColumn(SharePointListColumnType):
    allow_multiple_values: bool | None = Field(None)
    allow_unlimited_length: bool | None = Field(None)
    column_name: str | None = Field(None)
    list_id: str | None = Field(None)
    primary_lookup_column_id : str | None = Field(None)

class NumberColumn(SharePointListColumnType):
    decimal_places: Literal["automatic", "none", "one", "two", 
                            "three", "four", "five"] | None = Field(None)
    display_as: str | None = Field(None)
    maximum: float | None = Field(None)
    minimum: float | None = Field(None)

class PersonOrGroupColumn(SharePointListColumnType):
    allow_multiple_selection: bool | None = Field(None)
    choose_from_type: str | None = Field(None)
    display_as : str | None = Field(None)

class TermColumn(SharePointListColumnType):
    allow_multiple_values: bool | None = Field(None)
    show_fully_qualified_name: bool | None = Field(None)

class TextColumn(SharePointListColumnType):
    allow_multiple_lines: bool | None = Field(None)
    append_changes_to_existing_text: bool | None = Field(None)
    lines_for_editing: int | None = Field(None)
    max_length: int | None = Field(None)
    text_type: str | None = Field(None)

class ThumbnailColumn(SharePointListColumnType):
    pass


class GeolocationColumn(SharePointListColumnType):
    pass


class HyperlinkOrPictureColumn(SharePointListColumnType):
    is_picture: bool | None = Field(None)


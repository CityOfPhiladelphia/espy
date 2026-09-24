# list.py
from collections.abc import Iterator
from enum import StrEnum
from itertools import batched
from typing import Any, Dict

from httpx import Response
from pydantic import ValidationError
from espy.client import GraphAPIClient

# TODO: Make host name a variable, not a constant. Edit in optional config file?
from espy.constants import (
    GRAPH_URL,
    HOST_NAME,
    SHARE_POINT_LIST_EXCLUDED_COLUMNS,
    BATCH_SIZE
)
from espy.models.models import (
    ColumnKind,
    GraphAPIResponse,
    HTTPMethod,
    IncomingField,
    SharePointListColumn,
    SharePointListRow,
)
from espy.operations import GetRow, AddRow, EditRow, DeleteRow
from espy.urls import build_url
from functools import cached_property

# TODO: Print warning that a list containing a hyperlink or location column
# Cannot be updated with the api


class ListEndpoints(StrEnum):
    """An endpoint registry for all API operations made by the SharePointList
    class.

    Args:
        StrEnum (StrEnum): Inherits from the StrEnum class in the enum
        library.
    """

    LIST_ID = "{graph_url}/sites/{site_id}/lists/{list_name}"

    LIST_ROWS = "{graph_url}/sites/{site_id}/lists/{list_id}/items"

    LIST_COLUMNS = "{graph_url}/sites/{site_id}/lists/{list_id}/columns"

    GET_ROW_BY_ID = (
        "{graph_url}/sites/"
        "{site_id}/lists/{list_id}/items/{row_id}"
    )

    GET_ROW_BY_PK = (
        "/sites/{site_id}/lists/{list_id}/items?"
        "$expand=fields&$filter=fields/{column_name} eq '{value}'"
        )

    ADD_ROW = "/sites/{site_id}/lists/{list_id}/items"

    EDIT_OR_DELETE_ROW = "/sites/{site_id}/lists/{list_id}/items/{row_id}"


class SharePointList:
    """Models a SharePoint List.

    Provides functionality to access, modify, update, and select
    data from a SharePoint list.

    Attributes:
        client (GraphAPIClient): A GraphAPIClient instance, used to
        make requests against the Graph API.
        site_id (str): A string id for the SharePoint site the graph is on.
        list_id (str): A string identifier for the specific SharePoint list.
    """
    def __init__(self, client: GraphAPIClient, site_id: str, list_id: str):
        """Initializes the SharePointList object with id information. Most
        users will initialize the SharePointList object using the setup()
        method.

        Args:
            client (GraphAPIClient): A GraphAPIClient instance, used to
            make requests against the Graph API.
            site_id (str): A string id for the SharePoint site the graph is on.
            list_id (str): A string identifier for the specific SharePoint list.
        """
        self.client = client
        self.site_id = site_id
        self.list_id = list_id

    # Column data
    @cached_property
    def _columns(self) -> list[dict]:
        return self.list_columns()

    @cached_property
    def display_to_canonical(self) -> dict[str, str]:
        return { column['display_name'] : column['name']
                            for column in self._columns }

    @cached_property
    def canonical_to_display(self) -> dict[str, str]:
        return { column['name'] : column['display_name']
                for column in self._columns }

    @cached_property
    def column_types(self) -> dict[str, str]:
        return { column['display_name'] : column['type']
                for column in self._columns }

    @classmethod
    def setup(cls, site_name: str, list_name: str):
        """Authenticates and fetches the ids needed to create the SharePointList
        object.

        Args:
            site_name (str): The site name of the SharePoint list.
            list_name (str): The name of the SharePoint list.

        Returns:
            SharePointList: A SharePointList object, instantiated with client, \
                site_id, and list_id.
        """
        client = GraphAPIClient.authenticate()

        site_id = client.get_site_id(HOST_NAME, site_name)

        list_id_url = build_url(
            ListEndpoints.LIST_ID,
            graph_url=GRAPH_URL,
            site_id=site_id,
            list_name=list_name,
        )

        list_id = client.make_request(HTTPMethod.GET, list_id_url)\
            .json()["id"]

        return cls(client, site_id, list_id)

    def _check_incoming_field_name_valid(self, field_name: str) -> None:

        if field_name not in self.display_to_canonical:
            raise KeyError(f"{field_name} is not in the SharePoint List")
        

    def _check_incoming_field_type_valid(self, field_name: str, 
                                      field_value: Any, 
                                      column_types: dict[str, str]) \
                                        -> IncomingField:

        list_column_kind = column_types[field_name]

        validated_incoming_field = IncomingField(
            field_name=field_name,
            field_value=field_value,
            list_column_kind=ColumnKind(list_column_kind)
        )

        return validated_incoming_field

    def _check_incoming_field_is_pk(self, field_name: str) -> bool:
        for field in self._columns:
            if field['display_name'] == field_name:
                if field.get('indexed') == True:
                    return True

        raise KeyError(f"The specified field does not exist: {field_name}.")
            

    def _validate_incoming_data(
            self, input_data: dict[str, Any]
            ) -> dict[str, IncomingField]:

        validated_fields = {}

        invalid_column_names = []
        invalid_data_types = []
        compiled_errors = []

        for field_name, field_value in input_data.items():
            try:
                self._check_incoming_field_name_valid(field_name)
                validated_field = self._check_incoming_field_type_valid(
                                    field_name, field_value, self.column_types
                                )
                
                validated_fields[field_name] = validated_field

            except KeyError:
                invalid_column_names.append(field_name)

            except ValidationError:
                invalid_data_types.append(field_name)


        if invalid_column_names or invalid_data_types:
            if invalid_column_names:
                compiled_errors.append(f"""The following incoming columns do
                not exist in the SharePointList: 
                {','.join(invalid_column_names)}""")

            if invalid_data_types:
                compiled_errors.append(f"""The following incoming columns
                have the incorrect data type: {','.join(invalid_data_types)}""")

            raise KeyError(f"{'\n'.join(compiled_errors)}")

        return validated_fields

    def _format_payload_for_api(self, data: dict[str, Any]) -> dict[str, dict]:
        """
        Formats an incoming payload to be sent to the API. Validates incoming
        data and maps the user-facing field names to the canonical API names.
        
        Args:
            data (dict[str, Any]): The data to send to the API. May be used
            as downstream part of a POST or PATCH request.

        Returns: dict, a validated and
        formatted payload that the graph API will accept. 
        """
        # Map to the canonical names in the table
        # Make sure all keys are present in SharePointList
        fields_payload = {}
        fields_payload['fields'] = {}

        validated_data = self._validate_incoming_data(data)

        for display_name, column_data in validated_data.items():
            canonical_name = self.display_to_canonical[display_name]
            fields_payload['fields'][canonical_name] = column_data.field_value

        return fields_payload

    def _fetch_page(
        self, url: str | None, params: dict | None
    ) -> GraphAPIResponse[SharePointListRow] | None:
        """A helper funtion to fetch a single page of rows 
        from a SharePoint List. Used by the list_rows method to paginate
        through all data in a SharePoint list.

        Args:
            url (str | None): The URL for the list rows endpoint. Will need
            site_id and list_id. Can be built with the build_url function.
            params (dict): A dict of params to pass to the API call.

        Returns:
            GraphAPIResponse[SharePointListRow] | None: GraphAPI response
            envelope containing a list of SharePoint List row objects.
        """
        # For pagination reasons, return none if no URL is provided
        if not url:
            return None

        raw_data = self.client.make_request(
            HTTPMethod.GET, url, params=params).json()

        response_envelope = GraphAPIResponse[SharePointListRow].model_validate(
            raw_data
        )

        return response_envelope

    def _format_outgoing_row(self, row: dict[str, Any]) -> dict[str, Any]:
        """
        Formats an outgoing row by adding back any null columns not returned
        by the API. The Microsoft Graph API excludes fields that are null,
        resulting in inconsistent data shapes between rows if not corrected.

        Args:
            row: A validated list row returned by the Microsoft Graph API
        
        Returns (dict): A dictionary with null fields added back
        """

        # Includes fields that are null, for standardized output
        formatted_row = {}

        for canonical_column, display_column in \
            self.canonical_to_display.items():
            formatted_row[display_column] = row.get(canonical_column)

        return formatted_row
    
    def _get_row_by_id(self, row_id: int) -> dict[str, Any]:
        """
        Get a single row from a list by list id.

        Args:
            row_id (str): The id of the row to return

        Returns:
            dict: A dictionary with row data
        """
        row_url = build_url(
            ListEndpoints.GET_ROW_BY_ID,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
            row_id=row_id,
            kwargs={"$select": "id"}
        )

        raw_response = self.client.make_request(HTTPMethod.GET, row_url).json()
        validated_response = GraphAPIResponse[SharePointListRow].model_validate(raw_response)
        validated_row = validated_response.fields

        formatted_row = self._format_outgoing_row(validated_row)

        return formatted_row

    def _build_get(self, key_col: str, value: Any) \
        -> tuple[HTTPMethod, str, dict|None]:
        # First, we need to check if the incoming column exists:
        self._check_incoming_field_name_valid(key_col)

        # Then, we need to check if the index column is valid
        self._check_incoming_field_is_pk(key_col)

        canonical_field_name = self.display_to_canonical[key_col]

        method = HTTPMethod.GET

        url = build_url(
            ListEndpoints.GET_ROW_BY_PK,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
            column_name=canonical_field_name,
            value=value
        )

        body = None

        return (method, url, body)

    def _interpret_get(self, raw_response: Response):
        validated_response = GraphAPIResponse.model_validate(raw_response)

        if validated_response.value:
            validated_row = validated_response.value[0]['fields']
            formatted_row = self._format_outgoing_row(validated_row)

            # Add id back, since format outgoing row removes it
            # TODO: Is there a cleaner way around this?
            formatted_row['id'] = validated_row['id']
            return formatted_row

        raise ValueError("Row not found.")

    def _build_add(self, data: Dict) -> tuple[HTTPMethod, str, dict|None]:
        #We need to validate that the incoming data is valid
        self._validate_incoming_data(data)

        method = HTTPMethod.POST

        url = build_url(
            ListEndpoints.ADD_ROW,
            graph_url=GRAPH_URL,
            list_id=self.list_id,
            site_id=self.site_id,
        )

        body = self._format_payload_for_api(data)

        return (method, url, body)

    def _build_edit(self, row_id: int, data: Dict)\
          -> tuple[HTTPMethod, str, dict|None]:

        # We need to validate that the incoming data is valid
        self._validate_incoming_data(data)

        method = HTTPMethod.PATCH

        url = build_url(
            ListEndpoints.EDIT_OR_DELETE_ROW,
            graph_url=GRAPH_URL,
            list_id=self.list_id,
            site_id=self.site_id,
            row_id=row_id
        )

        body = self._format_payload_for_api(data)

        return (method, url, body)

    def _build_delete(self, row_id: int) -> tuple[HTTPMethod, str, dict|None]:

        method = HTTPMethod.DELETE

        url = build_url(
            ListEndpoints.EDIT_OR_DELETE_ROW,
            graph_url=GRAPH_URL,
            list_id=self.list_id,
            site_id=self.site_id,
            row_id=row_id
        )

        body = None

        return (method, url, body)

    def _dispatch_batch_request(self, request: Any)\
          -> tuple[HTTPMethod, str, dict|None]:

        match request:
            case GetRow(key_col, value):
                return self._build_get(key_col, value)
            case AddRow(data):
                return self._build_add(data)
            case EditRow(row_id, data):
                return self._build_edit(row_id, data)
            case DeleteRow(row_id):
                return self._build_delete(row_id)
            case _:
                raise ValueError(f"Unknown request type")

    def list_columns(self) -> list[dict]:
        """List the columns in a SharePoint list.

        Returns:
            list[SharePointListColumn]: A list of SharePointColumn objects.
        """
        columns_url = build_url(
            ListEndpoints.LIST_COLUMNS,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
            kwargs={"$expand": "fields"}
        )

        # Raw columns data will include hidden metadata columns that we want 
        # to exclude, since they will break our data validation rules
        # and are ultimately not useful. We filter those out here.
        column_list = []

        raw_response = self.client.make_request(
            HTTPMethod.GET, columns_url).json()
        
        response_envelope = GraphAPIResponse.model_validate(raw_response)

        if not response_envelope.value:
            raise ValueError("No columns to return.")

        for column in response_envelope.value:
            if column['name'] not in SHARE_POINT_LIST_EXCLUDED_COLUMNS: 
                column_list.append(
                    SharePointListColumn.model_validate(column).model_dump()
                )

        return column_list

    def list_rows(self) -> Iterator[dict]:
        """List all rows in a SharePoint List.

        Paginates through each page of the SharePoint list and returns
        all rows as a generator.

        Yields:
            Iterator[dict]: An iterator of dicts, with each dict representing
            a row of data.
        """
        next_link = build_url(
            ListEndpoints.LIST_ROWS,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
        )

        while response_envelope := self._fetch_page(next_link, {"$expand": "fields"}):
            # API returns results wrapped in a response envelope
            # that contains pagination metadata
            # actual value is nested within that

            next_link = response_envelope.next_link

            # Returns unnecessary extra column information, filter it out
            for row in response_envelope.value: # pyright: ignore
                validated_row = row.fields

                formatted_row = self._format_outgoing_row(validated_row)

                yield formatted_row

    def get_row(self, key_col: str, value: Any) -> dict[str, Any]:
        """
        Get a single row from a list by primary key.

        Args:
            key_col (str): The name of the primary key column to search on
            value (Any): The value of the priamry key column

        Returns:
            dict: A dictionary with row data
        """

        method, url, _ = self._build_get(key_col, value)

        single_request_url = f"{GRAPH_URL}{url}"

        raw_response = self.client.make_request(
            method, single_request_url).json()

        interpreted_response = self._interpret_get(raw_response)

        return interpreted_response
     
    def add_row(self, data: dict[str, Any]) -> Response:
        """Add a row to a SharePoint list.

        Note: This method does not currently work for Lists with
        a Location column, Person column, or Hyperlink or image column.

        Args:
            data (dict[str, Any]): A dict of row data to add. Keys in the dict
            must match the name of the name of the column in the SharePoint
            list.

        Returns:
            dict[str, Any]: A json response object from the API.
        """

        method, url, body = self._build_add(data)

        single_request_url = f"{GRAPH_URL}{url}"

        response = self.client.make_request(
            method, single_request_url, json=body
        )

        return response

    def edit_row(self, key_col: str, value: Any, 
                       data: dict[str, Any]) -> Response:
        """Edit a row in a SharePoint list.

        Note: This method does not currently work for Lists with
        a Location column, Person column, or Hyperlink or image column.

        Args:
            key_col (str): The name of the primary key column to search on.
            value: The value of the primary key column to search on.
            data (dict[str, Any]): A dict of row data to edit. Keys in the dict
            must match the name of the name of the column in the SharePoint
            list.

        Returns:
            dict[str, Any]: A json response object from the API.
        """
        returned_row = self.get_row(key_col, value)

        row_id = returned_row['id']

        method, url, body = self._build_edit(row_id, data)

        single_request_url = f"{GRAPH_URL}{url}"

        response = self.client.make_request(
            method, single_request_url, json=body
        )

        return response

    def delete_row(self, key_col: str, value: Any) -> Response:
        """Delete a row in a SharePoint list.
        Args:
            key_col (str): The name of the primary key column to search on.
            value: The value of the primary key column to search on.

        Returns:
            dict[str, Any]: A json response object from the API.
        """
        returned_row = self.get_row(key_col, value)
        row_id = returned_row['id']

        method, url, _ = self._build_delete(row_id)

        single_request_url = f"{GRAPH_URL}{url}"

        response = self.client.make_request(method, single_request_url)

        return response

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]:
        
        raise NotImplementedError
        # # First, we need to check if the incoming column exists:
        # self._check_incoming_field_name_valid(key_col)

        # # Then, we need to check if the index column is valid
        # self._check_incoming_field_is_pk(key_col)

        # return {}


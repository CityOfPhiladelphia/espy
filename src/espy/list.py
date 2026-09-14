# list.py
from collections.abc import Iterator
from enum import StrEnum
from typing import Any

from httpx import Response
from pydantic import ValidationError

from espy.client import GraphAPIClient

# TODO: Make host name a variable, not a constant. Edit in optional config file?
from espy.constants import (
    GRAPH_URL,
    HOST_NAME,
    SHARE_POINT_LIST_EXCLUDED_COLUMNS,
)
from espy.models.models import (
    ColumnKind,
    GraphAPIResponse,
    HTTPMethod,
    IncomingField,
    InvalidIncomingRowError,
    SharePointListColumn,
    SharePointListRow,
)
from espy.urls import build_url

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
        "{graph_url}/sites/"
        "{site_id}/lists/{list_id}/items?"
        "$expand=fields&$filter=fields/{column_name} eq {value}"
        )

    ADD_ROW = "{graph_url}/sites/{site_id}/lists/{list_id}/items"

    EDIT_ROW = "{graph_url}/sites/{site_id}/lists/{list_id}/items/{row_id}"


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
        self._column_list: list[dict] | None = None
        self._display_to_canonical_column: dict | None = None
        self._canonical_to_display_column: dict | None = None
        self._column_types: dict | None = None 

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

    def _get_column_mapping(self) -> tuple[dict[str, str], dict[str, str]]:
        """
        Gets a column mapping for the list mapping the column's user-visible
        name to the column's real name in the graph API.

        Returns:
            tuple[dict]: A mapping of the user visible name to the real name in the
            graph API, and a mapping of the real name to the user visible name.
        """
        if self._display_to_canonical_column and self._canonical_to_display_column:
            return (self._display_to_canonical_column, 
                    self._canonical_to_display_column
                    )

        columns = self.list_columns()
        
        self._display_to_canonical_column = { column['display_name'] : column['name']
                                for column in columns }

        self._canonical_to_display_column = { column['name'] : column['display_name']
                                             for column in columns }

        return (self._display_to_canonical_column, 
                self._canonical_to_display_column
                )

    def _get_column_types(self) -> dict[str, str]:
        if self._column_types:
            return self._column_types

        columns = self.list_columns()

        self._column_types = { column['display_name'] : column['type']
                                for column in columns }

        return self._column_types

    def _check_incoming_field_name_valid(self, field_name: str, 
                               column_mapping: dict[str, str]) -> None:
    
        if field_name not in column_mapping:
            raise InvalidIncomingRowError(f"{field_name} is not in the SharePoint List")
        

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

    def _check_incoming_field_is_pk(self, field_name: str, 
                                    field_list: list[dict[str, Any]]) -> bool:
        for field in field_list:
            if field['display_name'] == field_name:
                return field.get('indexed', False)

        raise InvalidIncomingRowError(f"The specified field does not exist: {field_name}.")
            

    def _validate_incoming_data(
            self, input_data: dict[str, Any]
            ) -> dict[str, IncomingField]:
        
        display_to_canonical, _ = self._get_column_mapping()
        column_types = self._get_column_types()

        validated_fields = {}

        invalid_column_names = []
        invalid_data_types = []
        compiled_errors = []

        for field_name, field_value in input_data.items():
            try:
                self._check_incoming_field_name_valid(
                    field_name, display_to_canonical
                    )
                validated_field = self._check_incoming_field_type_valid(
                                    field_name, field_value, column_types
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

            raise InvalidIncomingRowError(f"{'\n'.join(compiled_errors)}")

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

        display_to_canonical, _ = self._get_column_mapping()

        validated_data = self._validate_incoming_data(data)

        for display_name, column_data in validated_data.items():
            canonical_name = display_to_canonical[display_name]
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

        _, canonical_to_display = self._get_column_mapping()
        # Includes fields that are null, for standardized output
        formatted_row = {}

        for canonical_column, display_column in canonical_to_display.items():
            formatted_row[display_column] = row.get(canonical_column)

        return formatted_row

        
    def list_columns(self) -> list[dict]:
        """List the columns in a SharePoint list.

        Returns:
            list[SharePointListColumn]: A list of SharePointColumn objects.
        """

        if self._column_list:
            return self._column_list

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

        self._column_list = column_list
        return column_list
    
    def get_row_by_id(self, row_id: int) -> dict[str, Any]:
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

    def get_row_by_pk(self, pk_col: str, value: Any) -> dict[str, Any]:
        """
        Get a single row from a list by primary key.

        Args:
            pk_col (str): The name of the primary key column to search on
            value (Any): The value of the priamry key column

        Returns:
            dict: A dictionary with row data
        """

        display_to_canonical_column, _ = self._get_column_mapping()
        canonical_field_name = display_to_canonical_column[pk_col]

        row_url = build_url(
            ListEndpoints.GET_ROW_BY_PK,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
            column_name=canonical_field_name,
            value=value,
            kwargs={"$select": "id"}
        )

        raw_response = self.client.make_request(HTTPMethod.GET, row_url).json()
        validated_response = GraphAPIResponse.model_validate(raw_response)

        if validated_response.value:
            validated_row = validated_response.value[0]['fields']
            formatted_row = self._format_outgoing_row(validated_row)

            # Add id back, since format outgoing row removes it
            # TODO: Is there a cleaner way around this?
            formatted_row['id'] = validated_row['id']
            return formatted_row

        raise ValueError("Row not found.")

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
        
    def add_row(self, data: dict[str, Any]) -> Response:
        # TODO: Add functionality to make this add rows, adding one or more rows.
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
        add_row_url = build_url(
            ListEndpoints.ADD_ROW,
            graph_url=GRAPH_URL,
            list_id=self.list_id,
            site_id=self.site_id,
        )

        fields_payload = self._format_payload_for_api(data)

        response = self.client.make_request(
            HTTPMethod.POST, add_row_url, json=fields_payload
        )

        return response

    def edit_row_by_pk(self, pk_col: str, value: Any, 
                       data: dict[str, Any]) -> Response:
        """Edit a row in a SharePoint list.

        Note: This method does not currently work for Lists with
        a Location column, Person column, or Hyperlink or image column.

        Args:
            pk_col (str): The name of the primary key column to search on.
            value: The value of the primary key column to search on.
            data (dict[str, Any]): A dict of row data to edit. Keys in the dict
            must match the name of the name of the column in the SharePoint
            list.

        Returns:
            dict[str, Any]: A json response object from the API.
        """
        # First, we need to check if the incoming column exists:
        display_to_canonical, _ = self._get_column_mapping()
        self._check_incoming_field_name_valid(pk_col, display_to_canonical)

        # Then, we need to check if the index column is valid
        list_columns = self.list_columns()
        self._check_incoming_field_is_pk(pk_col, list_columns)

        # Then, we need to validate that the incoming data is valid
        self._validate_incoming_data(data)

        # Then, we need to get the id of the row to edit
        returned_row = self.get_row_by_pk(pk_col, value)
        row_id = returned_row['id']

        edit_row_url = build_url(
            ListEndpoints.EDIT_ROW,
            graph_url=GRAPH_URL,
            list_id=self.list_id,
            site_id=self.site_id,
            row_id=row_id
        )

        fields_payload = self._format_payload_for_api(data)

        response = self.client.make_request(
            HTTPMethod.PATCH, edit_row_url, json=fields_payload
        )

        return response

    def delete_row_by_pk(self, pk_col: str, value: Any) -> Response:
        """Delete a row in a SharePoint list.
        Args:
            pk_col (str): The name of the primary key column to search on.
            value: The value of the primary key column to search on.

        Returns:
            dict[str, Any]: A json response object from the API.
        """
        # First, we need to check if the incoming column exists:
        display_to_canonical, _ = self._get_column_mapping()
        self._check_incoming_field_name_valid(pk_col, display_to_canonical)

        # Then, we need to check if the index column is valid
        list_columns = self.list_columns()
        self._check_incoming_field_is_pk(pk_col, list_columns)

        # Then, we need to get the id of the row to edit
        returned_row = self.get_row_by_pk(pk_col, value)
        row_id = returned_row['id']
    
        delete_row_url = build_url(
            ListEndpoints.EDIT_ROW,
            graph_url=GRAPH_URL,
            list_id=self.list_id,
            site_id=self.site_id,
            row_id=row_id
        )

        response = self.client.make_request(
            HTTPMethod.DELETE, delete_row_url
        )

        return response

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...


if __name__ == "__main__":
    site_name = "311-servicing-department-integrations"
    list_name = "PPR 311 Requests"
    sp_list = SharePointList.setup(site_name=site_name, list_name=list_name)

    for row in sp_list.list_rows():
        print(row)
# list.py
from collections.abc import Iterator
from enum import StrEnum

from espy.client import GraphAPIClient

# TODO: Make host name a variable, not a constant. Edit in optional config file?
from espy.constants import GRAPH_URL, HOST_NAME
from espy.constants import (
    SHARE_POINT_LIST_EXCLUDED_COLUMNS,
)
from espy.models.models import (
    SharePointListColumn,
    GraphAPIResponse,
    HTTPMethod,
    SharePointListRow,
    ColumnKind,
    IncomingField,
    InvalidIncomingRowError
)
from espy.urls import build_url
from httpx import Response
from pydantic import ValidationError
from typing import Any

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

    LIST_ROWS = "{graph_url}/sites/{site_id}/lists/{list_id}/items?"

    LIST_COLUMNS = "{graph_url}/sites/{site_id}/lists/{list_id}/columns?"

    GET_ROW = (
        "{graph_url}/sites/"
        "{site_id}/lists/{list_id}/items/{row_id}?$expand=fields"
    )

    ADD_ROW = "{graph_url}/sites/{site_id}/lists/{list_id}/items"


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
        self._column_mapping: dict | None = None
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

    def get_row(self, row_id: str) -> dict[str, Any]:
        """
        Get a single row from a list.

        Args:
            row_id (str): The id of the row to return

        Returns:
            dict: A dictionary with row data
        """

        row_url = build_url(
            ListEndpoints.GET_ROW,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
            row_id=row_id,
        )

        raw_row_data = self.client.make_request(HTTPMethod.GET, row_url).json()
    
        validated_row_data = GraphAPIResponse\
            .model_validate(raw_row_data)

        return validated_row_data.model_dump()

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

    def _get_column_mapping(self) -> dict[str, str]:
        """
        Gets a column mapping for the list mapping the column's user-visible
        name to the column's real name in the graph API.

        Returns:
            dict: A mapping of the user visible name to the real name in the
            graph API.
        """
        if self._column_mapping:
            return self._column_mapping

        columns = self.list_columns()

        self._column_mapping = { column['display_name'] : column['name']
                                for column in columns }

        return self._column_mapping

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

    def _validate_incoming_data(
            self, input_data: dict[str, Any]
            ) -> dict[str, IncomingField]:
        
        column_mapping = self._get_column_mapping()
        column_types = self._get_column_types()

        validated_fields = {}

        invalid_column_names = []
        invalid_data_types = []
        compiled_errors = []

        for field_name, field_value in input_data.items():
            try:
                self._check_incoming_field_name_valid(
                    field_name, column_mapping
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

        params: dict | None = {"expand": "fields", "list": "fields"}

        while response_envelope := self._fetch_page(next_link, params):
            # API returns results wrapped in a response envelope
            # that contains pagination metadata
            # actual value is nested within that

            params = None
            next_link = response_envelope.next_link
            for row in response_envelope.value: # pyright: ignore
                yield row.model_dump()

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

        # Map to the canonical names in the table
        # Make sure all keys are present in SharePointList
        fields_payload = {}
        fields_payload['fields'] = {}

        self._column_mapping = self._get_column_mapping()

        validated_data = self._validate_incoming_data(data)

        for display_name, column_data in validated_data.items():
            canonical_name = self._column_mapping[display_name]
            fields_payload['fields'][canonical_name] = column_data.field_value

        response = self.client.make_request(
            HTTPMethod.POST, add_row_url, json=fields_payload
        )

        return response

    def edit_row(self, row_id: str, data: dict[str, Any]) -> dict[str, Any]: ...

    def delete_row(self, row_id: str) -> bool: ...

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...


if __name__ == "__main__":
    site_name = "ps360-metrics-share"
    list_name = "testing_lists"
    sp_list = SharePointList.setup(site_name=site_name, list_name=list_name)
    result = sp_list.list_columns()
    print(result)
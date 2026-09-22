# list.py
from collections.abc import Iterator
from enum import StrEnum
from typing import Any

from httpx import Response
from pydantic import ValidationError
from requests.exceptions import HTTPError

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
    IncomingBatch,
    IncomingRequest,
    Batch,
    BatchResult,
    BatchResponseBody,
    ErrorResult,
    InvalidIncomingRowError,
    SharePointListColumn,
    SharePointListRow,
)
from espy.urls import build_url
from functools import cached_property

# TODO: Print warning that a list containing a hyperlink or location column
# Cannot be updated with the api


class ListEndpoint(StrEnum):
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
        "sites/{site_id}/lists/{list_id}/items?"
        "$expand=fields&$filter=fields/{column_name} eq {value}"
        )

    ADD_ROW = "/sites/{site_id}/lists/{list_id}/items"

    EDIT_OR_DELETE_ROW = "/sites/{site_id}/lists/{list_id}/items/{row_id}"

    SUBMIT_BATCH = "https://graph.microsoft.com/v1.0/$batch"

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
    def _columns(self) -> list[SharePointListColumn]:
        return self.list_columns()

    @cached_property
    def display_to_canonical(self) -> dict[str, str]:
        return { column.display_name : column.name
                            for column in self._columns }

    @cached_property
    def canonical_to_display(self) -> dict[str, str]:
        return { column.name : column.display_name
                for column in self._columns }

    @cached_property
    def column_types(self) -> dict[str, str]:
        return {column.display_name: column.type
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
            ListEndpoint.LIST_ID,
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
            if field.display_name == field_name:
                if field.indexed == True:
                    return True

        raise InvalidIncomingRowError(f"The specified field is not a primary key: {field_name}.")
            

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

            raise InvalidIncomingRowError(f"{'\n'.join(compiled_errors)}")

        return validated_fields

    def _batch_list(self, data: list, size: int = 20) -> Iterator:
        """Yield successive batches of up to 20 items from a list. 
        SharePoint Graph API can only handle 20 requests at a time."""

        for i in range(0, len(data), size):
            yield data[i : i + size]

    def _format_incoming_row(self, data: dict[str, Any]) -> SharePointListRow:
        """
        Formats an incoming row to be sent to the API. Validates incoming
        data and maps the user-facing field names to the canonical API names.
        
        Args:
            data (dict[str, Any]): The data to send to the API. May be used
            as downstream part of a POST or PATCH request.

        Returns: dict, a validated and
        formatted payload that the graph API will accept. 
        """
        # Map to the canonical names in the table
        # Make sure all keys are present in SharePointList
        fields_payload = SharePointListRow()

        validated_data = self._validate_incoming_data(data)

        for display_name, column_data in validated_data.items():
            canonical_name = self.display_to_canonical[display_name]
            fields_payload.fields[canonical_name] = column_data.field_value

        return fields_payload

    def _fetch_page(
        self, url: str | None, params: dict | None
    ) -> GraphAPIResponse | None:
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

        response_envelope = GraphAPIResponse.model_validate(
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

        
    def list_columns(self) -> list[SharePointListColumn]:
        """List the columns in a SharePoint list.

        Returns:
            list[SharePointListColumn]: A list of SharePointColumn objects.
        """
        columns_url = build_url(
            ListEndpoint.LIST_COLUMNS,
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
                    SharePointListColumn.model_validate(column)
                )

        self._column_list = column_list
        return column_list
    
    def _build_url_and_body_from_method(self, method: HTTPMethod, item: Any,
                                         **kwargs) -> tuple[str, SharePointListRow | None]:
        """
        Creates a correctly formatted url to the API based on the HTTP method.
        Args:
            method (HTTPMethod): The HTTP method being used
            item (Any): An item that will be passed to the API as part of a request body. 
            Will be returned if the method requires it.
        """
        match method:
            # Get Rows
            case HTTPMethod.GET:
                url = build_url(
                ListEndpoint.GET_ROW_BY_PK, 
                site_id = self.site_id,
                list_id = self.list_id,
                column_name=kwargs.get("key_col"),
                value=item, 
                kwargs={"$select": "id"}
                )

                return url, None

            # Add Rows
            case HTTPMethod.POST:
                url = build_url(
                    ListEndpoint.ADD_ROW,
                    list_id=self.list_id,
                    site_id=self.site_id,
                )

                return url, item

            # In the case of an edit or delete request,
            # we need to pass an item containing an "id"

            # Edit Rows
            case HTTPMethod.PATCH:
                url = build_url(
                ListEndpoint.EDIT_OR_DELETE_ROW, 
                site_id = self.site_id,
                list_id = self.list_id,
                row_id = item['fields']['id']
                )

                return url, item

            # Delete Rows
            case HTTPMethod.DELETE:
                url = build_url(
                ListEndpoint.EDIT_OR_DELETE_ROW, 
                site_id = self.site_id,
                list_id = self.list_id,
                row_id = item['fields']['id']
                )

                return url, None

        raise ValueError(f"Cannot build URL. Unsupported method: {method}")

    def _create_request_batch(self, method: HTTPMethod, batch: list[Any], 
                              **kwargs):
        """
        Creates a batch of requests for the API's batch endpoint. Determines
        how to format those requests based on the method provided.
        """

        request_batch = []
        for idx, item in enumerate(batch):

            if method == HTTPMethod.POST:
                formatted_item = self._format_incoming_row(item)

            else:
                formatted_item = item 

            url, body = self._build_url_and_body_from_method(
                method, formatted_item, **kwargs
            )

            request_batch.append(
                IncomingRequest(
                    id=str(idx),
                    method=method,
                    url=url,
                    body=body
                )
            )

        return IncomingBatch(requests=request_batch)

    def _parse_batch_request_response(self, batch_response: Response)\
          -> BatchResult[SharePointListRow]:
        """
        For the response from each request in a batch request, parse
        that a request and return whether or not it succeeded or failed.

        Args:
            batch: The batch that was processed.
            batch_response: The response from a batch request to the SharePoint API.
        
        Returns:
            BatchResult: An object containing information about batch request
            successes and failures.
        
        """
        batch_result = BatchResult()
        
        batch_responses = Batch.model_validate(batch_response.json())

        for response in batch_responses.responses:
            response_body = response.body

            if isinstance(response_body, ErrorResult):
                raise HTTPError(f"{response_body.error.code}: {response_body.error.message}")

            elif isinstance(response_body, BatchResponseBody):
                for row in response_body.value:
                    batch_result.responses.append(row)

            elif isinstance(response_body, SharePointListRow):
                batch_result.responses.append(response_body)

            else:
                raise ValueError("Data returned by API in unparseable format.")
            
        return batch_result

    def get_rows(self, key_col: str, values: list[Any])\
         -> Iterator[dict[str, Any]]:
        """
        Get rows by primary key.

        Args:
            key_col (str): The name of the primary key column to search on
            value list[any]: The primary key values to search for.

        Returns:
            list[dict]: A dictionary with row data
        """
        # Check if the incoming column exists
        self._check_incoming_field_name_valid(key_col)

        # Check if the index column is valid
        self._check_incoming_field_is_pk(key_col)

        canonical_key_name = self.display_to_canonical.get(key_col)

        for batch in self._batch_list(values):
        # For each row that we want to get, we need to construct
        # a separate get request URL
            payload = self._create_request_batch(
                HTTPMethod.GET,
                batch,
                key_col=canonical_key_name
                )
            
            response = self.client.make_request(
                HTTPMethod.POST,
                ListEndpoint.SUBMIT_BATCH,
                json=payload.model_dump()
            )

            response.raise_for_status()

            batch_response = self._parse_batch_request_response(response)

            if not batch_response.responses:
                raise ValueError(
                    f"No rows found where '{key_col}' matches the given values."
                )

            for response in batch_response.responses:
                validated_response = response.fields
                formatted_row = self._format_outgoing_row(validated_response)
                yield formatted_row


    def list_rows(self) -> Iterator[dict[str, Any]]:
        """List all rows in a SharePoint List.

        Paginates through each page of the SharePoint list and returns
        all rows as a generator.

        Yields:
            Iterator[dict]: An iterator of dicts, with each dict representing
            a row of data.
        """
        next_link = build_url(
            ListEndpoint.LIST_ROWS,
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
                validated_row = row['fields']
                formatted_row = self._format_outgoing_row(validated_row)
                yield formatted_row
        
    def add_rows(self, data: list[dict[str, Any]]) -> Iterator[dict[str, Any]]:
        """Add a row to a SharePoint list.

        Note: This method does not currently work for Lists with
        a Location column, Person column, or Hyperlink or image column.

        Args:
            data (list[dict[str, Any]]): A list of dicts of row data to add. 
            Keys in the dicts  must match the name of the name of the column in the SharePoint
            list.

        Returns:
            list[Response] -> A list of response objects from the API.
        """
        for batch in self._batch_list(data):
            # For each row to add, we need to construct a
            # separate request to make.
            payload = self._create_request_batch(
                HTTPMethod.POST, batch
            )

            response = self.client.make_request(
                HTTPMethod.POST,
                ListEndpoint.SUBMIT_BATCH,
                json=payload.model_dump()
            )

            response.raise_for_status()
            batch_response = self._parse_batch_request_response(response)

            for response in batch_response.responses:
                validated_response = response.fields
                formatted_row = self._format_outgoing_row(validated_response)
                yield formatted_row
    
    # def edit_row(self, key_col: str, value: Any, 
    #                    data: dict[str, Any]) -> Response:
    #     """Edit a row in a SharePoint list.

    #     Note: This method does not currently work for Lists with
    #     a Location column, Person column, or Hyperlink or image column.

    #     Args:
    #         key_col (str): The name of the primary key column to search on.
    #         value: The value of the primary key column to search on.
    #         data (dict[str, Any]): A dict of row data to edit. Keys in the dict
    #         must match the name of the name of the column in the SharePoint
    #         list.

    #     Returns:
    #         dict[str, Any]: A json response object from the API.
    #     """
    #     # First, we need to check if the incoming column exists:
    #     self._check_incoming_field_name_valid(key_col, 
    #                                           self.display_to_canonical)

    #     # Then, we need to check if the index column is valid
    #     list_columns = self.list_columns()
    #     self._check_incoming_field_is_pk(key_col, list_columns)

    #     # Then, we need to validate that the incoming data is valid
    #     self._validate_incoming_data(data)

    #     # Then, we need to get the id of the row to edit
    #     returned_row = self.get_rows(key_col, value)
    #     row_id = returned_row['id']

    #     edit_row_url = build_url(
    #         ListEndpoint.EDIT_ROW,
    #         graph_url=GRAPH_URL,
    #         list_id=self.list_id,
    #         site_id=self.site_id,
    #         row_id=row_id
    #     )

    #     fields_payload = self._format_row_for_api(data)

    #     response = self.client.make_request(
    #         HTTPMethod.PATCH, edit_row_url, json=fields_payload
    #     )

    #     return response

    # def delete_row(self, key_col: str, value: Any) -> Response:
    #     """Delete a row in a SharePoint list.
    #     Args:
    #         key_col (str): The name of the primary key column to search on.
    #         value: The value of the primary key column to search on.

    #     Returns:
    #         dict[str, Any]: A json response object from the API.
    #     """
    #     # First, we need to check if the incoming column exists:
    #     self._check_incoming_field_name_valid(key_col, self.display_to_canonical)

    #     # Then, we need to check if the index column is valid
    #     list_columns = self.list_columns()
    #     self._check_incoming_field_is_pk(key_col, list_columns)

    #     # Then, we need to get the id of the row to edit
    #     returned_row = self.get_rows(key_col, value)
    #     row_id = returned_row['id']
    
    #     delete_row_url = build_url(
    #         ListEndpoint.EDIT_ROW,
    #         graph_url=GRAPH_URL,
    #         list_id=self.list_id,
    #         site_id=self.site_id,
    #         row_id=row_id
    #     )

    #     response = self.client.make_request(
    #         HTTPMethod.DELETE, delete_row_url
    #     )

    #     return response

    # def upsert_row(
    #     self, key_col: str, data: dict[str, Any]
    # ) -> dict[str, Any]:
    #             # First, we need to check if the incoming column exists:
    #     self._check_incoming_field_name_valid(key_col, self.display_to_canonical)

    #     # Then, we need to check if the index column is valid
    #     list_columns = self.list_columns()
    #     self._check_incoming_field_is_pk(key_col, list_columns)

    #     return {}
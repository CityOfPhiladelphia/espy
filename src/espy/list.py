# list.py
from collections.abc import Callable, Iterator, Sequence
from enum import StrEnum
from functools import cached_property
from itertools import batched
from typing import Any

from pydantic import ValidationError

from espy.client import GraphAPIClient, resolve_hostname
from espy.constants import (
    GRAPH_URL,
    MAX_BATCH_SIZE,
    SHAREPOINT_LIST_EXCLUDED_COLUMNS,
)
from espy.models.models import (
    ACCEPTABLE_PYTHON_TYPES,
    READ_ONLY_COLUMN_KINDS,
    APICredentials,
    BatchError,
    BatchResult,
    ColumnKind,
    GraphBatchResponse,
    GraphBatchSubResponse,
    GraphCollection,
    HTTPMethod,
    SharePointListColumn,
    SharePointListRow,
)
from espy.operations import AddRow, BatchOperation, DeleteRow, EditRow, GetRow
from espy.urls import build_url

# TODO: Print warning that a list containing a hyperlink or location column
# Cannot be updated with the api

# The (method, url, body) parts of a single Graph request.
type RequestParts = tuple[HTTPMethod, str, dict | None]

# Turns a successful response body into the value returned to the user.
# Each operation type has its own, e.g. _interpret_get for GetRow.
type Interpreter = Callable[[dict | None], Any]


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

    GET_ROW_BY_ID = "{graph_url}/sites/{site_id}/lists/{list_id}/items/{row_id}"

    # Exclude graph url from these as they may be needed for batch requests
    # which take a different base URL
    GET_ROW_BY_PK = (
        "sites/{site_id}/lists/{list_id}/items?"
        "$expand=fields&$filter=fields/{column_name} eq '{value}'"
    )

    ADD_ROW = "sites/{site_id}/lists/{list_id}/items"

    EDIT_OR_DELETE_ROW = "sites/{site_id}/lists/{list_id}/items/{row_id}"

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
        return {column.display_name: column.name for column in self._columns}

    @cached_property
    def canonical_to_display(self) -> dict[str, str]:
        return {column.name: column.display_name for column in self._columns}

    @cached_property
    def column_types(self) -> dict[str, ColumnKind]:
        return {column.display_name: column.type for column in self._columns}

    @classmethod
    def setup(
        cls,
        *,
        site_name: str,
        list_name: str,
        hostname: str | None = None,
        creds: APICredentials | None = None,
    ):
        """Authenticates and fetches the ids needed to create the SharePointList
        object. All arguments are keyword-only.

        Args:
            site_name (str): The site name of the SharePoint list.
            list_name (str): The name of the SharePoint list.
            hostname (str | None, optional): SharePoint hostname, i.e.
            "example.sharepoint.com". If None, read from SHAREPOINT_HOSTNAME.
            creds (APICredentials | None, optional): tenant_id, client_id and
            client_secret. If None, read from AZURE_TENANT_ID, AZURE_CLIENT_ID
            and AZURE_CLIENT_SECRET.

        Returns:
            SharePointList: A SharePointList object, instantiated with client, \
                site_id, and list_id.
        """
        client = GraphAPIClient.authenticate(creds)

        site_id = client.get_site_id(resolve_hostname(hostname), site_name)

        list_id_url = build_url(
            ListEndpoints.LIST_ID,
            graph_url=GRAPH_URL,
            site_id=site_id,
            list_name=list_name,
        )

        list_id = client.make_request(HTTPMethod.GET, list_id_url).json()["id"]

        return cls(client, site_id, list_id)

    def _check_incoming_field_name_valid(self, field_name: str) -> None:

        if field_name not in self.display_to_canonical:
            raise KeyError(f"{field_name} is not in the SharePoint List")

    def _check_incoming_field_type_valid(
        self, field_name: str, field_value: Any
    ) -> None:

        column_kind = self.column_types[field_name]

        if column_kind in READ_ONLY_COLUMN_KINDS:
            raise TypeError(
                f"Field name {field_name} has column type that does not "
                f"support write operations with the API: {column_kind.value}."
            )

        valid_data_types = ACCEPTABLE_PYTHON_TYPES[column_kind]

        # Null values should not be flagged as a bad data type
        if field_value is not None and not isinstance(
            field_value, valid_data_types
        ):
            raise TypeError(
                f"Field name {field_name} has the wrong data type. Data type "
                f"must be one of "
                f"{','.join(t.__name__ for t in valid_data_types)}"
            )

    def _check_incoming_field_is_pk(self, field_name: str) -> bool:
        for column in self._columns:
            if column.display_name == field_name and column.indexed:
                return True

        raise KeyError(f"The specified field is not indexed: {field_name}.")

    def _validate_incoming_data(self, input_data: dict[str, Any]) -> None:

        invalid_column_names = []
        invalid_data_types = []
        compiled_errors = []

        for field_name, field_value in input_data.items():
            try:
                self._check_incoming_field_name_valid(field_name)
                self._check_incoming_field_type_valid(field_name, field_value)

            except KeyError:
                invalid_column_names.append(field_name)

            except TypeError:
                invalid_data_types.append(field_name)

        if invalid_column_names or invalid_data_types:
            if invalid_column_names:
                compiled_errors.append(
                    f"""The following incoming columns do not exist in the SharePointList: {",".join(invalid_column_names)}"""
                )

            if invalid_data_types:
                compiled_errors.append(
                    f"""The following incoming columns have the incorrect data type: {",".join(invalid_data_types)}"""
                )

            raise KeyError(f"{'\n'.join(compiled_errors)}")

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
        self._validate_incoming_data(data)

        # Map to the canonical names in the table
        return {
            "fields": {
                self.display_to_canonical[display_name]: value
                for display_name, value in data.items()
            }
        }

    def _fetch_page(
        self, url: str | None, params: dict | None
    ) -> GraphCollection[SharePointListRow] | None:
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
            HTTPMethod.GET, url, params=params
        ).json()

        response_envelope = GraphCollection[SharePointListRow].model_validate(
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

        for (
            canonical_column,
            display_column,
        ) in self.canonical_to_display.items():
            formatted_row[display_column] = row.get(canonical_column)

        # Add id back
        formatted_row["id"] = row["id"]

        return formatted_row

    def _build_get(
        self, key_col: str, value: Any
    ) -> tuple[HTTPMethod, str, dict | None]:
        # First, we need to check if the incoming column exists:
        self._check_incoming_field_name_valid(key_col)

        # Then, we need to check if the index column is valid
        self._check_incoming_field_is_pk(key_col)

        canonical_field_name = self.display_to_canonical[key_col]

        method = HTTPMethod.GET

        url = build_url(
            ListEndpoints.GET_ROW_BY_PK,
            site_id=self.site_id,
            list_id=self.list_id,
            column_name=canonical_field_name,
            value=value,
        )

        body = None

        return (method, url, body)

    def _interpret_get(self, body: dict | None) -> dict:
        rows = (
            GraphCollection[SharePointListRow].model_validate(body or {}).value
        )

        if not rows:
            raise ValueError("Row not found.")

        return self._format_outgoing_row(rows[0].fields)

    def _build_add(self, data: dict) -> tuple[HTTPMethod, str, dict | None]:
        method = HTTPMethod.POST

        url = build_url(
            ListEndpoints.ADD_ROW,
            list_id=self.list_id,
            site_id=self.site_id,
        )

        body = self._format_payload_for_api(data)

        return (method, url, body)

    def _interpret_add(self, body: dict | None) -> dict:
        validated_row = SharePointListRow.model_validate(body).fields
        formatted_row = self._format_outgoing_row(validated_row)

        return formatted_row

    def _build_edit(
        self, row_id: int, data: dict
    ) -> tuple[HTTPMethod, str, dict | None]:

        method = HTTPMethod.PATCH

        url = build_url(
            ListEndpoints.EDIT_OR_DELETE_ROW,
            list_id=self.list_id,
            site_id=self.site_id,
            row_id=row_id,
        )

        body = self._format_payload_for_api(data)

        return (method, url, body)

    def _interpret_edit(self, body: dict | None) -> dict:
        validated_row = SharePointListRow.model_validate(body).fields
        formatted_row = self._format_outgoing_row(validated_row)

        return formatted_row

    def _build_delete(self, row_id: int) -> tuple[HTTPMethod, str, dict | None]:

        method = HTTPMethod.DELETE

        url = build_url(
            ListEndpoints.EDIT_OR_DELETE_ROW,
            list_id=self.list_id,
            site_id=self.site_id,
            row_id=row_id,
        )

        body = None

        return (method, url, body)

    def _interpret_delete(self, body: dict | None) -> dict | None:
        return body

    def _dispatch_batch_operation(
        self, operation: BatchOperation
    ) -> tuple[RequestParts, Interpreter]:

        match operation:
            case GetRow(key_col, value):
                return self._build_get(key_col, value), self._interpret_get
            case AddRow(data):
                return self._build_add(data), self._interpret_add
            case EditRow(row_id, data):
                return self._build_edit(row_id, data), self._interpret_edit
            case DeleteRow(row_id):
                return self._build_delete(row_id), self._interpret_delete
            case _:
                raise ValueError("Unknown request type")

    def _build_batch(
        self, offset: int, operations: Sequence[BatchOperation]
    ) -> tuple[list[dict], list[Interpreter]]:
        """Build the Graph request bodies for one group of operations.

        Request ids start at ``offset`` so they stay unique across groups.
        ``interpreters[i]`` is the interpreter for ``operations[i]``.
        """

        requests: list[dict] = []
        interpreters: list[Interpreter] = []

        for idx, operation in enumerate(operations):
            (method, url, body), interpreter = self._dispatch_batch_operation(
                operation
            )

            requests.append(
                {
                    "id": str(offset + idx),
                    "method": method,
                    "url": url,
                    "headers": {"Content-Type": "application/json"},
                    "body": body,
                }
            )

            interpreters.append(interpreter)

        return requests, interpreters

    def _to_batch_result(
        self,
        operation: BatchOperation,
        sub_response: GraphBatchSubResponse | None,
        interpreter: Interpreter,
    ) -> BatchResult:
        """Turn one Graph sub-response into a BatchResult.

        Never raises: a missing response, an HTTP error, or a body the
        interpreter can't parse all become a BatchResult with ``error`` set.
        """
        if sub_response is None:
            return BatchResult(
                operation=operation,
                error=BatchError(
                    status=None, message="No response returned by Graph"
                ),
            )

        if not 200 <= sub_response.status < 300:
            err = (sub_response.body or {}).get("error", {})
            return BatchResult(
                operation=operation,
                error=BatchError(
                    status=sub_response.status,
                    message=err.get("message", ""),
                    code=err.get("code"),
                ),
            )

        try:
            return BatchResult(
                operation=operation, value=interpreter(sub_response.body)
            )
        except (ValueError, ValidationError) as e:
            return BatchResult(
                operation=operation,
                error=BatchError(status=sub_response.status, message=str(e)),
            )

    def list_columns(self) -> list[SharePointListColumn]:
        """List the columns in a SharePoint list.

        Returns:
            list[SharePointListColumn]: A list of SharePointColumn objects.
        """
        columns_url = build_url(
            ListEndpoints.LIST_COLUMNS,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
            kwargs={"$expand": "fields"},
        )

        # Raw columns data will include hidden metadata columns that we want
        # to exclude, since they will break our data validation rules
        # and are ultimately not useful. We filter those out here.
        column_list = []

        raw_response = self.client.make_request(
            HTTPMethod.GET, columns_url
        ).json()

        response_envelope = GraphCollection[dict].model_validate(raw_response)

        if not response_envelope.value:
            raise ValueError("No columns to return.")

        for column in response_envelope.value:
            if column["name"] not in SHAREPOINT_LIST_EXCLUDED_COLUMNS:
                column_list.append(SharePointListColumn.model_validate(column))

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

        # Only the first request needs explicit params; every @odata.nextLink
        # already embeds the full query string (including $skiptoken), and httpx
        # replaces a URL's query string entirely when params is passed.
        params = {"$expand": "fields"}

        while response_envelope := self._fetch_page(next_link, params):
            next_link = response_envelope.next_link
            params = None

            for row in response_envelope.value:  # pyright: ignore
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

        single_request_url = f"{GRAPH_URL}/{url}"

        response = self.client.make_request(method, single_request_url)

        return self._interpret_get(response.json())

    def add_row(self, data: dict[str, Any]) -> dict:
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

        # For a single request, we need to add the GRAPH URL on
        # the URL returned without it is suitable for batch requests
        single_request_url = f"{GRAPH_URL}/{url}"

        response = self.client.make_request(
            method, single_request_url, json=body
        )

        return self._interpret_add(response.json())

    def edit_row(self, key_col: str, value: Any, data: dict[str, Any]) -> dict:
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

        row_id = returned_row["id"]

        method, url, body = self._build_edit(row_id, data)

        single_request_url = f"{GRAPH_URL}/{url}"

        response = self.client.make_request(
            method, single_request_url, json=body
        )

        return self._interpret_edit(response.json())

    def delete_row(self, key_col: str, value: Any) -> dict | None:
        """Delete a row in a SharePoint list.
        Args:
            key_col (str): The name of the primary key column to search on.
            value: The value of the primary key column to search on.

        Returns:
            dict[str, Any]: A json response object from the API.
        """
        returned_row = self.get_row(key_col, value)
        row_id = returned_row["id"]

        method, url, _ = self._build_delete(row_id)

        single_request_url = f"{GRAPH_URL}/{url}"

        response = self.client.make_request(method, single_request_url)

        return self._interpret_delete(response.json())

    def upsert_row(self, key_col: str, data: dict[str, Any]) -> dict[str, Any]:

        raise NotImplementedError
        # # First, we need to check if the incoming column exists:
        # self._check_incoming_field_name_valid(key_col)

        # # Then, we need to check if the index column is valid
        # self._check_incoming_field_is_pk(key_col)

        # return {}

    def batch(self, operations: Sequence[BatchOperation]) -> list[BatchResult]:
        """
        Given a list of user supplied BatchOperations, collate
        and process the operations as a batch.
        """

        results: list[BatchResult] = []

        # Graph accepts at most MAX_BATCH_SIZE requests per batch call
        for group_num, group in enumerate(batched(operations, MAX_BATCH_SIZE)):
            offset = group_num * MAX_BATCH_SIZE
            requests, interpreters = self._build_batch(offset, group)

            response = self.client.make_request(
                HTTPMethod.POST,
                ListEndpoints.SUBMIT_BATCH,
                json={"requests": requests},
            )
            envelope = GraphBatchResponse.model_validate(response.json())

            # Graph may return sub-responses in any order, so look them up
            # by request id. This is not the row id, which is also called
            # 'id' but lives inside each response body.
            by_id = {int(sub.id): sub for sub in envelope.responses}

            # Walk the group in input order so results line up with operations
            for idx, (operation, interpreter) in enumerate(
                zip(group, interpreters, strict=True)
            ):
                sub_response = by_id.get(offset + idx)
                results.append(
                    self._to_batch_result(operation, sub_response, interpreter)
                )

        return results

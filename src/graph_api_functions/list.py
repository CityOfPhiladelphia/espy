# list.py
from collections.abc import Iterator
from enum import StrEnum
from typing import Any

from graph_api_functions.client import ClientEndpoints, GraphAPIClient

# TODO: Make host name a variable, not a constant. Edit in optional config file?
from graph_api_functions.constants import GRAPH_URL, HOST_NAME
from graph_api_functions.models.graph_api_models import (
    SHARE_POINT_LIST_EXCLUDED_COLUMNS,
)
from graph_api_functions.models.models import (
    GraphAPIResponse,
    HTTPMethod,
    SharePointListColumn,
    SharePointListRow,
)
from graph_api_functions.urls import build_url

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
        users will initialize the SharePointList object using the get_list()
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

    @classmethod
    def get_list(cls, site_path: str, list_name: str):
        """Authenticates and fetches the ids needed to create the SharePointList
        object.

        Args:
            site_path (str): The site path of the SharePoint list.
            list_name (str): The name of the SharePoint list.

        Returns:
            SharePointList: A SharePointList object, instantiated with client,
            site_id, and list_id.
        """
        # TODO: New method to get site and list ID
        client = GraphAPIClient.authenticate()

        site_id_url = build_url(
            ClientEndpoints.SITE_ID,
            graph_url=GRAPH_URL,
            hostname=HOST_NAME,
            site_path=site_path,
        )

        site_id = client.make_request("GET", site_id_url)["id"]

        site_id = client.get_site_id()

        list_id_url = build_url(
            ListEndpoints.LIST_ID,
            graph_url=GRAPH_URL,
            site_id=site_id,
            list_name=list_name,
        )

        list_id = client.make_request("GET", list_id_url)["id"]

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

        raw_row_data = self.client.make_request(HTTPMethod.GET, row_url)

        return raw_row_data
        # validated_row_data = SharePointListRow(raw_row_data)

        # return validated_row_data.model_dump()

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
        )

        raw_columns_data = self.client.make_request(HTTPMethod.GET, columns_url)

        response_envelope = GraphAPIResponse[
            SharePointListColumn
        ].model_validate(raw_columns_data)

        filtered = [
            column
            for column in response_envelope.value
            if not column.read_only
            and column.display_name not in SHARE_POINT_LIST_EXCLUDED_COLUMNS
        ]

        return filtered

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

        column_mapping = {}

        columns = self.list_columns()

        for column in columns:
            column_mapping[column.display_name] = column.name

        self._column_mapping = column_mapping

        return column_mapping

    def _fetch_page(
        self, url: str | None, params: dict
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

        raw_data = self.client.make_request(HTTPMethod.GET, url, params=params)

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
            for row in response_envelope.value:
                yield row.model_dump()

    def add_row(self, data: dict[str, Any]) -> dict[str, Any]:
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

        fields_payload = {"fields": data}

        response = self.client.make_request(
            "POST", add_row_url, json=fields_payload
        )

        return response

    def edit_row(self, row_id: str, data: dict[str, Any]) -> dict[str, Any]: ...

    def delete_row(self, row_id: str) -> bool: ...

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...

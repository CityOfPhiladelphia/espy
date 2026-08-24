# list.py
from collections.abc import Iterator
from enum import StrEnum
from typing import Any

from graph_api_functions.client import GraphAPIClient, ClientEndpoints
#TODO: Make host name a variable, not a constant. Edit in optional config file?
from graph_api_functions.constants import HOST_NAME, GRAPH_URL
from graph_api_functions.models import (
    GraphAPIResponse,
    HTTPMethod,
    SharePointListRow,
    SharePointListColumn,
)
from graph_api_functions.urls import build_url


class ListEndpoints(StrEnum):
    """
    An endpoint registry for all API operations made by the SharePointList
    class.
    """
    LIST_ID = ("{graph_url}/sites/{site_id}/lists/{list_name}")

    LIST_ROWS = ("{graph_url}/sites/"
                 "{site_id}/lists/{list_id}/items?")

    LIST_COLUMNS = ("{graph_url}/sites/"
                    "{site_id}/lists/{list_id}/columns?")

    GET_ROW = (
        "{graph_url}/sites/"
        "{site_id}/lists/{list_id}/items/{row_id}?$expand=fields"
    )

    ADD_ROW = (
        "{graph_url}/sites/{site_id}/lists/{list_id}/items"
    )


class SharePointList:
    def __init__(self, client: GraphAPIClient, site_id: str, list_id: str):
        self.client = client
        self.site_id = site_id
        self.list_id = list_id
        self._column_mapping = None

    @classmethod
    def get_list(cls, site_path: str, list_name: str):
        """
        Authenticates and fetches the ids needed to create the SharePointList
        object.

        Args:
            site_path: str, the site path of the SharePoint list.
            list_name: str, the name of the SharePoint list.
        """
        # TODO: New method to get site and list ID
        client = GraphAPIClient.authenticate()

        site_id_url = build_url(
              ClientEndpoints.SITE_ID, 
              graph_url=GRAPH_URL, 
              hostname=HOST_NAME, 
              site_path=site_path
              )

        site_id = client.make_request("GET", site_id_url)["id"]

        list_id_url = build_url(
              ListEndpoints.LIST_ID,
              graph_url=GRAPH_URL,
              site_id=site_id,
              list_name=list_name
        )

        list_id = client.make_request("GET", list_id_url)["id"]

        return cls(client, site_id, list_id)

    def get_row(self, row_id: str) -> dict[str, Any]:
        """
        Get a single row from a list.
        
        Args:
            row_id: The id of the row to return

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

    def list_columns(self) -> list:
        columns_url = build_url(
            ListEndpoints.LIST_COLUMNS,
            graph_url=GRAPH_URL,
            site_id=self.site_id,
            list_id=self.list_id,
        )

        raw_columns_data = self.client.make_request(HTTPMethod.GET, columns_url)

        response_envelope = GraphAPIResponse[SharePointListColumn]\
           .model_validate(raw_columns_data)

        return [column for column in response_envelope.model_dump()["value"]]

    def _get_column_mapping(self) -> dict[str, str]:
        """
        Gets a column mapping for the list mapping the column's user-visible
        name to the column's real name in the graph API.
        """

        column_mapping = {}

        columns = self.list_columns()

        for column in columns:
            column_mapping[column["display_name"]] = column["name"]

        self._column_mapping = column_mapping
        
        return column_mapping


    def _fetch_page(
        self, url: str | None, params
    ) -> GraphAPIResponse[SharePointListRow] | None:

        # For pagination reasons, return none if no URL is provided
        if not url:
            return None

        raw_data = self.client.make_request(HTTPMethod.GET, url, params=params)

        response_envelope = GraphAPIResponse[SharePointListRow].model_validate(
            raw_data
        )

        return response_envelope

    def list_rows(self) -> Iterator[dict]:
        next_link = build_url(
            ListEndpoints.LIST_ROWS,
            graph_url=GRAPH_URL,
            site_id=self.site_id, 
            list_id=self.list_id
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
        add_row_url = build_url(ListEndpoints.ADD_ROW, graph_url=GRAPH_URL)

        response = self.client.make_request(
            "POST", add_row_url)

        print(response)

    def edit_row(self, row_id: str, data: dict[str, Any]) -> dict[str, Any]: ...

    def delete_row(self, row_id: str) -> bool: ...

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...
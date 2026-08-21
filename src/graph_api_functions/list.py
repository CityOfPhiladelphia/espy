# list.py
from collections.abc import Iterator

from graph_api_functions.client import GraphAPIClient, _URLResolver
from graph_api_functions.models import (
    HTTPMethod,
    GraphAPIResponse,
    SharePointListRow,
)
from graph_api_functions.constants import HOST_NAME

from enum import StrEnum
from typing import Any

import string


class ListEndpoints(StrEnum):
    LIST_ROWS = "https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?"

    LIST_COLUMNS = "https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/columns?"

    GET_ROW = (
        "https://graph.microsoft.com/v1.0/sites/"
        "{site_id}/lists/{list_id}/items/{row_id}?$expand=fields"
    )


def build_url(url: ListEndpoints, **kwargs):
    """
    Given a GraphAPI endpoint template, builds the actual URL to request
    against. Returns an error if keys are missing.
    """

    # get all placeholder fields in the string
    fields = {
        field for _, field, _, _ in string.Formatter().parse(url) if field
    }

    # If the caller has forgotten a field, raise an error
    missing_fields = fields - kwargs.keys()

    if missing_fields:
        raise KeyError(
            f"The following fields are missing: {', '.join(missing_fields)}"
        )

    # Return url with correct keyword arguments passed in
    return url.format(**kwargs)


class SharePointList:
    def __init__(self, client: GraphAPIClient, site_id: str, list_id: str):
        self.client = client
        self.site_id = site_id
        self.list_id = list_id

    @classmethod
    def get_list(cls, site_path: str, list_name: str):
        client = GraphAPIClient.authenticate()
        site_id = _URLResolver.get_site_id(client, HOST_NAME, site_path)
        list_id = _URLResolver.get_list_id(client, site_id, list_name)

        return cls(client, site_id, list_id)

    def get_row(self, row_id: str) -> dict[str, Any]:
        row_url = build_url(
            ListEndpoints.GET_ROW,
            site_id=self.site_id,
            list_id=self.list_id,
            row_id=row_id,
        )

        raw_row_data = self.client.make_request(HTTPMethod.GET, row_url)

        return raw_row_data

    def list_columns(self) -> dict:
        columns_url = build_url(
            ListEndpoints.LIST_COLUMNS,
            site_id=self.site_id,
            list_id=self.list_id,
        )

        raw_columns_data = self.client.make_request(HTTPMethod.GET, columns_url)

        return raw_columns_data

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
            ListEndpoints.LIST_ROWS, site_id=self.site_id, list_id=self.list_id
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

    def add_row(self, data: dict[str, Any]) -> dict[str, Any]: ...

    def edit_row(self, row_id: str, data: dict[str, Any]) -> dict[str, Any]: ...

    def delete_row(self, row_id: str) -> bool: ...

    def upsert_row(
        self, key_col: str, data: dict[str, Any]
    ) -> dict[str, Any]: ...

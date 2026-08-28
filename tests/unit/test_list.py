from collections.abc import Iterator
from unittest.mock import MagicMock

import pytest
import httpx

from graph_api_functions.list import SharePointList
from graph_api_functions.models.models import GraphAPIResponse, SharePointListRow


def mocked_response(fixture_dict: dict) -> httpx.Response:
    mock_response = MagicMock()
    mock_response.json.return_value = fixture_dict
    mock_response.status_code = 200

    return mock_response

def mocked_get(fixture_dict):
    mock_get = MagicMock()
    mock_get.return_value = mocked_response(fixture_dict)

    return mock_get

@pytest.fixture
def configured_test_list(configured_client) -> SharePointList:
    test_list = SharePointList(configured_client, "site_id", "list_id")

    return test_list


@pytest.fixture
def mocked_page_data_with_next_link() -> dict:
    mock_page_data = {
        "odata_context": "123",
        "next_link": "example.com",
        "value": [
            {"fields": {"id": "1"}},
        ],
    }

    return mock_page_data


@pytest.fixture
def mocked_single_page_data() -> dict:
    mock_page_data = {
        "odata_context": "456",
        "value": [
            {"fields": {"id": "2"}},
        ],
    }

    return mock_page_data

@pytest.fixture
def mocked_row_data() -> dict:
    mock_row_data = {
        "odata_context": "789",
        "fields": {
            "Title": "Test",
            "name": "Billy Penn",
            "address": "1234 Market Street"
        }
    }

    return mock_row_data


def test_fetch_page_returns_graph_api_response(
    configured_test_list, mocked_single_page_data, monkeypatch
):
    mock_get = mocked_get(mocked_single_page_data)

    monkeypatch.setattr(
        "graph_api_functions.list.GraphAPIClient.make_request", mock_get
    )

    page_data = configured_test_list._fetch_page(
        "example.com", params={"test": True}
    )

    assert isinstance(page_data, GraphAPIResponse[SharePointListRow])


def test_list_page_yields_dict(
    configured_test_list,
    mocked_page_data_with_next_link,
    mocked_single_page_data,
    monkeypatch,
):
    mock_get = mocked_get(mocked_single_page_data)

    mock_get.side_effect = [
        mocked_response(mocked_page_data_with_next_link),
        mocked_response(mocked_single_page_data),
    ]

    monkeypatch.setattr(
        "graph_api_functions.list.GraphAPIClient.make_request", mock_get
    )

    result = configured_test_list.list_rows()

    ## Assert that what is returned is actually an iterator
    assert isinstance(result, Iterator)

    record = next(result)

    ## Assert that the iterator yields a dict
    assert isinstance(record, dict)

    record = next(result)

    ## Assert that iterator stops yielding when exhausted
    assert isinstance(record, dict)

    with pytest.raises(StopIteration):
        next(result)


def test_get_row_yields_dict(
    configured_test_list,
    mocked_row_data,
    monkeypatch,
):

    mock_get = mocked_get(mocked_row_data)

    monkeypatch.setattr(
        "graph_api_functions.list.GraphAPIClient.make_request", mock_get
    )

    result = configured_test_list.get_row("1")

    assert isinstance(result, dict)

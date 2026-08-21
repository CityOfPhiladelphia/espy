from collections.abc import Iterator
from unittest.mock import MagicMock

import pytest

from graph_api_functions.list import SharePointList
from graph_api_functions.models import GraphAPIResponse


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


def test_fetch_page_returns_graph_api_response(
    configured_test_list, mocked_single_page_data, monkeypatch
):
    mock_response = MagicMock()
    mock_response.return_value = mocked_single_page_data
    monkeypatch.setattr(
        "graph_api_functions.list.GraphAPIClient.make_request", mock_response
    )

    page_data = configured_test_list._fetch_page(
        "example.com", params={"test": True}
    )

    assert isinstance(page_data, GraphAPIResponse)


def test_list_page_yields_dict(
    configured_test_list,
    mocked_page_data_with_next_link,
    mocked_single_page_data,
    monkeypatch,
):

    mock_get = MagicMock()
    mock_get.side_effect = [
        mocked_page_data_with_next_link,
        mocked_single_page_data,
    ]

    monkeypatch.setattr(
        "graph_api_functions.list.GraphAPIClient.make_request", mock_get
    )

    result = configured_test_list.list_rows()

    ## Assert that what is returned is actually an iterator
    assert isinstance(result, Iterator)

    record = next(result)

    print(record)

    ## Assert that the iterator yields a dict
    assert isinstance(record, dict)

    record = next(result)

    ## Assert that iterator stops yielding when exhausted
    assert isinstance(record, dict)

    with pytest.raises(StopIteration):
        next(result)

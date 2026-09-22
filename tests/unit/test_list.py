from collections.abc import Iterator
from unittest.mock import MagicMock

import httpx
import pytest

from espy.list import SharePointList
from espy.models.models import (
    COLUMN_KIND_PYTHON_TYPES,
    ColumnKind,
    GraphAPIResponse,
    InvalidIncomingRowError,
    SharePointListRow,
    SharePointListColumn
)


def mocked_response(fixture_dict: dict) -> httpx.Response:
    mock_response = MagicMock()
    mock_response.json.return_value = fixture_dict
    mock_response.status_code = 200

    return mock_response

def mocked_request(fixture_dict):
    mock_request = MagicMock()
    mock_request.return_value = mocked_response(fixture_dict)

    return mock_request

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
def mocked_batch_data() -> dict:
    mock_batch_data = {
        "responses": [{
            "id": "0",
            "status": 200,
            "headers": {"header": "test"},
            "body": {"value": [
                {"fields":
                 {
                    "Title": "Test",
                    "col_one": "123456",
                    "col_two": "1234 Market Street"
                 }}
            ]}
    }]
    }

    return mock_batch_data

@pytest.fixture
def mocked_empty_batch_data() -> dict:
    mock_empty_batch_data = {
        "responses": [{
            "id": "0",
            "status": 200,
            "headers": {"header": "test"},
            "body": {"value": []}
        }]
    }

    return mock_empty_batch_data

@pytest.fixture
def raw_column_data() -> list[dict]:
    return [
        {'description': '',
         'display_name': 'col_one',
         'enforce_unique_values': True,
         'hidden': False,
         'id': '1234',
         'indexed': True,
         'name': 'col_one',
         'read_only': False,
         'required': True,
         'type': 'text'},
        {'description': '',
         'display_name': 'col_two',
         'enforce_unique_values': False,
         'hidden': False,
         'id': '5678',
         'indexed': False,
         'name': 'col_two',
         'read_only': False,
         'required': True,
         'type': 'boolean'},
        ]

@pytest.fixture
def mocked_column_api_response(raw_column_data) -> dict:
    return {"value": raw_column_data}

@pytest.fixture
def mocked_column_data(raw_column_data) -> list[SharePointListColumn]:
    return [SharePointListColumn.model_validate(column)
            for column in raw_column_data]

def test_fetch_page_returns_graph_api_response(
    configured_test_list, mocked_single_page_data, monkeypatch
):
    mock_get = mocked_request(mocked_single_page_data)

    monkeypatch.setattr(
        "espy.list.GraphAPIClient.make_request", mock_get
    )

    page_data = configured_test_list._fetch_page(
        "example.com", params={"test": True}
    )

    assert isinstance(page_data, GraphAPIResponse)


def test_list_page_yields_dict(
    configured_test_list,
    mocked_column_data,
    mocked_page_data_with_next_link,
    mocked_single_page_data,
    monkeypatch,
):  

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    mock_get = mocked_request(mocked_single_page_data)

    mock_get.side_effect = [
        mocked_response(mocked_page_data_with_next_link),
        mocked_response(mocked_single_page_data),
    ]

    monkeypatch.setattr(
        "espy.list.GraphAPIClient.make_request", mock_get
    )

    result = configured_test_list.list_rows()

    assert isinstance(result, Iterator)

    record = next(result)

    assert isinstance(record, dict)

    record = next(result)

    assert isinstance(record, dict)

    with pytest.raises(StopIteration):
        next(result)

def test_get_rows_yields_share_point_list_rows(
    configured_test_list,
    mocked_column_data,
    mocked_batch_data,
    monkeypatch,
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    mock_get = mocked_request(mocked_batch_data)

    monkeypatch.setattr(
        "espy.list.GraphAPIClient.make_request", mock_get
    )

    result = configured_test_list.get_rows("col_one", ["123456"])
    next_resp = next(result)

    assert isinstance(next_resp, SharePointListRow)

def test_get_rows_raises_key_error_when_column_not_present(
    configured_test_list,
    mocked_column_data,
    mocked_batch_data,
    monkeypatch,
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    mock_get = mocked_request(mocked_batch_data)

    monkeypatch.setattr(
        "espy.list.GraphAPIClient.make_request", mock_get
    )

    # get_rows is a generator, so the validation only runs once consumed
    with pytest.raises(KeyError):
        next(configured_test_list.get_rows("col_three", ["123"]))

def test_get_rows_raises_value_error_when_value_not_present(
    configured_test_list,
    mocked_column_data,
    mocked_empty_batch_data,
    monkeypatch,
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    mock_get = mocked_request(mocked_empty_batch_data)

    monkeypatch.setattr(
        "espy.list.GraphAPIClient.make_request", mock_get
    )

    with pytest.raises(ValueError):
        next(configured_test_list.get_rows("col_one", ["456"]))

def test_list_columns_returns_list_of_share_point_list_column(
        configured_test_list,
        mocked_column_api_response,
        monkeypatch
):
    mock_get = mocked_request(mocked_column_api_response)

    monkeypatch.setattr(
        "espy.list.GraphAPIClient.make_request", mock_get
    )

    result = configured_test_list.list_columns()

    assert isinstance(result, list)
    assert isinstance(result[0], SharePointListColumn)

def test_add_rows_breaks_with_bad_column_name(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):  
    data = [{"col_three": "1234 Market St"}]

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    # add_rows is a generator, so it must be consumed to run validation
    with pytest.raises(InvalidIncomingRowError):
        list(configured_test_list.add_rows(data))

def test_column_kind_enum_matches_type_mapping():
    assert set(ColumnKind) == set(COLUMN_KIND_PYTHON_TYPES.keys())

def test_add_rows_breaks_with_bad_data_type(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):
    data = [{"col_two": "True"}]

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns",
        mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        list(configured_test_list.add_rows(data))

def test_edit_row_breaks_with_bad_column_name(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):  
    data = {"col_three": "1234 Market St"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data["value"]

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        configured_test_list.edit_row("col_four", "123456", data)

def test_edit_row_breaks_with_bad_data_type(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):
    data = {"col_two": "fdhfdfh"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data["value"]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", 
        mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        configured_test_list.edit_row("col_one", "123456", data)

def test_edit_row_breaks_with_non_index_col(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):
    data = {"col_two": "fdhfdfh"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data["value"]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", 
        mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        configured_test_list.edit_row("col_two", "123456", data)

def test_delete_row_breaks_with_bad_column_name(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):  

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data["value"]

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        configured_test_list.delete_row("col_four", "123456")

def test_delete_row_breaks_with_non_index_col(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data["value"]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", 
        mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        configured_test_list.delete_row("col_two", "123456")

def test_upsert_row_returns_error_when_key_col_not_present(
        configured_test_list,
        mocked_column_data,
        monkeypatch
):
    data = {"col_three": "1234 Market St"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = mocked_column_data["value"]

    monkeypatch.setattr("espy.list.SharePointList.list_columns", mock_list_columns_func)

    with pytest.raises(InvalidIncomingRowError):
        configured_test_list.upsert_row("col_four", data)
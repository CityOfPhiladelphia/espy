from collections.abc import Iterator
from unittest.mock import MagicMock, patch

import httpx
import pytest

from espy.list import SharePointList
from espy.models.models import (
    ACCEPTABLE_PYTHON_TYPES,
    READ_ONLY_COLUMN_KINDS,
    ColumnKind,
    GraphCollection,
    SharePointListColumn,
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


def test_setup_authenticates_and_resolves_ids(fake_creds, monkeypatch):
    monkeypatch.setenv("SHAREPOINT_HOSTNAME", "env.sharepoint.com")
    mock_instance = MagicMock()
    mock_instance.get_site_id.return_value = "test_site"
    mock_instance.make_request.return_value = mocked_response({"id": "list_id"})

    with patch("espy.list.GraphAPIClient") as MockedClient:
        MockedClient.authenticate.return_value = mock_instance

        sp_list = SharePointList.setup(
            site_name="TeamSite", list_name="MyList", creds=fake_creds
        )

    MockedClient.authenticate.assert_called_once_with(fake_creds)
    assert sp_list.client is mock_instance
    assert sp_list.site_id == "test_site"
    assert sp_list.list_id == "list_id"
    mock_instance.get_site_id.assert_called_once_with(
        "env.sharepoint.com", "TeamSite"
    )


def test_setup_uses_custom_hostname(fake_creds):
    mock_instance = MagicMock()
    mock_instance.make_request.return_value = mocked_response({"id": "list_id"})

    with patch("espy.list.GraphAPIClient") as MockedClient:
        MockedClient.authenticate.return_value = mock_instance

        SharePointList.setup(
            site_name="TeamSite",
            list_name="MyList",
            hostname="example.sharepoint.com",
            creds=fake_creds,
        )

    mock_instance.get_site_id.assert_called_once_with(
        "example.sharepoint.com", "TeamSite"
    )


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
            "col_one": "123456",
            "col_two": "1234 Market Street",
        },
        "value": [
            {
                "fields": {
                    "Title": "Test",
                    "col_one": "123456",
                    "col_two": "1234 Market Street",
                    "id": "123",
                }
            }
        ],
    }

    return mock_row_data


@pytest.fixture
def mocked_batch_data() -> dict:
    mock_batch_data = {
        "responses": [
            {
                "id": "0",
                "status": 200,
                "headers": {"header": "test"},
                "body": {
                    "value": [
                        {
                            "fields": {
                                "Title": "Test",
                                "col_one": "123456",
                                "col_two": "1234 Market Street",
                            }
                        }
                    ]
                },
            }
        ]
    }

    return mock_batch_data


@pytest.fixture
def mocked_column_data() -> dict:

    mock_column_data = {
        "odata_context": "101112",
        "value": [
            {
                "description": "",
                "display_name": "col_one",
                "enforce_unique_values": True,
                "hidden": False,
                "id": "1234",
                "indexed": True,
                "name": "col_one",
                "read_only": False,
                "required": True,
                "text": {},
            },
            {
                "description": "",
                "display_name": "col_two",
                "enforce_unique_values": False,
                "hidden": False,
                "id": "5678",
                "indexed": False,
                "name": "col_two",
                "read_only": False,
                "required": True,
                "boolean": {},
            },
        ],
    }

    return mock_column_data


def test_fetch_page_returns_graph_api_response(
    configured_test_list, mocked_single_page_data, monkeypatch
):
    mock_get = mocked_request(mocked_single_page_data)

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_get)

    page_data = configured_test_list._fetch_page(
        "example.com", params={"test": True}
    )

    assert isinstance(page_data, GraphCollection)


def test_list_page_yields_dict(
    configured_test_list,
    mocked_column_data,
    mocked_page_data_with_next_link,
    mocked_single_page_data,
    monkeypatch,
):

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    mock_get = mocked_request(mocked_single_page_data)

    mock_get.side_effect = [
        mocked_response(mocked_page_data_with_next_link),
        mocked_response(mocked_single_page_data),
    ]

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_get)

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
    mocked_column_data,
    mocked_row_data,
    monkeypatch,
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    mock_get = mocked_request(mocked_row_data)

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_get)

    result = configured_test_list.get_row("col_one", "123")
    assert isinstance(result, dict)


def test_get_row_raises_key_error_when_column_not_present(
    configured_test_list,
    mocked_column_data,
    mocked_row_data,
    monkeypatch,
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    mock_get = mocked_request(mocked_row_data)

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_get)

    with pytest.raises(KeyError):
        configured_test_list.get_row("col_three", "123")


def test_get_row_raises_value_error_when_value_not_present(
    configured_test_list,
    mocked_column_data,
    monkeypatch,
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    mock_get = mocked_request({"value": []})

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_get)

    with pytest.raises(ValueError):
        configured_test_list.get_row("col_one", "456")


def test_list_column_returns_list_of_columns(
    configured_test_list, mocked_column_data, monkeypatch
):
    mock_get = mocked_request(mocked_column_data)

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_get)

    result = configured_test_list.list_columns()

    assert isinstance(result, list)
    assert isinstance(result[0], SharePointListColumn)
    assert result[0].type == ColumnKind.TEXT
    assert result[1].type == ColumnKind.BOOLEAN



def test_add_row_breaks_with_bad_column_name(
    configured_test_list, mocked_column_data, monkeypatch
):
    data = {"col_three": "1234 Market St"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.add_row(data)


def test_column_kind_enum_matches_type_mapping():
    writable_kinds = set(ACCEPTABLE_PYTHON_TYPES.keys())

    assert writable_kinds.isdisjoint(READ_ONLY_COLUMN_KINDS)
    assert set(ColumnKind) == writable_kinds | READ_ONLY_COLUMN_KINDS



def test_add_row_breaks_with_bad_data_type(
    configured_test_list, mocked_column_data, monkeypatch
):
    data = {"col_two": "True"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.add_row(data)


def test_edit_row_breaks_with_bad_column_name(
    configured_test_list, mocked_column_data, monkeypatch
):
    # col three does not exist
    data = {"col_three": "1234 Market St"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.edit_row("col_four", "123456", data)


def test_edit_row_breaks_with_bad_data_type(
    configured_test_list, mocked_column_data, monkeypatch
):
    # col_two is a boolean column, so a string value is the wrong type
    data = {"col_two": "fdhfdfh"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    # get_row returns a formatted row, which includes the row id
    mock_list_get_row_func = MagicMock()
    mock_list_get_row_func.return_value = {
        "col_one": "123456",
        "col_two": True,
        "id": "123",
    }

    monkeypatch.setattr(
        "espy.list.SharePointList.get_row", mock_list_get_row_func
    )

    mock_request = MagicMock()

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", mock_request)

    with pytest.raises(KeyError, match="incorrect data type"):
        configured_test_list.edit_row("col_one", "123456", data)

    # Validation should fail before the PATCH request is sent
    mock_request.assert_not_called()


def test_edit_row_breaks_with_non_index_col(
    configured_test_list, mocked_column_data, monkeypatch
):
    data = {"col_two": "fdhfdfh"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.edit_row("col_two", "123456", data)


def test_delete_row_breaks_with_bad_column_name(
    configured_test_list, mocked_column_data, monkeypatch
):

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.delete_row("col_four", "123456")


def test_delete_row_breaks_with_non_index_col(
    configured_test_list, mocked_column_data, monkeypatch
):
    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.delete_row("col_two", "123456")


@pytest.mark.skip(reason="Not implemented yet")
def test_upsert_row_returns_error_when_key_col_not_present(
    configured_test_list, mocked_column_data, monkeypatch
):
    data = {"col_three": "1234 Market St"}

    mock_list_columns_func = MagicMock()
    mock_list_columns_func.return_value = [
        SharePointListColumn.model_validate(column)
        for column in mocked_column_data["value"]
    ]

    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns", mock_list_columns_func
    )

    with pytest.raises(KeyError):
        configured_test_list.upsert_row("col_four", data)

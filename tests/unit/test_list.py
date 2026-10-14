from collections.abc import Iterator
from unittest.mock import MagicMock, patch

import httpx
import pytest

from espy.constants import MAX_BATCH_SIZE
from espy.list import SharePointList
from espy.models.models import (
    ACCEPTABLE_PYTHON_TYPES,
    READ_ONLY_COLUMN_KINDS,
    ColumnKind,
    GraphCollection,
    SharePointListColumn,
)
from espy.operations import AddRow, DeleteRow, EditRow, GetRow


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


# ---------------------------------------------------------------------------
# Batch operations
# ---------------------------------------------------------------------------


@pytest.fixture
def batch_list(configured_test_list, mocked_column_data, monkeypatch):
    """A SharePointList with columns patched in (col_one is indexed)."""
    monkeypatch.setattr(
        "espy.list.SharePointList.list_columns",
        MagicMock(
            return_value=[
                SharePointListColumn.model_validate(column)
                for column in mocked_column_data["value"]
            ]
        ),
    )

    return configured_test_list


class FakeGraph:
    """Stands in for GraphAPIClient.make_request for $batch calls.

    ``respond`` receives each sub-request dict and returns a
    ``(status, body)`` tuple. Every call's list of sub-requests is kept
    in ``self.batches`` so tests can inspect what was sent.
    """

    def __init__(self, respond, shuffle=False, drop_last=False):
        self.respond = respond
        self.shuffle = shuffle
        self.drop_last = drop_last
        self.batches: list[list[dict]] = []

    def __call__(self, method, url, **kwargs):
        requests = kwargs["json"]["requests"]
        self.batches.append(requests)

        responses = []
        for req in requests:
            status, body = self.respond(req)
            responses.append(
                {
                    "id": req["id"],
                    "status": status,
                    "headers": {},
                    "body": body,
                }
            )

        if self.drop_last:
            responses = responses[:-1]

        if self.shuffle:
            responses.reverse()

        return mocked_response({"responses": responses})


def row_body(row_id, col_one="x", col_two=True):
    return {
        "fields": {"col_one": col_one, "col_two": col_two, "id": str(row_id)}
    }


def lookup_body(row_id, col_one="x"):
    return {"value": [row_body(row_id, col_one)]}


def install(monkeypatch, fake):
    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", fake)


def test_batch_returns_one_result_per_operation_in_order(
    batch_list, monkeypatch
):
    fake = FakeGraph(lambda req: (201, row_body(req["id"])), shuffle=True)
    install(monkeypatch, fake)

    ops = [AddRow({"col_one": str(i)}) for i in range(3)]
    results = batch_list.batch(ops)

    assert [r.operation for r in results] == ops
    assert all(r.ok for r in results)
    # Responses came back reversed; ids must still line up with inputs
    assert [r.value["id"] for r in results] == ["0", "1", "2"]


def test_batch_splits_into_groups_of_max_batch_size(batch_list, monkeypatch):
    fake = FakeGraph(lambda req: (201, row_body(req["id"])))
    install(monkeypatch, fake)

    total = MAX_BATCH_SIZE * 2 + 5
    results = batch_list.batch([AddRow({"col_one": "a"})] * total)

    assert [len(b) for b in fake.batches] == [
        MAX_BATCH_SIZE,
        MAX_BATCH_SIZE,
        5,
    ]
    # Request ids are unique across groups
    sent_ids = [req["id"] for batch in fake.batches for req in batch]
    assert sent_ids == [str(i) for i in range(total)]
    assert len(results) == total
    assert [r.value["id"] for r in results] == sent_ids


def test_batch_empty_makes_no_request(batch_list, monkeypatch):
    fake = FakeGraph(lambda req: (200, {}))
    install(monkeypatch, fake)

    assert batch_list.batch([]) == []
    assert fake.batches == []


def test_batch_reports_partial_failure_without_raising(batch_list, monkeypatch):
    def respond(req):
        if req["id"] == "1":
            return 400, {"error": {"code": "invalidRequest", "message": "bad"}}
        return 201, row_body(req["id"])

    install(monkeypatch, FakeGraph(respond))

    results = batch_list.batch([AddRow({"col_one": str(i)}) for i in range(3)])

    assert [r.ok for r in results] == [True, False, True]
    assert results[1].error.status == 400
    assert results[1].error.code == "invalidRequest"
    assert results[1].error.message == "bad"
    assert results[1].value is None


def test_batch_missing_sub_response_becomes_error(batch_list, monkeypatch):
    fake = FakeGraph(lambda req: (201, row_body(req["id"])), drop_last=True)
    install(monkeypatch, fake)

    results = batch_list.batch([AddRow({"col_one": "a"}) for _ in range(2)])

    assert results[0].ok
    assert not results[1].ok
    assert results[1].error.status is None


def test_batch_unparseable_body_becomes_error(batch_list, monkeypatch):
    # A GET that returns no matching rows -> "Row not found."
    install(monkeypatch, FakeGraph(lambda req: (200, {"value": []})))

    results = batch_list.batch([GetRow("col_one", "nope")])

    assert not results[0].ok
    assert results[0].error.status == 200
    assert "not found" in results[0].error.message.lower()


def test_batch_builds_expected_request_shapes(batch_list, monkeypatch):
    fake = FakeGraph(lambda req: (200, row_body(1)))
    install(monkeypatch, fake)

    batch_list.batch(
        [
            GetRow("col_one", "abc"),
            AddRow({"col_one": "a"}),
            EditRow(5, {"col_one": "b"}),
            DeleteRow(6),
        ]
    )

    get, add, edit, delete = fake.batches[0]
    assert get["method"] == "GET" and get["body"] is None
    assert "col_one eq abc" in get["url"]
    assert add["method"] == "POST"
    assert add["body"] == {"fields": {"col_one": "a"}}
    assert edit["method"] == "PATCH" and edit["url"].endswith("/items/5")
    assert edit["body"] == {"fields": {"col_one": "b"}}
    assert delete["method"] == "DELETE" and delete["url"].endswith("/items/6")


def test_batch_unknown_operation_raises(batch_list):
    with pytest.raises(ValueError, match="Unknown request type"):
        batch_list._dispatch_batch_operation(object())


def test_add_rows_invalid_data_raises_before_any_request(
    batch_list, monkeypatch
):
    fake = FakeGraph(lambda req: (201, row_body(1)))
    install(monkeypatch, fake)

    with pytest.raises(KeyError):
        batch_list.add_rows([{"col_one": "ok"}, {"not_a_column": "x"}])

    assert fake.batches == []


def test_add_rows_returns_formatted_rows(batch_list, monkeypatch):
    install(monkeypatch, FakeGraph(lambda req: (201, row_body(9, "new"))))

    results = batch_list.add_rows([{"col_one": "new"}])

    assert results[0].ok
    assert results[0].value == {"col_one": "new", "col_two": True, "id": "9"}


def test_get_rows_mixes_found_and_missing(batch_list, monkeypatch):
    def respond(req):
        if "eq 1" in req["url"]:
            return 200, lookup_body(1, "1")
        return 200, {"value": []}

    install(monkeypatch, FakeGraph(respond))

    results = batch_list.get_rows("col_one", ["1", "2"])

    assert results[0].ok and results[0].value["col_one"] == "1"
    assert not results[1].ok


def test_get_rows_rejects_non_indexed_column(batch_list, monkeypatch):
    fake = FakeGraph(lambda req: (200, {}))
    install(monkeypatch, fake)

    with pytest.raises(KeyError):
        batch_list.get_rows("col_two", ["x"])

    assert fake.batches == []


def test_edit_rows_looks_up_then_edits_found_rows(batch_list, monkeypatch):
    def respond(req):
        if req["method"] == "GET":
            if "eq missing" in req["url"]:
                return 200, {"value": []}
            return 200, lookup_body(7, "found")
        return 200, row_body(7, "edited")

    fake = FakeGraph(respond)
    install(monkeypatch, fake)

    results = batch_list.edit_rows(
        "col_one",
        ["found", "missing"],
        [{"col_one": "edited"}, {"col_one": "never"}],
    )

    # Two batches: lookup of both, then edit of the one that was found
    assert len(fake.batches) == 2
    assert [r["method"] for r in fake.batches[1]] == ["PATCH"]
    assert fake.batches[1][0]["url"].endswith("/items/7")

    assert results[0].ok and results[0].value["col_one"] == "edited"
    assert isinstance(results[0].operation, EditRow)
    assert not results[1].ok
    assert isinstance(results[1].operation, GetRow)


def test_edit_rows_raises_on_length_mismatch(batch_list, monkeypatch):
    fake = FakeGraph(lambda req: (200, {}))
    install(monkeypatch, fake)

    with pytest.raises(ValueError, match="same length"):
        batch_list.edit_rows("col_one", ["a", "b"], [{"col_one": "x"}])

    assert fake.batches == []


def test_edit_rows_reports_edit_failure(batch_list, monkeypatch):
    def respond(req):
        if req["method"] == "GET":
            return 200, lookup_body(7)
        return 409, {"error": {"code": "conflict", "message": "locked"}}

    install(monkeypatch, FakeGraph(respond))

    results = batch_list.edit_rows("col_one", ["a"], [{"col_one": "b"}])

    assert not results[0].ok
    assert results[0].error.code == "conflict"


def test_delete_rows_looks_up_then_deletes_found_rows(batch_list, monkeypatch):
    def respond(req):
        if req["method"] == "GET":
            if "eq missing" in req["url"]:
                return 200, {"value": []}
            return 200, lookup_body(3)
        return 204, None

    fake = FakeGraph(respond)
    install(monkeypatch, fake)

    results = batch_list.delete_rows("col_one", ["found", "missing"])

    assert len(fake.batches) == 2
    assert [r["method"] for r in fake.batches[1]] == ["DELETE"]
    assert fake.batches[1][0]["url"].endswith("/items/3")

    assert results[0].ok
    assert isinstance(results[0].operation, DeleteRow)
    assert not results[1].ok


def test_delete_rows_with_no_matches_skips_second_batch(
    batch_list, monkeypatch
):
    fake = FakeGraph(lambda req: (200, {"value": []}))
    install(monkeypatch, fake)

    results = batch_list.delete_rows("col_one", ["a", "b"])

    assert len(fake.batches) == 1
    assert not any(r.ok for r in results)


def test_delete_row_returns_none_on_empty_204(batch_list, monkeypatch):
    calls = []

    def fake_request(_self, method, url, **kwargs):
        calls.append(method)
        response = MagicMock()
        response.status_code = 204
        # An empty 204 body makes the real Response.json() raise
        response.json.side_effect = ValueError("no content")
        if method == "GET":
            response.status_code = 200
            response.json.side_effect = None
            response.json.return_value = lookup_body(3)
        return response

    monkeypatch.setattr("espy.list.GraphAPIClient.make_request", fake_request)

    assert batch_list.delete_row("col_one", "x") is None
    assert calls == ["GET", "DELETE"]

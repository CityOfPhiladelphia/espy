import pytest

from graph_api_functions.client import _URLResolver
from unittest.mock import MagicMock


def test_get_site_id_returns_str(configured_client, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"id": "123"}

    monkeypatch.setattr(
        "graph_api_functions.client.GraphAPIClient.make_request", mock_response
    )

    assert isinstance(
        _URLResolver.get_site_id(
            configured_client, "test_host.com", "site_path/"
        ),
        str,
    )


def test_get_drive_id_returns_str(configured_client, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {
        "value": [
            {"name": "wrong_library", "id": "1"},
            {"name": "right_library", "id": "2"},
        ]
    }

    monkeypatch.setattr(
        "graph_api_functions.client.GraphAPIClient.make_request", mock_response
    )

    assert isinstance(
        _URLResolver.get_drive_id(configured_client, "123", "right_library"),
        str,
    )


def test_get_drive_id_fetches_correct_library(configured_client, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {
        "value": [
            {"name": "wrong_library", "id": "1"},
            {"name": "right_library", "id": "2"},
        ]
    }

    monkeypatch.setattr(
        "graph_api_functions.client.GraphAPIClient.make_request", mock_response
    )

    assert (
        _URLResolver.get_drive_id(configured_client, "123", "right_library")
        == "2"
    )


def test_get_workbook_id_returns_str(configured_client, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"id": "123"}

    monkeypatch.setattr(
        "graph_api_functions.client.GraphAPIClient.make_request", mock_response
    )

    assert isinstance(
        _URLResolver.get_workbook_id(
            configured_client, "drive_id", "workbook/path"
        ),
        str,
    )


def test_get_list_id_returns_str(configured_client, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"id": "123"}

    monkeypatch.setattr(
        "graph_api_functions.client.GraphAPIClient.make_request", mock_response
    )

    assert isinstance(
        _URLResolver.get_list_id(configured_client, "site_id", "list_name"), str
    )

# tests/unit/test_client.py
from enum import StrEnum
from unittest.mock import MagicMock

import pytest

from EsPy.models.models import HTTPMethod, UnsupportedMethodError


def test_get_headers_constructs_valid_header_dict(configured_client):
    headers = configured_client._get_headers()

    assert headers == {
        "Authorization": "Bearer mock-token",
        "Content-Type": "application/json",
    }


def test_execute_request_invokes_correct_function(
    configured_client, monkeypatch
):
    mock_utility = MagicMock(return_value="Mocked Output")

    monkeypatch.setattr("EsPy.client.httpx.get", mock_utility)

    response = configured_client._execute_request(HTTPMethod.GET, "Hello World")

    mock_utility.assert_called_once()

    assert response == "Mocked Output"

def test_execute_request_returns_error_with_wrong_method(configured_client):

    class MockHTTPMethod(StrEnum):
        MOCK_METHOD = "mock_method"

    with pytest.raises(UnsupportedMethodError):
        configured_client._execute_request(
            MockHTTPMethod.MOCK_METHOD, "TEST_ENDPOINT"
        )

# tests/unit/test_client.py
from enum import StrEnum
from unittest.mock import MagicMock

import pytest

from espy.client import GraphAPIClient, resolve_hostname
from espy.models.models import HTTPMethod, UnsupportedMethodError


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

    monkeypatch.setattr("espy.client.httpx.get", mock_utility)

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


def test_authenticate_builds_credential_from_creds(fake_creds, monkeypatch):
    mock_cred_cls = MagicMock()
    monkeypatch.setattr("espy.client.ClientSecretCredential", mock_cred_cls)

    client = GraphAPIClient.authenticate(fake_creds)

    mock_cred_cls.assert_called_once_with(
        tenant_id="fake-tenant",
        client_id="fake-client",
        client_secret="fake-secret",
    )
    assert client.credential is mock_cred_cls.return_value


def test_authenticate_reads_creds_from_env(monkeypatch):
    monkeypatch.setenv("AZURE_TENANT_ID", "env-tenant")
    monkeypatch.setenv("AZURE_CLIENT_ID", "env-client")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "env-secret")
    mock_cred_cls = MagicMock()
    monkeypatch.setattr("espy.client.ClientSecretCredential", mock_cred_cls)

    GraphAPIClient.authenticate()

    mock_cred_cls.assert_called_once_with(
        tenant_id="env-tenant",
        client_id="env-client",
        client_secret="env-secret",
    )


def test_authenticate_names_missing_env_vars(monkeypatch):
    monkeypatch.setenv("AZURE_TENANT_ID", "env-tenant")
    monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)

    with pytest.raises(OSError) as exc_info:
        GraphAPIClient.authenticate()

    message = str(exc_info.value)
    assert "AZURE_CLIENT_ID" in message
    assert "AZURE_CLIENT_SECRET" in message
    assert "AZURE_TENANT_ID" not in message


def test_resolve_hostname_prefers_explicit_value(monkeypatch):
    monkeypatch.setenv("SHAREPOINT_HOSTNAME", "env.sharepoint.com")

    assert resolve_hostname("arg.sharepoint.com") == "arg.sharepoint.com"


def test_resolve_hostname_falls_back_to_env(monkeypatch):
    monkeypatch.setenv("SHAREPOINT_HOSTNAME", "env.sharepoint.com")

    assert resolve_hostname(None) == "env.sharepoint.com"


def test_resolve_hostname_raises_when_unset(monkeypatch):
    monkeypatch.delenv("SHAREPOINT_HOSTNAME", raising=False)

    with pytest.raises(OSError, match="SHAREPOINT_HOSTNAME"):
        resolve_hostname(None)

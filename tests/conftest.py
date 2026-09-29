from unittest.mock import MagicMock

import pytest

from espy.client import GraphAPIClient
from espy.models.models import APICredentials


@pytest.fixture
def configured_client(monkeypatch) -> GraphAPIClient:
    """Provides a mock GraphAPIClient with pre-configured request return values."""

    mock_cred_cls = MagicMock()

    fake_token = MagicMock()
    fake_token.token = "mock-token"

    mock_cred_cls.get_token.return_value = fake_token

    return GraphAPIClient(mock_cred_cls)


@pytest.fixture
def fake_creds() -> APICredentials:
    """Placeholder credentials for tests that exercise setup/authenticate."""
    return {
        "tenant_id": "fake-tenant",
        "client_id": "fake-client",
        "client_secret": "fake-secret",
    }

from unittest.mock import MagicMock

import pytest

from EsPy.client import GraphAPIClient


@pytest.fixture
def configured_client(monkeypatch) -> GraphAPIClient:
    """Provides a mock GraphAPIClient with pre-configured request return values."""

    mock_cred_cls = MagicMock()

    fake_token = MagicMock()
    fake_token.token = "mock-token"

    mock_cred_cls.get_token.return_value = fake_token

    return GraphAPIClient(mock_cred_cls)

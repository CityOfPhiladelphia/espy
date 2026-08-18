from unittest.mock import MagicMock
import pytest
from graph_api_functions.client import GraphAPIClient

@pytest.fixture
def configured_client(monkeypatch) -> GraphAPIClient:
    """Provides a mock GraphAPIClient with pre-configured request return values."""

    mock_cred_cls = MagicMock()
    mock_cred_instance = mock_cred_cls.return_value
    
    fake_token = MagicMock()
    fake_token.token = "mock-token"

    mock_cred_instance.get_token.return_value = fake_token

    # Reassigns the variable in src.client for the duration of the test
    monkeypatch.setattr("graph_api_functions.client.ClientSecretCredential", mock_cred_cls)
    return GraphAPIClient("tenant", "id", "secret")
import pytest
from graph_api_functions.list import SharePointList
from unittest.mock import MagicMock

@pytest.fixture
def configure_test_list(configured_client):
    test_list = SharePointList(configured_client, "site_id", "list_id")

    return test_list

def test_list_rows_returns_list_of_dicts(test_list):

    mock_response = MagicMock()
    mock_response.json.return_value = {"status": "success", "code": 200}
import pytest

from graph_api_functions.client import _URLResolver
from unittest.mock import MagicMock


@pytest.fixture
def resolver(configured_client):
    resolver = _URLResolver(configured_client)

    return resolver  

def test_get_site_id_returns_str(resolver, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"id": "123"}
    
    monkeypatch.setattr(
        "graph_api_functions.urls.GraphAPIClient.make_request",
        mock_response)
    
    assert isinstance(resolver.get_site_id("test_host.com", "site_path/" ), str)

def test_get_drive_id_returns_str(resolver, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"value": [
        {"name": "wrong_library", "id": "1"}, 
        {"name": "right_library", "id": "2"}]
        }
    
    monkeypatch.setattr(
        "graph_api_functions.urls.GraphAPIClient.make_request",
        mock_response)
    
    assert isinstance(resolver.get_drive_id("123", "right_library" ), str)

def test_get_drive_id_fetches_correct_library(resolver, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"value": [
        {"name": "wrong_library", "id": "1"}, 
        {"name": "right_library", "id": "2"}]
        }
    
    monkeypatch.setattr(
        "graph_api_functions.urls.GraphAPIClient.make_request",
        mock_response)
    
    assert resolver.get_drive_id("123", "right_library") == "2"

def test_get_workbook_id_returns_str(resolver, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"id": "123"}
    
    monkeypatch.setattr(
        "graph_api_functions.urls.GraphAPIClient.make_request",
        mock_response)
    
    assert isinstance(
        resolver.get_workbook_id("drive_id", "workbook/path"), str
        )

def test_get_list_id_returns_str(resolver, monkeypatch):
    mock_response = MagicMock()
    mock_response.return_value = {"id": "123"}
    
    monkeypatch.setattr(
        "graph_api_functions.urls.GraphAPIClient.make_request",
        mock_response)
    
    assert isinstance(
        resolver.get_list_id("site_id", "list_name"), str
        )
from graph_api_functions.urls import UrlConstructor
from graph_api_functions.models import SharePointURL


def test_site_url_returns_url():
    assert isinstance(UrlConstructor.site_url('test.com', 'test_path'), SharePointURL)

def test_drive_url_returns_url():
    assert isinstance(UrlConstructor.drive_url('drive_id'), SharePointURL)

def test_workbook_url_returns_url():
    assert isinstance(UrlConstructor.workbook_url(
        'drive_id', 'path/to/workbook'), SharePointURL)

def test_list_url_returns_url():
    assert isinstance(UrlConstructor.list_url(
        'site_id', 'list_name'), SharePointURL)
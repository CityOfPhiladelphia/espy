from graph_api_functions.client import _UrlConstructor


def test_site_url_returns_str():
    assert isinstance(_UrlConstructor.site_url('test.com', 'test_path'), str)

def test_drive_url_returns_str():
    assert isinstance(_UrlConstructor.drive_url('drive_id'), str)

def test_workbook_url_returns_str():
    assert isinstance(_UrlConstructor.workbook_url(
        'drive_id', 'path/to/workbook'), str)

def test_list_url_returns_str():
    assert isinstance(_UrlConstructor.list_url(
        'site_id', 'list_name'), str)
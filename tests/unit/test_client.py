# tests/unit/test_client.py
def test_get_headers_constructs_valid_header_dict(configured_client):
    headers = configured_client._get_headers()

    assert headers == {
        "Authorization": "Bearer mock-token",
        "Content-Type": "application/json"
    }

def test_execute_request_returns_valid_dict(configured_client):
    response = configured_client.execute_request

    assert isinstance(response, dict)
# client.py
from azure.identity import ClientSecretCredential
from graph_api_functions.constants import SCOPE
from graph_api_functions.models import HTTPMethod, UnsupportedMethodError
import httpx

class GraphAPIClient():
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )

    def _get_headers(self) -> dict[str, str]:

        # Microsoft caches the token, so getting token repeatedly
        # should not be a problem. Automatic refresh is handled
        # this way.
        token = self._credential.get_token(SCOPE)

        return {
                "Authorization": f"Bearer {token.token}",
                "Content-Type": "application/json"
            }

    def _execute_request(self, method: HTTPMethod, 
        endpoint: str, **kwargs) -> httpx.Response:

        # Map selected method to HTTPX method
        func_map = {
            "GET": httpx.get,
            "POST": httpx.post
        }

        try:
            httpx_func = func_map[method]

        except KeyError:
            raise UnsupportedMethodError

        return httpx_func(endpoint, **kwargs)

    def _unpack_response(self, response: httpx.Response) -> dict:

        response.raise_for_status()

        return response.json()

    def make_request(self, method: HTTPMethod, endpoint: str, **kwargs) -> dict:

        response = self._execute_request(
            method=method,
            endpoint=endpoint,
            **kwargs
            )

        return self._unpack_response(response)
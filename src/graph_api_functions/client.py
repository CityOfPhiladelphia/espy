# client.py
from azure.identity import ClientSecretCredential
from graph_api_functions.constants import SCOPE
from graph_api_functions.models import HTTPMethod, UnsupportedMethodError
import httpx

class GraphAPIClient():
    def __init__(self, credential: ClientSecretCredential):
        self.credential = credential

    def _get_headers(self) -> dict[str, str]:
        """
        Creates the header needed to authenticate each request to the
        API.
        """
        # Microsoft caches the token, so getting token repeatedly
        # should not be a problem. Automatic refresh is handled
        # this way.
        token = self.credential.get_token(SCOPE)

        return {
                "Authorization": f"Bearer {token.token}",
                "Content-Type": "application/json"
            }

    def _execute_request(self, method: HTTPMethod, 
        endpoint: str, **kwargs) -> httpx.Response:
        """
        Private method.

        Executes an API request. Takes one of a predetermined
        set of methods, and executes using the httpx library.

        Args:
            method (HTTPMethod): An HTTPMethod Enum
            endpoint (str): The API endpoint to perform the operation on
            **kwargs
        
        Returns:
            HTTPX response
        """
        # Map selected method to HTTPX method
        func_map = {
            "GET": httpx.get,
            "POST": httpx.post,
            "PUT": httpx.put,
            "PATCH": httpx.patch
        }

        try:
            httpx_func = func_map[method]

        except KeyError:
            raise UnsupportedMethodError

        headers = self._get_headers()

        return httpx_func(endpoint, headers=headers, **kwargs)

    def _unpack_response(self, response: httpx.Response) -> dict:
        """
        Private method.

        Unpacks an httpx response into a json object. Throws an error
        if not successful.

        """

        response.raise_for_status()

        return response.json()

    #TODO: Does this need to be 3 functions?
    def make_request(self, method: HTTPMethod, endpoint: str, **kwargs) -> dict:
        """
        Public method.

        Makes a request using HTTPX, and returns the unpacked json of that
        request.

        Args:
            method (HTTPMethod): An HTTPMethod Enum
            endpoint (str): The API endpoint to perform the operation on
            **kwargs
        """
        response = self._execute_request(
            method=method,
            endpoint=endpoint,
            **kwargs
            )

        return self._unpack_response(response)  
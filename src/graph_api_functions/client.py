# client.py
import citygeo_secrets as cgs
import httpx
from azure.identity import ClientSecretCredential
from enum import StrEnum

from graph_api_functions.constants import GRAPH_APP, SCOPE
from graph_api_functions.models import HTTPMethod, UnsupportedMethodError

class ClientEndpoints(StrEnum):
    SITE_ID = "{graph_url}/sites/{hostname}:{site_path}"

    DRIVE_ID = "{graph_url}/drives/{site_id}"

class GraphAPIClient:
    #TODO: Function that gets id information from ClientEndpoints
    def __init__(self, credential: ClientSecretCredential):
        self.credential = credential

    @staticmethod
    def build_client_secret_credential(creds: dict) -> ClientSecretCredential:
        tenant_id = creds["Tenant ID"]
        client_id = creds["Application ID"]
        client_secret = creds["Secret Value"]

        return ClientSecretCredential(tenant_id, client_id, client_secret)

    @classmethod
    def authenticate(cls):
        """Authenticate to SharePoint by generating the Client Secret
        credential."""
        creds = cgs.get_secrets(GRAPH_APP)[GRAPH_APP]
        credential = cls.build_client_secret_credential(creds)

        return cls(credential)

    def _get_headers(self) -> dict[str, str]:
        """
        Creates the header needed to authenticate each request to the
        API.

        Returns:
            dict representing the header to send in the request
        """
        # Microsoft caches the token, so getting token repeatedly
        # should not be a problem. Automatic refresh is handled
        # this way.
        token = self.credential.get_token(SCOPE)

        return {
            "Authorization": f"Bearer {token.token}",
            "Content-Type": "application/json",
        }

    def _execute_request(
        self, method: HTTPMethod, endpoint: str, **kwargs
    ) -> httpx.Response:
        """
        Private method.

        Executes an API request. Takes one of a predetermined
        set of methods, and executes using the httpx library.

        Args:
            method (HTTPMethod): An HTTPMethod Enum
            endpoint (str): The API endpoint to perform the operation on
            **kwargs: can set additional parameters such as params, timeout,
                      follow_redirects, and json

        Returns:
            HTTPX response
        """
        # Map selected method to HTTPX method
        func_map = {
            "GET": httpx.get,
            "POST": httpx.post,
            "PUT": httpx.put,
            "PATCH": httpx.patch,
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

        Args:
            response: the raw json of the httpx response.

        Returns:
            dict: the response as a dictionary.
        """

        response.raise_for_status()

        return response.json()

    # TODO: Does this need to be 3 functions?
    def make_request(self, method: HTTPMethod, endpoint: str, **kwargs) -> dict:
        """
        Public method.

        Makes a request using HTTPX, and returns the unpacked json of that
        request.

        Args:
            method (HTTPMethod): An HTTPMethod Enum
            endpoint (str): The API endpoint to perform the operation on
            **kwargs: can set additional parameters such as params, timeout,
                      follow_redirects, and json

        Returns:
            dict: dictionary of the httpx response
        """
        response = self._execute_request(
            method=method, endpoint=endpoint, **kwargs
        )

        return self._unpack_response(response)



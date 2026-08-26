# client.py
import citygeo_secrets as cgs
import httpx
from azure.identity import ClientSecretCredential
from enum import StrEnum

from graph_api_functions.constants import GRAPH_APP, SCOPE, GRAPH_URL
from graph_api_functions.urls import build_url
from graph_api_functions.models.models import HTTPMethod, UnsupportedMethodError


class ClientEndpoints(StrEnum):
    SITE_ID = "{graph_url}/sites/{hostname}:/sites/{site_name}"

    DRIVE_ID = "{graph_url}/sites/{site_id}/drives"

class GraphAPIClient:
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
        # TODO: Change this to accept a dictionary of creds since people will have different graph apps 
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

    # def _unpack_response(self, response: httpx.Response) -> dict:
    #     """
    #     Private method.

    #     Unpacks an httpx response into a json object. Throws an error
    #     if not successful.

    #     Args:
    #         response: the raw json of the httpx response.

    #     Returns:
    #         dict: the response as a dictionary.
    #     """


    #     return response.json()

    def make_request(self, method: HTTPMethod, endpoint: str, **kwargs):
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

        return response

    def get_site_id(self, hostname:str, site_name:str)-> str:
        site_id_url = build_url(
                    ClientEndpoints.SITE_ID,
                    graph_url=GRAPH_URL,
                    hostname=hostname,
                    site_name=site_name,
                )

        response = self.make_request(HTTPMethod.GET, site_id_url).json()

        return response["id"]

    def get_drive_id(self, site_id: str, document_library: str) -> str:
        drive_id_url = build_url(
            ClientEndpoints.DRIVE_ID,
            graph_url=GRAPH_URL,
            site_id=site_id
        )
        drives = self.make_request(HTTPMethod.GET, drive_id_url).json()["value"]

        for drive in drives:
            if drive["name"] == document_library:
                return drive['id']


        raise RuntimeError(
            f"Document library '{document_library}' not found."
        )
# client.py
import citygeo_secrets as cgs
import httpx
from azure.identity import ClientSecretCredential

from graph_api_functions.constants import GRAPH_APP, GRAPH_URL, SCOPE
from graph_api_functions.models import HTTPMethod, UnsupportedMethodError


class GraphAPIClient():
    def __init__(self, credential: ClientSecretCredential):
        self.credential = credential

    @staticmethod
    def build_client_secret_credential(creds: dict) -> ClientSecretCredential:
        tenant_id               = creds["Tenant ID"]
        client_id               = creds["Application ID"]
        client_secret           = creds["Secret Value"]

        return ClientSecretCredential(tenant_id, client_id, client_secret)

    @classmethod
    def authenticate(cls):
        """Authenticate to SharePoint by generating the Client Secret
        credential."""
        creds      = cgs.get_secrets(GRAPH_APP)[GRAPH_APP]
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

        Args:
            response: the raw json of the httpx response.

        Returns: 
            dict: the response as a dictionary.
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
            **kwargs: can set additional parameters such as params, timeout, 
                      follow_redirects, and json
            
        Returns:
            dict: dictionary of the httpx response 
        """
        response = self._execute_request(
            method=method,
            endpoint=endpoint,
            **kwargs
            )

        return self._unpack_response(response)  


class _UrlConstructor:
    """
    A class containing static methods to format URLs
    in the shape needed for the SharePoint API
    """
    @staticmethod
    def site_id_url(hostname: str, site_path) -> str:
        url = (
            f"{GRAPH_URL}/sites/"
            f"{hostname}:{site_path}"
        )

        return url

    @staticmethod
    def drive_id_url(site_id: str) -> str:
        url = f"{GRAPH_URL}/drives/{site_id}"

        return url

    @staticmethod
    def workbook_id_url(drive_id: str, workbook_path: str) -> str:
        url = (
                f"{GRAPH_URL}/drives/{drive_id}"
                f"/root:/{workbook_path}"
            )

        return url

    @staticmethod
    def list_id_url(site_id: str, list_name: str) -> str:

        url = (
            f"{GRAPH_URL}/sites/"
            f"{site_id}/lists/{list_name}"
            )

        return url


class _URLResolver:
    """
    A class that returns SharePointGraph API object ids.

    Args:
        client: A Microsoft GraphAPIClient instance
    """
    def __init__(self, client: GraphAPIClient):
         self.client = client

    def get_site_id(self, hostname: str, site_path: str) -> str:
        """
        Grabs the site_id of the SharePoint site specified by site_path. 
        """
        site_id_url = _UrlConstructor.site_id_url(hostname, site_path)

        data = self.client.make_request(HTTPMethod.GET, site_id_url)

        return data["id"]

    def get_drive_id(self, site_id: str, document_library: str) -> str:
        """
        Gets the id of the document library drive.
        Searches through all available drives until it finds the one specified 
        by the document_library attribute. 

        Args:
            site_id(str): the id of the sharepoint site
            document_library(str): the name of the document library you want to access  

        Returns:
            str: The id of the drive 

        Throws Error if library is not found 
        """

        drive_url = _UrlConstructor.drive_id_url(site_id)

        drives_response = self.client.make_request(HTTPMethod.GET, drive_url)

        drives = drives_response["value"]

        result = next(
            (drive for drive in drives \
             if drive.get("name") == document_library),
            None)

        if result:
            return result["id"]

        raise RuntimeError(
            f"Document library '{document_library}' not found."
        )


    def get_workbook_id(self, drive_id: str, workbook_path: str) -> str:
        """
        Gets the id of the workbook for the specified workbook_path 

        Args:
        drive_id(str): the id of the document drive that holds the workbook 
        workbook_path(str): the path to the workbook

        Returns:
        str: the id of the workbook 
        """

        workbook_url = _UrlConstructor.workbook_id_url(drive_id, workbook_path)

        data = self.client.make_request(
            HTTPMethod.GET, workbook_url)

        return data["id"]

    def get_list_id(self, site_id: str, list_name: str) -> str:
        """
        Gets the list id based on the title of the list.

        Args:
            site_id(str): The id of the sharepoint site
            list_name(str): The name of the list

        Returns:
            str: the id of the list
        """

        list_url = _UrlConstructor.list_id_url(site_id, list_name)

        data = self.client.make_request(HTTPMethod.GET, list_url)

        return data["id"]
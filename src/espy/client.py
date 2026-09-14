import os
from enum import StrEnum

import citygeo_secrets as cgs
import httpx
from azure.identity import ClientSecretCredential

from espy.constants import GRAPH_APP, GRAPH_URL, HOST_NAME, SCOPE
from espy.models.models import HTTPMethod, UnsupportedMethodError
from espy.urls import build_url


class ClientEndpoints(StrEnum):
    SITE_ID = "{graph_url}/sites/{hostname}:/sites/{site_name}"

    DRIVE_ID = "{graph_url}/sites/{site_id}/drives"

    UPLOAD_FILE = (
        "{graph_url}/sites/{site_id}"
        "/drive/root:/{dest_path}/{file_name}:/content"
        )

    GET_CONTENT = (
        "{graph_url}/sites/{site_id}"
        "/drives/{drive_id}/root:/{file_path}:/content"
    )


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
            "GET"   : httpx.get,
            "POST"  : httpx.post,
            "PUT"   : httpx.put,
            "PATCH" : httpx.patch,
            "DELETE": httpx.delete
        }

        try:
            httpx_func = func_map[method]

        except KeyError:
            raise UnsupportedMethodError

        headers = self._get_headers()

        return httpx_func(endpoint, headers=headers, **kwargs)

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

        try:
            response.raise_for_status()

        except httpx.HTTPError as e:
            error_msg = f"""{e} | Server Error Body: 
            {response.json()}"""

            raise httpx.HTTPError(error_msg) from e
        
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

    def upload_local_file(self, site_name:str, local_path:str, dest_path:str) -> dict: 
        """
        Upload a file to a SharePoint Documents folder. 

        Args:
            site_name (str) : The name of the sharepoint site you want to upload a file to 
            local_path (str): the exact local path where your file is 
            dest_path (str): the path, relative to the Documents folder, where you want to save the file
            
            Example: Setting dest_path="FolderName"
            will create a file found at Documents/FolderName/file.xlsx 
            Setting dest_path to "" will save it at Documents/file.xslx 

        Returns:
            dict: response json
        """

        file_name = os.path.basename(local_path)

        with open(local_path, "rb") as f:
            file_data = f.read()

        upload_url = build_url(
                    ClientEndpoints.UPLOAD_FILE,
                    graph_url=GRAPH_URL,
                    site_id=self.get_site_id(hostname=HOST_NAME, site_name=site_name),
                    dest_path=dest_path,
                    file_name=file_name                    
                )

        response = self.make_request(HTTPMethod.PUT, 
                                     upload_url,
                                     data=file_data)

        return response.json()

    def get_content(self, file_path:str) -> bytes:
        """
        Get the raw bytes of the excel workbook specified by
        file_path. 

        Args:
            file_path (str): The path of the file you want, relative to the
            instantiated document library. 

        Returns:
            bytes: The raw bytes representing the excel file.
        """
        content_url = build_url(ClientEndpoints.GET_CONTENT,
                                 graph_url=GRAPH_URL,
                                 site_id=self.site_id,
                                 drive_id=self.drive_id,
                                 file_path=file_path
         )

        request = self.make_request(HTTPMethod.GET,
                                            content_url, 
                                            timeout=60, 
                                            follow_redirects=True)

        return request.content 
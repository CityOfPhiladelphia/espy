"urls.py"

from graph_api_functions.client import GraphAPIClient
from graph_api_functions.constants import GRAPH_URL
from graph_api_functions.models import SharePointURL, HTTPMethod

class UrlConstructor():
    """
    A class containing static methods to format URLs
    in the shape needed for the SharePoint API
    """
    @staticmethod
    def site_url(hostname: str, site_path) -> SharePointURL:
        url = (
            f"{GRAPH_URL}/sites/"
            f"{hostname}:{site_path}"
        )

        return SharePointURL(url=url)

    @staticmethod
    def drive_url(site_id: str) -> SharePointURL:
        url = f"{GRAPH_URL}/drives/{site_id}"

        return SharePointURL(url=url)

    @staticmethod
    def workbook_url(drive_id: str, workbook_path: str) -> SharePointURL:
        url = (
                f"{GRAPH_URL}/drives/{drive_id}"
                f"/root:/{workbook_path}"
            )

        return SharePointURL(url=url)

    @staticmethod
    def list_url(site_id: str, list_name: str) -> SharePointURL:

        url = (
            f"{GRAPH_URL}/sites/"
            f"{site_id}/lists/{list_name}"
            )

        return SharePointURL(url=url)

class URLResolver():
    def __init__(self, client: GraphAPIClient):
         self.client = client

    def get_site_id(self, hostname: str, site_path: str) -> str:
        """
        Grabs the site_id of the SharePoint site specified by site_path. 
        """
        site_id_url = UrlConstructor.site_url(hostname, site_path).url

        data = self.client.make_request(HTTPMethod.GET, site_id_url)

        return data["id"]

    def get_drive_id(self, site_id: str, document_library) -> str:
        """
        Gets the id of the document library drive.
        Searches through all available drives until it finds the one specified 
        by the document_library attribute. 

        Arguments:
            site_id - the id of the sharepoint site
    
        Returns:
            str - The id of the drive 

        Throws Error if library is not found 
        """

        drive_url = UrlConstructor.drive_url(site_id).url

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

        Arguments:
        drive_id - the id of the document drive that holds the workbook 
        workbook_path - the path to the workbook

        Returns:
        str - the id of the workbook 
        """

        workbook_url = UrlConstructor.workbook_url(drive_id, workbook_path).url

        data = self.client.make_request(HTTPMethod.GET, workbook_url)
        
        return data["id"]

    def get_list_id(self, site_id: str, list_name: str) -> str:
        """
        Gets the list id based on the title of the list.

        Arguments:
            site_id - the id of the sharepoint site
            list_name - The name of the list
        
        Returns:
        str - the id of the list
        """

        list_url = UrlConstructor.list_url(site_id, list_name).url

        data = self.client.make_request(HTTPMethod.GET, list_url)

        return data["id"]
from graph_api_functions.models import SharePointURL
from graph_api_functions.constants import GRAPH_URL

class UrlConstructor():

    @staticmethod
    def site_url(hostname: str, site_path) -> SharePointURL:
        url = (
            f"{GRAPH_URL}/sites/"
            f"{hostname}:{site_path}"
        )

        return SharePointURL(url=url)

    @staticmethod
    def drive_url(drive_id: str) -> SharePointURL:
        url = f"{GRAPH_URL}/drives/{drive_id}"

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
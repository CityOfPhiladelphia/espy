from graph_api_functions.client import GraphAPIClient, ClientEndpoints
from graph_api_functions.constants import GRAPH_URL
from graph_api_functions.list import ListEndpoints


import string

from graph_api_functions.models import HTTPMethod


def build_url(url: ListEndpoints, **kwargs) -> str:
    """
    Given a GraphAPI endpoint template, builds the actual URL to request
    against. Returns an error if keys are missing.

    Args:
        url: ListEndpoints, A ListEndpoints enum object. 
            Must be a valid endpoint, with placeholders.
        **kwargs: Key word arguments that fill out the templated values contained
        in url.

    Returns:
        str, a formatted string
    """

    # get all placeholder fields in the string
    fields = {
        field for _, field, _, _ in string.Formatter().parse(url) if field
    }

    # If the caller has forgotten a field, raise an error
    missing_fields = fields - kwargs.keys()

    if missing_fields:
        raise KeyError(
            f"The following fields are missing: {', '.join(missing_fields)}"
        )

    # Return url with correct keyword arguments passed in
    return url.format(**kwargs)

# class _UrlConstructor:
#     """
#     A class containing static methods to format URLs
#     in the shape needed for the SharePoint API
#     """

#     @staticmethod
#     def site_id_url(hostname: str, site_path) -> str:
#         url = f"{GRAPH_URL}/sites/{hostname}:{site_path}"

#         return url

#     @staticmethod
#     def drive_id_url(site_id: str) -> str:
#         url = f"{GRAPH_URL}/drives/{site_id}"

#         return url

#     @staticmethod
#     def workbook_id_url(drive_id: str, workbook_path: str) -> str:
#         url = f"{GRAPH_URL}/drives/{drive_id}/root:/{workbook_path}"

#         return url

#     @staticmethod
#     def list_id_url(site_id: str, list_name: str) -> str:

#         url = f"{GRAPH_URL}/sites/{site_id}/lists/{list_name}"

#         return url

#     @staticmethod
#     def list_url(site_id: str, list_id: str) -> str:

#         url = f"{GRAPH_URL}/sites/{site_id}/lists/{list_id}/items"

#         return url



import citygeo_secrets as cgs
import httpx
from azure.identity import ClientSecretCredential

GRAPH_APP = "AppReg: CityGeo-Databridge-Updates (All Fields)"
GRAPH_URL = "https://graph.microsoft.com/v1.0"
creds = cgs.get_secrets(GRAPH_APP)

"""
TO DO: 
1. Make subclasses of GraphAPI class to delineate SharePoint Excel objects from SharePoint Lists
---> There would be methods for GraphAPI class such connecting, _get, _post, etc that would be inherited 
---> But there would be more granular classes for each objects specific methods. 

2. Look into Pydantic for validation and what not 

3. Flesh out type annotation better --> Need to import that future library for custom annotatiosn 

4. Flesh out documentation in more detail 

5. Some methods are calling httpx directly from when I was testing -- force those to use the _get method tho. 

6. Add functions for loading the pulled data into a database 
---> Want to use generators with StringIO buffering and chunking in conjuction with csv.DictReader to achieve this 
---> This is the most space efficient approach since it relies on memory and not disk space.

7. Make __init__ files so that we can import this and use the functions 
"""


class GraphAPI:
    #############################################
    ##### THESE ARE CLASS ATTRIBUTES THAT #######
    ##### WILL BE SHARE AMONGST ALL SUBS  #######
    #############################################
    tenant_id = creds[GRAPH_APP]["Tenant ID"]
    client_id = creds[GRAPH_APP]["Application ID"]
    client_secret = creds[GRAPH_APP]["Secret Value"]

    credential = ClientSecretCredential(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
    )

    token = credential.get_token("https://graph.microsoft.com/.default")
    headers = {
        "Authorization": f"Bearer {token.token}",
        "Content-Type": "application/json",
    }

    def __init__(
        self,
        site_path,
        document_library=None,
        wb_path=None,
        table_name=None,
        worksheet_name=None,
        list_name=None,
    ):
        ###############################################
        ## THIS CONSTRUCTOR WILL NEED TO BE REWORKED ##
        ## NAMELY GETTING RID OF ATTRIBS NOT NEEDED  ##
        ##              IN THE BASE CLASS            ##
        ###############################################
        self.hostname = "phila.sharepoint.com"
        self.site_path = site_path
        self.document_library = document_library
        self.workbook_path = wb_path
        self.table_name = table_name
        self.worksheet_name = worksheet_name
        self.list_name = list_name

        self.site_id = None
        self.drive_id = None
        self.item_id = None
        self.list_id = None

        self.connect()

    def _get(self, url: str, params: dict | None = None) -> dict:
        """
        Generic function for requesting a specific url.

        Params:
        url - The url you are requesting

        Returns:
        dict - The GET response
        """
        request = httpx.get(
            url,
            headers=GraphAPI.headers,
            params=params,
            timeout=60,
            follow_redirects=True,
        )
        request.raise_for_status()

        return request.json()

    def _post(self, url: str, body: dict):
        """
        Generic function for posting body to a specific url.

        Params:
        url  - The url you want to post to
        body - json of the body of the post message
        """
        request = httpx.post(
            url, headers=GraphAPI.headers, json=body, timeout=60
        )
        request.raise_for_status()

    def _patch(self, url: str, body: dict):
        """
        Generic function for making a PATCH request

        Params:
        url  - The url you want to patch to
        body - json of the body of the post message
        """

        # THIS FUNCT WAS NOT FLESHED OUT -- MIGHT REQUIRE DIFFERENT HEADERS
        request = httpx.patch(
            url, headers=GraphAPI.headers, json=body, timeout=60
        )
        request.raise_for_status()

    def get_site_id(self) -> str:
        """
        Grabs the site_id of the SharePoint site specified by site_path.
        """
        url = f"{GRAPH_URL}/sites/{self.hostname}:{self.site_path}"

        data = self._get(url)

        return data["id"]

    # ERMAGHERD THE BELLOW FUNCTIONS ARE SPECIFIC TO SUB CLASSES SO
    # THEY ARE COMMENTED OUT HERE ERMAGHERD

    # def get_drive_id(self, site_id) -> str:
    #     '''
    #     Gets the id of the document library drive.
    #     Searches through all available drives until it finds the one specified
    #     by the document_library attribute.

    #     Arguments:
    #     site_id - the id of the sharepoint site

    #     Returns:
    #     str - The id of the drive

    #     Throws Error if library is not found
    #     '''
    #     url = f"{GRAPH_URL}/sites/{site_id}/drives"
    #     drives = self._get(url)["value"]

    #     for drive in drives:
    #         if drive["name"] == self.document_library:
    #             return drive["id"]

    #     raise RuntimeError(
    #         f"Document library '{self.document_library}' not found."
    #     )

    # def get_workbook_id(self, drive_id: str):
    #     '''
    #     Gets the id of the workbook for the specified workbook_path

    #     Arguments:
    #     drive_id - the id of the document drive that holds the workbook

    #     Returns:
    #     str - the id of the workbook
    #     '''
    #     url = (
    #         f"{GRAPH_URL}/drives/{drive_id}"
    #         f"/root:/{self.workbook_path}"
    #     )

    #     item = self._get(url)

    #     return item["id"]

    # def get_list_id(self) -> str:
    #     '''
    #     Gets the list id based on the title of the list.

    #     Arguments:
    #     title - the title of the list

    #     Returns:
    #     str - the id of the list
    #     '''
    #     url = (
    #         f"{GRAPH_URL}/sites/"
    #         f"{self.site_id}/lists/{self.list_name}"
    #     )
    #     item = self._get(url)

    #     return item["id"]

    # def connect(self):
    #     '''
    #     Function to establish the necessary id's for a sharepoint resource.
    #     '''
    #     if self.site_id is not None:
    #         return

    #     self.site_id = self.get_site_id()

    #     if self.list_name:
    #         self.list_id = self.get_list_id()

    #     if self.document_library:
    #         self.drive_id = self.get_drive_id(self.site_id)

    #     if self.workbook_path is not None and self.table_name is not None and self.worksheet_name is not None:
    #         self.item_id = self.get_workbook_id(self.drive_id)

    # def protect_worksheet(self, password: str):
    #     '''
    #     Turns on sheet protection for the specified worksheet

    #     Arguments:
    #     password - str that represents the sheet protection password
    #     '''
    #     if password is None:
    #         print("No password specified, not doing anything")
    #         return

    #     if self.worksheet_name is None:
    #         print("No worksheet has been specified. Doing nothing...")
    #         return

    #     print("Re-Protecting sheet...")
    #     url = (
    #         f"{GRAPH_URL}/drives/{self.drive_id}"
    #         f"/items/{self.item_id}"
    #         f"/workbook"
    #         f"/worksheets/{self.worksheet_name}"
    #         f"/protection/protect"
    #     )

    #     body = {
    #         "password": password
    #     }

    #     self._post(url, body)

    #     print("Protection re-enabled!!!")

    # def unprotect_worksheet(self, password):
    #     '''
    #     Function that turns off protection for the worksheet.

    #     Arguments:
    #     password - str that represents the sheet protection password
    #     '''
    #     print("Unprotecting sheet...")

    #     if password is None:
    #         print("No password specified, not doing anything")
    #         return

    #     if self.worksheet_name is None:
    #         print("No worksheet has been specified. Doing nothing...")
    #         return

    #     url = (
    #         f"{GRAPH_URL}/drives/{self.drive_id}"
    #         f"/items/{self.item_id}"
    #         f"/workbook"
    #         f"/worksheets/{self.worksheet_name}"
    #         f"/protection/unprotect"
    #     )

    #     body = {
    #         "password": password
    #     }

    #     self._post(url, body)

    #     print("Unprotect successful!")

    # def append_row(self, row: List, password=None):
    #     '''
    #     Appends data, represented by row, to the sharepoint workbook table.

    #     Arguments:
    #     row - List representing the data to be appended to the sharepoint table
    #     '''

    #     self.unprotect_worksheet(password)
    #     url = (
    #         f"{GRAPH_URL}/drives/{self.drive_id}"
    #         f"/items/{self.item_id}"
    #         f"/workbook/tables/{self.table_name}"
    #         f"/rows/add"
    #     )

    #     body = {
    #         "values": row
    #     }

    #     self._post(url, body)

    #     print("Data appended successfully!")
    #     self.protect_worksheet(password)

    # def upload_file(self, local_path: str, dest_folder: str) -> dict:
    #     """
    #     Upload a file to a SharePoint document library folder.

    #     Arguments -- whered my shit go?!?!?!
    #     """
    #     site_id     = self.site_id
    #     headers     = GraphAPI.headers

    #     file_name = os.path.basename(local_path)

    #     with open(local_path, "rb") as f:
    #         file_data = f.read()

    #     upload_url = (
    #         f"https://graph.microsoft.com/v1.0/sites/{site_id}"
    #         f"/drive/root:/{dest_folder}/{file_name}:/content"
    #     )

    #     print(f"Uploading '{file_name}' -> {dest_folder}/{file_name}")

    #     resp = httpx.put(upload_url, headers=headers, data=file_data)
    #     resp.raise_for_status()
    #     return resp.json()

    # def get_sharepoint_content(self, file_path, save_path):
    #     '''
    #     Downloads sharepoint excel at file_path for the instantiated sharepoint page. Saves
    #     as new excel file at save_path. Sorry this doc is booty, im exhausted...

    #     Arguments:
    #     file_path - path to the file we want to steal
    #     save_path - path to where we want to store it.
    #     '''
    #     site_id = self.site_id
    #     drive_id = self.drive_id

    #     file_content_url = f"{GRAPH_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}:/content"
    #     request = httpx.get(file_content_url, headers=GraphAPI.headers, timeout=60, follow_redirects=True)

    #     content =  request.content

    #     content_stream = BytesIO(content)
    #     workbook = load_workbook(content_stream)
    #     workbook.save(save_path)

    # def get_list_content(self, extract_cols: list | None =None):
    #     '''
    #     Return the content of the list items by returning a generator of dicts.
    #     Each dictionary can contain multiple rows of the List and are later unpacked.
    #     Handles pagination if the result set is not complete.

    #     Arguments:
    #     extract_cols - a list of columns names you wish to extract

    #     Returns:
    #     Iterator[dict] - Returns iterator of dictionaries for each response.
    #     '''
    #     site_id = self.site_id
    #     list_id = self.list_id

    #     list_content_url = (
    #         f"{GRAPH_URL}/sites/{site_id}"
    #         f"/lists/{list_id}/items"
    #     )

    #     params = {
    #         "$expand": "fields",
    #         "$top": "300",
    #     }

    #     page = 0
    #     while list_content_url:
    #         page += 1
    #         print(f"Fetching page {page} of list items...")
    #         data = self._get(list_content_url, params if page == 1 else None) # params can't be sent during subsequent calls
    #         items = data.get("value", [])

    #         for item in items:
    #             fields = item.get("fields", {})
    #             if extract_cols:
    #                 yield {key: fields[key] for key in extract_cols}
    #             else:
    #                 yield fields

    #         list_content_url = data.get("@odata.nextLink")


if __name__ == "__main__":
    """Appending data to a specified workbook example"""
    # site_path        = "/sites/ps360-metrics-share"
    # document_library = "Documents"
    # workbook_path    = 'etl_tools_test_workbook.xlsx'
    # table_name       = 'testing'
    # worksheet_name   = 'Dataset'
    # sp_client = GraphAPI(
    #     site_path,
    #     document_library,
    #     workbook_path,
    #     table_name,
    #     worksheet_name
    # )

    # data = [[1,2,3],
    #         [4,5,6]]

    # sp_client.append_row(data, password="yo password here")

    """ Uploading a file to a sharepoint site example: """
    # site_path        = "/sites/ps360-metrics-share"
    # document_library = "Documents"
    # local_path       = "/home/ubuntu/Repos/testing/lol.xlsx"
    # dest_folder      = document_library

    # sp_client = GraphAPI(site_path, document_library)
    # sp_client.upload_file(local_path, dest_folder)

    """Downloading Sharepoint excel example:"""
    # site_path        = "/sites/ps360-metrics-share"
    # document_library = "Documents"
    # file_path        = 'Philly Stat - Law/rtk_requests.xlsx'
    # save_path        = "haha.xlsx"
    # sp_client = GraphAPI(site_path, document_library)

    # sp_client.get_sharepoint_content(file_path, save_path)

    """Extracting from Microsoft Lists"""
    site_path = "/sites/ps360-metrics-share"
    list_name = "testing_lists"
    sp_client = GraphAPI(site_path, list_name=list_name)

    extract_cols = [
        "text_col",
        "choice_col",
        "date_col",
        "mult_line_col",
        "yes_no_col",
        "location_col",
        "image_col",
    ]
    raw_items = list(sp_client.get_list_content(extract_cols))
    print(raw_items)

from enum import StrEnum

from httpx import Response 

from graph_api_functions.client import GraphAPIClient
from graph_api_functions.constants import GRAPH_URL
from graph_api_functions.models.models import HTTPMethod
from graph_api_functions.models.graph_api_models import MalformedRowError, PrimaryKeyValueNotFound
from graph_api_functions.urls import build_url


class ExcelEndpoints(StrEnum):
    """
    Endpoint registry for all API operations made by the ExcelWorkbook
    class.

    Args:
        StrEnum (StrEnum): Inherits from the StrEnum class in the enum
        library.
    """
    WORKBOOK_ID = "{graph_url}/drives/{drive_id}/root:/{workbook_path}"

    PROTECT = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/worksheets/{worksheet_name}/protection/protect"
    )
    
    UNPROTECT = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/worksheets/{worksheet_name}/protection/unprotect"
    )

    ADD_ROW = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/tables/{table_name}/rows/add"
    )

    DELETE_ROW = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/tables/{table_name}/rows/itemAt(index={index})"
    )

    UPDATE_ROW = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/tables/{table_name}/rows/itemAt(index={index})"
    )

    LIST_ROWS = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/tables/{table_name}/rows"
    )

    LIST_COLS = (
        "{graph_url}/drives/{drive_id}/items/{workbook_id}"
        "/workbook/tables/{table_name}/columns"
    )

    
class ExcelWorksheet:
    """
    Models a SharePoint Excel Workbook.

    Provides functionality to access and add rows to a Sharepoint Excel file. 

    Attributes:
        client(GraphAPIClient): A GraphAPIClient instance, used to
        make requests against the Graph API.

        site_id (str)         : A string id for the SharePoint site the graph is on.

        drive_id (str)        : A string identifier for a specfic sharepoint drive. 

        workbook_id (str)     : A string identifier for a specfic excel workbook.

        worksheet_name (str)  : String representing the name of a worksheet in Excel.

        table_name (str)      : String representing the name of a table object in Excel. 
    """
    def __init__(
        self,
        client: GraphAPIClient,
        site_id: str,
        drive_id: str, 
        workbook_id: str,
        worksheet_name: str|None=None,
        table_name: str|None=None,
    ):

        self.client         = client
        self.site_id        = site_id
        self.drive_id       = drive_id
        self.workbook_id    = workbook_id
        self.worksheet_name = worksheet_name
        self.table_name     = table_name

    @classmethod
    def setup(
        cls, 
        hostname:str, 
        site_name:str, 
        document_library:str,
        workbook_path:str,
        worksheet_name:str|None=None,
        table_name:str|None=None
        ):
        """
        Sets up the ExcelWorkbook class. Creates a client and fetches
        necessary ids to have ExcelWorkbook operate on. 

        Args:
            hostname        : Sharepoint hostname. i.e. "phila.sharepoint.com"
            site_name       : The name of the sharepoint site. i.e. "ps360-metrics-share"
            document_library: The name of the document library. i.e. "Documents"
            workbook_path   : The path to the file you want to access 
            relative to document_library. i.e. "Philly Stat - OIT/OIT_data.xlsx" 
            worksheet_name  : The name of the worksheet you want to affect. i.e. "Metrics" 
            table_name      : The name of the table object in the excel. i.e. "Table1" 

        Returns:
            ExcelWorkbook: An ExcelWorkboook object.
        """
        client = GraphAPIClient.authenticate()

        site_id = client.get_site_id(hostname, site_name)
        drive_id = client.get_drive_id(site_id, document_library)

        workbook_id_url = build_url(
            ExcelEndpoints.WORKBOOK_ID,
            graph_url=GRAPH_URL,
            drive_id=drive_id, 
            workbook_path=workbook_path
        )
        workbook_id = client.make_request(HTTPMethod.GET, workbook_id_url).json()["id"]

        return cls(client, site_id, drive_id, workbook_id, worksheet_name, table_name)


    def get_row(self):
        raise NotImplementedError("Method for excel be implemented in the future.")

    def add_rows(self, rows:list, password:str|None=None) -> Response:
        """
        Append a row of data to a specific table in a specific excel worksheet. 
        For this function to work: 
            a) Excel must contain a table object in it, which
            must have been set when instantiating ExcelWorksheet. 
            
            b) The list you pass must have a value for each column of the table, or 
            the operation will fail. 

        Args:
            row (list): The row you want to add, in list form. 
                e.g. [[1,2,3]] would add a row with values 1, 2, and 3. 

            password (str, optional): Password to pass to sheet protection to temporarily un/reprotect sheet
                Defaults to None.

        Returns:
            Response - httpx response 
        """
        self._check_rows(rows)

        self.toggle_protection(password, protect=False)

        add_row_url = build_url(
            ExcelEndpoints.ADD_ROW,
            graph_url=GRAPH_URL,
            drive_id=self.drive_id,
            workbook_id=self.workbook_id,
            table_name=self.table_name
        )

        json = {
            "values": rows
        }
        
        response = self.client.make_request(HTTPMethod.POST, 
                                            add_row_url, 
                                            json=json, 
                                            timeout=60)

        self.toggle_protection(password, protect=True)

        return response


    def delete_row_at_index(self, index:int, password:str|None=None) -> Response:
        """
        Deletes a row, specified by index, from an excel table. 
        Index is 0-based, so the first row in the table is set by 
        index=0,

        Args:
            index (int): The row number you want to delete, 0-based.
            password (str | None, optional): Sheet protection password. Defaults to None.

        Returns:
            Response: The HTTP response.
        """
        self.toggle_protection(password, protect=False)
        del_row_url = build_url(
            ExcelEndpoints.DELETE_ROW,
            graph_url=GRAPH_URL,
            drive_id=self.drive_id,
            workbook_id=self.workbook_id,
            table_name=self.table_name,
            index=index
        )
        response = self.client.make_request(HTTPMethod.DELETE, del_row_url)
        self.toggle_protection(password, protect=True)

        return response

    def delete_row_by_pk(self,pk_col:str, pk_val: str|int, password:str|None=None) -> Response:
        """
        Deletes a row based on a primart key column and value.

        Args:
            pk_col (str): The name of the primary key column 
            pk_val (str | int): The value we are looking to match on 
            password (str | None, optional):  Sheet protection password. Defaults to None.

        Raises:
            KeyError: Throws when key is not found as a column in the tbale 
            PrimaryKeyValueNotFound: Thrown when the value to look for doesn't exist. 

        Returns:
            Response: HTTP Response
        """
        rows = self.list_rows()

        if not len(rows):
            raise ValueError("Cannot delete from empty table!")
        
        if pk_col not in rows[0]:
            raise KeyError(f"Key {pk_col} does not exist in table!")

        del_index = None 
        for index, row in enumerate(rows): 
            if row[pk_col] == pk_val: 
                del_index = index 
                break

        if del_index:
            response = self.delete_row_at_index(del_index, password=password)
        else:
            raise PrimaryKeyValueNotFound(f"Could not find value: {pk_val} under the primary key column {pk_col}!")

        
        return response

    def update_row_at_index(self, index:int, value:list, password:str|None=None) -> Response:
        """
        Updates an existing row of an excel table. 

        Arguments: 
            index(int): The index of the row you want to update, 0-based. 
            value(list): List containing the data to send to update 
            password(str|None, optional): Sheet protection password 

        Returns:
            Response: The HTTP response. 
        """
        self.toggle_protection(password, protect=False)
        update_row_url = build_url(
            ExcelEndpoints.UPDATE_ROW,
            graph_url=GRAPH_URL,
            drive_id=self.drive_id,
            workbook_id=self.workbook_id,
            table_name=self.table_name,
            index=index
        )

        json = {
            "values": value
        }

        response = self.client.make_request(HTTPMethod.PATCH, update_row_url, json=json)
        self.toggle_protection(password, protect=True)

        return response



    def _check_rows(self, rows: list):
        num_cols = len(self.list_columns())

        for row in rows:
            if len(row) != num_cols:
                raise MalformedRowError(f"A row of data contains {len(row)} values, but requires exactly {num_cols} values.")


    def toggle_protection(self, password:str, protect:bool) -> None:
        """
        Turns on/off sheet protection in the excel worksheet. 

        Args:
            password (str): Password for sheet protection
            protect (bool): Boolean representing if you want it on (True) or off (False) 
        """
        if not password:
            return 
        
        if not self.worksheet_name:
            return 

        if not protect:
            url = build_url(
                        ExcelEndpoints.UNPROTECT,
                        graph_url=GRAPH_URL,
                        drive_id=self.drive_id,
                        workbook_id=self.workbook_id,
                        worksheet_name=self.worksheet_name
                    )
        else:
            url = build_url(
                        ExcelEndpoints.PROTECT,
                        graph_url=GRAPH_URL,
                        drive_id=self.drive_id,
                        workbook_id=self.workbook_id,
                        worksheet_name=self.worksheet_name
                    )

        json = { "password": password }

        self.client.make_request(HTTPMethod.POST, url, json=json)


    def list_rows(self) -> list[dict]:
        """
        Returns the rows of an excel table, represented as a list
        of dictionaries. The the keys to the dictionary represent 
        the column name and the value is the value of the column

        Returns:
            list[dict]: The rows of an excel table. 
        """

        column_names = self.list_columns()

        list_rows_url = build_url(
            ExcelEndpoints.LIST_ROWS,
            graph_url=GRAPH_URL,
            drive_id=self.drive_id,
            workbook_id=self.workbook_id,
            table_name=self.table_name
        )
        response = self.client.make_request(HTTPMethod.GET, list_rows_url).json()
        items = response.get("value", [])

        rows = [] 

        for item in items: 
            row = {key: item['values'][0][i] for i, key in enumerate(column_names)}
            rows.append(row)  

        return rows 


    def list_columns(self) -> list:
        """
        Return a list of column names for a specific excel table.
        Column names are returned in the order they appear in the table. 

        Returns:
            list: A list containing the column names
        """

        list_cols_url = build_url(
            ExcelEndpoints.LIST_COLS,
            graph_url=GRAPH_URL,
            drive_id=self.drive_id,
            workbook_id=self.workbook_id,
            table_name=self.table_name
        )

        response = self.client.make_request(HTTPMethod.GET, list_cols_url).json()

        values = response.get("value", [])

        cols = [col['name'] for col in values]

        return cols

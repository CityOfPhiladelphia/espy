from unittest.mock import MagicMock, call, patch
 
import pytest
 
from espy.excel import ExcelWorksheet
from espy.models.graph_api_models import (
    MalformedRowError,
    PrimaryKeyValueNotFound,
)
from espy.models.models import HTTPMethod

@pytest.fixture
def mock_client():
    return MagicMock()

@pytest.fixture
def worksheet(mock_client): # Calls the above fixture 
    return ExcelWorksheet(
        client=mock_client,
        site_id="a site",
        drive_id="a drive",
        workbook_id="a workbook",
        worksheet_name="a name",
        table_name="Table1",
    )
 
 
@pytest.fixture
def worksheet_no_protection(mock_client):
    """A worksheet with no worksheet_name, so toggle_protection short-circuits."""
    return ExcelWorksheet(
        client=mock_client,
        site_id="a site",
        drive_id="a drive",
        workbook_id="a workbook",
        worksheet_name=None,
        table_name="Table1",
    )
 
 
def make_response(json_data=None):
    response = MagicMock()
    if json_data is not None:
        response.json.return_value = json_data
    return response


### Testing the setup functions ### 
class TestSetup:
    def test_setup_with_resolved_ids(self):
        # Set up a mock client 
        mock_instance = MagicMock()
        mock_instance.get_site_id.return_value = "test_site"
        mock_instance.get_drive_id.return_value = "test_drive"
        mock_instance.make_request.return_value = make_response({"id":"some_id"})

        with patch("espy.excel.GraphAPIClient") as MockedClient: 
            MockedClient.authenticate.return_value = mock_instance

            excel = ExcelWorksheet.setup(
                hostname="garb.com",
                site_name="garbsite",
                document_library="garbdoc",
                workbook_path="garbpath",
                worksheet_name="garbsheet",
                table_name="garbtable",
            )

        assert isinstance(excel, ExcelWorksheet)
        assert excel.client is mock_instance
        assert excel.site_id == "test_site"
        assert excel.drive_id == "test_drive"
        assert excel.workbook_id == "some_id"
        assert excel.worksheet_name == "garbsheet"
        assert excel.table_name == "garbtable"

    def test_setup_with_default_vals(self):
        mock_instance = MagicMock()
        mock_instance.get_site_id.return_value = "test_site"
        mock_instance.get_drive_id.return_value = "test_drive"
        mock_instance.make_request.return_value = make_response({"id":"some_id"})
 
        with patch("espy.excel.GraphAPIClient") as MockedClient: 
            MockedClient.authenticate.return_value = mock_instance
 
            excel = ExcelWorksheet.setup(
                hostname="contoso.sharepoint.com",
                site_name="TeamSite",
                document_library="Documents",
                workbook_path="Reports/data.xlsx",
            )
 
        assert excel.worksheet_name is None
        assert excel.table_name is None


### Testing the add row functions ### 
class TestAddRows:
    def test_add_rows_success(self, worksheet, mock_client):
        with patch.object(worksheet, "list_columns", return_value=["col1", "col2"]), \
        patch.object(worksheet, "toggle_protection") as mock_toggle:
            mock_response = make_response()
            mock_client.make_request.return_value = mock_response
 
            rows = [[1, 2]]
            response = worksheet.add_rows(rows, password="pw")
 
        assert response is mock_response # Test that add rows returns what make_request returns 

        # Check that toggle protection is called twice  
        mock_toggle.assert_has_calls(
            [call("pw", protect=False), call("pw", protect=True)]
        )

        # Check make request is called once
        mock_client.make_request.assert_called_once()

        # Check the arguments were passed correctly 
        args, kwargs = mock_client.make_request.call_args
        assert args[0] == HTTPMethod.POST
        assert "Table1/rows/add" in args[1]
        assert kwargs["json"] == {"values": rows}
        assert kwargs["timeout"] == 60
 
    def test_add_rows_malformed(self, worksheet, mock_client):
        with patch.object(worksheet, "list_columns", return_value=["col1", "col2"]):
            with pytest.raises(MalformedRowError):
                worksheet.add_rows([[1, "Alice", "extradata"]])

        # Makle sure make request not called with garb rows 
        mock_client.make_request.assert_not_called()

class TestDeleteRow:
    def test_delete_row_at_index(self, worksheet, mock_client):
        with patch.object(worksheet, "toggle_protection") as mock_toggle:
            mock_response = make_response()
            mock_client.make_request.return_value = mock_response
 
            response = worksheet.delete_row_at_index(2, password="pw")
 
        assert response is mock_response # again, make sure the make_request is returned 

        # Make sure makes 2 calls to toggle 
        mock_toggle.assert_has_calls(
            [call("pw", protect=False), call("pw", protect=True)]
        )

        # Make sure arguments are good
        args, kwargs = mock_client.make_request.call_args
        assert args[0] == HTTPMethod.DELETE
        assert "index=2" in args[1]

    def test_delete_row_by_pk_delegates_to_index_delete(self, worksheet):
        with patch.object(
            worksheet, "_find_index_of_pk", return_value=3  # set return val 
        ) as mock_find, \
        patch.object(
            worksheet, "delete_row_at_index", return_value="del"
        ) as mock_delete:
            result = worksheet.delete_row_by_pk("ID", 42, password="pw")
 
        mock_find.assert_called_once_with("ID", 42)

        # Ensure delete called with right index 
        mock_delete.assert_called_once_with(3, password="pw")

        assert result == "del"

class TestUpdateRow:
    def test_update_row_at_index(self, worksheet, mock_client):
        with patch.object(
            worksheet, "list_columns", return_value=["col1", "col2"]
        ), patch.object(worksheet, "toggle_protection") as mock_toggle:
            mock_response = make_response()
            mock_client.make_request.return_value = mock_response
 
            response = worksheet.update_row_at_index(1, [[1, 2]], password="pw")
 
        assert response is mock_response
        mock_toggle.assert_has_calls(
            [call("pw", protect=False), call("pw", protect=True)]
        )
        args, kwargs = mock_client.make_request.call_args
        assert args[0] == HTTPMethod.PATCH
        assert "index=1" in args[1]
        assert kwargs["json"] == {"values": [[1, 2]]}

    def test_update_row_at_index_malformed(self, worksheet, mock_client):
        with patch.object(worksheet, "list_columns", return_value=["col1", "col2"]):
            with pytest.raises(MalformedRowError):
                worksheet.update_row_at_index(1, [[1]])
 
        mock_client.make_request.assert_not_called()


class TestListCols:
    def test_list_columns(self, worksheet, mock_client):
        mock_client.make_request.return_value = make_response(
            {"value": [{"name": "col1"}, {"name": "col2"}, {"name": "col3"}]}
        )
 
        cols = worksheet.list_columns()
 
        assert cols == ["col1", "col2", "col3"]
        args, kwargs = mock_client.make_request.call_args
        assert args[0] == HTTPMethod.GET
        assert "Table1/columns" in args[1]

    def test_list_columns_no_cols(self, worksheet, mock_client):
        mock_client.make_request.return_value = make_response({"value": []})
 
        assert worksheet.list_columns() == []

class TestCheckRows:
    def test_valid_rows(self, worksheet):
        with patch.object(worksheet, "list_columns", return_value=["col1", "col2"]):
            worksheet._check_rows([[1, 2], [3, 4]])  
 
    def test_malformed_row(self, worksheet):
        with patch.object(worksheet, "list_columns", return_value=["col1", "col2"]), \
        pytest.raises(MalformedRowError):
            worksheet._check_rows([[1]])

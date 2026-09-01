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
        table_name="the table",
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
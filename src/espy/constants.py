# constants.py
MAX_BATCH_SIZE = 20

SCOPE = "https://graph.microsoft.com/.default"
GRAPH_URL = "https://graph.microsoft.com/v1.0"
GRAPH_APP = "AppReg: CityGeo-Databridge-Updates (All Fields)"
HOST_NAME = "phila.sharepoint.com"

SHAREPOINT_LIST_EXCLUDED_COLUMNS = {
    "@odata.etag",
    "Title",
    # "ID",
    "LinkTitle",
    "_ColorTag",
    "ComplianceAssetId",
    "ContentType",
    "Modified",
    "Created",
    "Author",
    "AuthorLookupId",
    "EditorLookupId",
    "Editor",
    "_UIVersionString",
    "Attachments",
    "Edit",
    "LinkTitleNoMenu",
    "DocIcon",
    "ItemChildCount",
    "FolderChildCount",
    "_ComplianceFlags",
    "_ComplianceTag",
    "_ComplianceTagWrittenTime",
    "_ComplianceTagUserId",
    "_IsRecord",
    "Label setting",
    "AppAuthor",
    "AppAuthorLookupId",
    "AppEditor",
    "AppEditorLookupId",
}

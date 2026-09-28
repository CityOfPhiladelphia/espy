# constants.py

SCOPE = "https://graph.microsoft.com/.default"
GRAPH_URL = "https://graph.microsoft.com/v1.0"

# Environment variable read when no hostname is passed
HOSTNAME_ENV_VAR = "SHAREPOINT_HOSTNAME"

# APICredentials key -> environment variable read when no creds are passed.
# These are the standard names used by azure-identity.
CREDENTIAL_ENV_VARS = {
    "tenant_id": "AZURE_TENANT_ID",
    "client_id": "AZURE_CLIENT_ID",
    "client_secret": "AZURE_CLIENT_SECRET",
}

SHARE_POINT_LIST_EXCLUDED_COLUMNS = [
    "@odata.etag",
    "Title",
    "ID",
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
]

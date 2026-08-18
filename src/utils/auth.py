import citygeo_secrets as cgs
from azure.identity import ClientSecretCredential

from graph_api_functions.constants import GRAPH_APP

def get_graph_app_secrets() -> dict:
    # CGS nests the response under GRAPH_APP twice; unwrapping
    return cgs.get_secrets(GRAPH_APP)[GRAPH_APP]

def build_client_secret_credential(creds: dict) -> ClientSecretCredential:
    tenant_id               = creds["Tenant ID"]
    client_id               = creds["Application ID"]
    client_secret           = creds["Secret Value"]

    return ClientSecretCredential(tenant_id, client_id, client_secret)


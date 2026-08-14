# client.py
from azure.identity import ClientSecretCredential
from graph_api_functions.models import HTTPMethod

SCOPE = "https://graph.microsoft.com/.default"

class GraphAPIClient():
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret
        )

    def _get_headers(self) -> dict[str, str]:

        # Microsoft caches the token, so getting token repeatedly
        # should not be a problem. Automatic refresh is handled
        # this way.
        token = self._credential.get_token(SCOPE)

        return {
                "Authorization": f"Bearer {token.token}",
                "Content-Type": "application/json"
            }

    def execute_request(self, method: HTTPMethod, endpoint: str, json: dict):
        ...
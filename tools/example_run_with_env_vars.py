import os
from collections.abc import Callable
from contextlib import contextmanager
from example import main

# Use your own secret manager here
class FakeSecretManager:
    def __init__(self, secret_name: str):
        self.secret_name = secret_name
        
    def fake_get_secret(self) -> dict:
        """Returns a dictionary"""
        return {}



@contextmanager
def set_env_vars():
    """
    Temporarily sets environment variables within a 'with' block.
    Restores the original environment upon exit.
    """
    creds = FakeSecretManager("fake_secret").fake_get_secret()

    AZURE_TENANT_ID = creds["Tenant ID"]
    AZURE_CLIENT_ID = creds["Application ID"]
    AZURE_CLIENT_SECRET = creds["Secret Value"]
    SHAREPOINT_HOSTNAME = "phila.sharepoint.com"

    env_vars = {
        "AZURE_TENANT_ID": AZURE_TENANT_ID,
        "AZURE_CLIENT_ID": AZURE_CLIENT_ID,
        "AZURE_CLIENT_SECRET": AZURE_CLIENT_SECRET,
        "SHAREPOINT_HOSTNAME": SHAREPOINT_HOSTNAME,
    }

    os.environ.update(env_vars)

    try:
        yield

    finally:
        for key, _ in env_vars.items():
            os.environ.pop(key, None)


def run_with_env_vars(func: Callable):
    with set_env_vars():
        func()


if __name__ == "__main__":
    run_with_env_vars(main)

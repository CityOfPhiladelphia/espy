# base.py

from abc import ABC, abstractmethod
from requests import Response

class SharePointFile(ABC):
    def __init__(self, url: str):
        self.url = url

class RowwiseSharePointFile(SharePointFile, ABC):
    @abstractmethod
    def list_rows(self) -> Response:
        pass
# interfaces.py

from abc import ABC, abstractmethod
from requests import Response

class RowOperations(ABC):
    @abstractmethod
    def add_row(self) -> Response:
        raise NotImplementedError

    @abstractmethod
    def edit_row(self) -> Response:
        raise NotImplementedError

    @abstractmethod
    def delete_row(self) -> Response:
        raise NotImplementedError

    @abstractmethod
    def upsert_row(self) -> Response:
        raise NotImplementedError

    @abstractmethod
    def get_row(self) -> Response:
        raise NotImplementedError
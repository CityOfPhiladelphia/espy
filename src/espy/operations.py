from typing import Any, Dict
from pydantic.dataclasses import dataclass

@dataclass(frozen=True)
class GetRow:
    key_col: str
    value: Any

@dataclass(frozen=True)
class AddRow:
    data: Dict

@dataclass(frozen=True)
class EditRow:
    row_id: int
    data: Dict

@dataclass(frozen=True)
class DeleteRow:
    row_id: int

    



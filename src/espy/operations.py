from typing import Any, Dict, TypeAlias
from pydantic.dataclasses import dataclass

@dataclass(frozen=True)
class GetRow:
    """
    Get Row operation on a SharePoint list.
    
    Args:
        key_col (str): The primary key column of the SharePoint list.
        value (Any): The value of the key column to search for.
    """
    key_col: str
    value: Any

@dataclass(frozen=True)
class AddRow:
    """
    Add Row operation on a SharePoint list.
    
    Args:
        data (dict): A dictionary representing the row to be added to the
        SharePoint list.

    """
    data: Dict

@dataclass(frozen=True)
class EditRow:
    """
    Edit Row operation on a SharePoint list. 

    Must use row_id, which can be obtained
    through a Get Row operation.
    
    Args:
        row_id (int): The row id of the row, as provided by SharePoint.
        data (dict): A dictionary representing the new values for the row
        in the SharePoint list.
    """
    row_id: int
    data: Dict

@dataclass(frozen=True)
class DeleteRow:    
    """
    Delete Row operation on a SharePoint list.. 
    
    Must use row_id, which can be obtained
    through a Get Row operation.
    
    Args:
        row_id (int): The row id of the row, as provided by SharePoint.
    """
    row_id: int

BatchOperation: TypeAlias = GetRow | AddRow | EditRow | DeleteRow



import json
from collections.abc import Mapping
from typing import Any


def get_dict_structure(d):
    if isinstance(d, dict):
        return {k: get_dict_structure(v) for k, v in d.items()}
    elif isinstance(d, list):
        # Shows structure of the first item in the list as a template
        return [get_dict_structure(d[0])] if d else []
    else:
        return type(d).__name__


def flatten_data(data: Mapping[str, Any], parent_key: str = "") -> dict:
    items: list = []

    for key, value in data.items():
        new_key = f"{parent_key}_{key}" if parent_key else key

        # Handle cases where the dict/JSON value is in the form of a string
        if isinstance(value, str):
            value_stripped = value.strip()

            if value_stripped.startswith("{") and value_stripped.endswith("}"):
                try:
                    value = json.loads(value)

                except json.JSONDecodeError:
                    pass

        if isinstance(value, Mapping):
            items.extend(flatten_data(value, new_key).items())

        else:
            items.append((new_key, value))

    return dict(items)

import string
from enum import StrEnum


def build_url(url: StrEnum, **kwargs) -> str:
    """
    Given a GraphAPI endpoint template, builds the actual URL to request
    against. Returns an error if keys are missing.

    Args:
        url: ListEndpoints, A ListEndpoints enum object.
            Must be a valid endpoint, with placeholders.
        **kwargs: Key word arguments that fill out the templated values contained
        in url.

    Returns:
        str, a formatted string
    """

    fields = {
        field for _, field, _, _ in string.Formatter().parse(url) if field
    }

    # If the caller has forgotten a field, raise an error
    missing_fields = fields - kwargs.keys()

    if missing_fields:
        raise KeyError(
            f"The following fields are missing: {', '.join(missing_fields)}"
        )

    return url.format(**kwargs)
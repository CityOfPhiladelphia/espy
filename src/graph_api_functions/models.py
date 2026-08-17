from enum import StrEnum
from pydantic import BaseModel, HttpUrl

## Enums
class HTTPMethod(StrEnum):
    GET = "GET"
    POST = "POST"
    PUT = "PUT"

## Errors
class UnsupportedMethodError(BaseException):
    pass

## Data Validation
class SharePointURL(BaseModel):
    url: HttpUrl

## API Response Classes

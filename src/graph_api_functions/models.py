from enum import StrEnum
from pydantic import BaseModel, HttpUrl, DirectoryPath

## Enums
class HTTPMethod(StrEnum):
    GET = "GET"
    POST = "POST"
    PUT = "PUT"
    PATCH = "PATCH"

## Errors
class UnsupportedMethodError(BaseException):
    pass

## API Input Classes
class SharePointSiteInfo(BaseModel):
    hostname: HttpUrl
    site_path: DirectoryPath

## API Response Classes

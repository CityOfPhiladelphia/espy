# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [3.0.0] – 2026-10-13

### Added
- SharePointList batch methods:
    - `get_rows`
    - `add_rows`
    - `edit_rows`
    - `delete_rows`
    - Low-level `batch` method
- Documentation for the batch methods, including example code.

### Changed
- **Breaking:** SharePointList single request methods:
    - `add_row` returns a `dict`, not `httpx.Response`
    - `edit_row` returns a `dict`, not `httpx.Response`
    - `delete_row` returns `None`, not `httpx.Response`
    - `list_columns` returns a list of `SharePointListColumn` objects, no longer a list of dicts
- **Breaking:** Failed requests in `client.py` now raise `httpx.HTTPStatusError`
  instead of `httpx.HTTPError`. `HTTPStatusError` is a subclass of
  `HTTPError`, so existing `except httpx.HTTPError` handlers still work.
- **Breaking:** The constant `SHARE_POINT_LIST_EXCLUDED_COLUMNS` (a list) was
  renamed to `SHAREPOINT_LIST_EXCLUDED_COLUMNS` (a set). This only affects code
  that imports it directly.
- `edit_row` now looks up the row before validating the incoming data. Invalid
  data still raises `KeyError`, but after the lookup request has been sent.

### Security
- Credentials read from environment variables are held as `SecretStr`, so they
  are masked if an `APICredentials` object is printed or logged.

### Upgrading from 2.x

The single request methods now return parsed data rather than the raw
`httpx.Response`.

```python
# 2.x
response = sp_list.add_row(row)
new_row = response.json()

# 3.0
new_row = sp_list.add_row(row)  # already a dict
```

`edit_row` follows the same pattern as `add_row`. `delete_row` returns
`None`; a failed delete raises `httpx.HTTPStatusError`.

Because responses are no longer returned, code that used `.status_code` or
`.raise_for_status()` should instead catch `httpx.HTTPStatusError`.

`list_columns` now returns model objects, so use attribute access:

```python
# 2.x
cols = sp_list.list_columns()
cols[0]["display_name"]

# 3.0
cols[0].display_name
```

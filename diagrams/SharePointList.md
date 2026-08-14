---
title: SharePoint Interface Architecture
---

```mermaid

classDiagram

    class GraphAPIClient {
        << Client >>
        + String tenant_id
        + String client_id
        + String client_secret
        + get_headers() dict
        + execute_request(method: Enum, endpoint: str, json: dict) dict
    }

    class TabularStorage {
        << Protocol >>
        + get_row(row_id) dict
        + list_rows() List~dict~
        + add_row(data) dict
        + edit_row(row_id, data) dict
        + delete_row(row_id) bool
        + upsert_row(key_col, data) dict
    }

    class ExcelWorkbook {
        << Aggregate >>
        - GraphAPIClient client
        + String site_id
        + String file_path
        + list_tables() List~str~
        + get_table(table_name) ExcelTable
        + create_table(address, table_name) ExcelTable
    }

    class ExcelTable {
        << Entity >>
        - GraphAPIClient client
        + String site_id
        + String file_path
        + String table_name
        + get_row(row_id) dict
        + list_rows() List~dict~
        + add_row(data) dict
        + edit_row(row_id, data) dict
        + delete_row(row_id) bool
        + upsert_row(key_col, data) dict
        + lock() dict
    }

    class SharePointList {
        << Entity >>
        - GraphAPIClient client
        + String site_id
        + String list_id
        + get_row(row_id) dict
        + list_rows() List~dict~
        + add_row(data) dict
        + edit_row(row_id, data) dict
        + delete_row(row_id) bool
        + upsert_row(key_col, data) dict
    }

    %% TabularStorage implementations
    TabularStorage <|.. ExcelTable : satisfies
    TabularStorage <|.. SharePointList : satisfies

    %% Composition / Dependencies
    ExcelWorkbook o-- GraphAPIClient : uses
    ExcelWorkbook "1" *-- "many" ExcelTable : contains & creates
    ExcelTable o-- GraphAPIClient : uses
    SharePointList o-- GraphAPIClient : uses
```
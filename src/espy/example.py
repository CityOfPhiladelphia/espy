from datetime import datetime
from requests.exceptions import HTTPError

from espy.list import SharePointList
from espy.models.models import ErrorResult, BatchResponseBody, SharePointListRow 
from espy.operations import GetRow, AddRow, EditRow, DeleteRow

SITE_NAME = "ps360-metrics-share"
LIST_NAME = "PPR 311 Test List"

def test_get_rows(sp_list: SharePointList):

    key_col = "Service Request Number"
    values = ["19914185", "19914171", "19914114"]

    operations = [GetRow(key_col=key_col, value=value) for value in values]

    response = sp_list.batch(operations)

    print(response.model_dump())

def test_add_rows(sp_list: SharePointList):

    data_to_add = [
        {
            "Date/Time Opened": str(datetime.now()),
            "Address/Intersection": "1234 MARKET ST",
            "Service Request Number": "99999999",
            "Description": "Inserted as a test.",
            "Tree Inside Park or at Rec Center": None,
            "ZipCode": "19107",
            "Status": "New",
            "Notes": None
        },
                {
            "Date/Time Opened": str(datetime.now()),
            "Address/Intersection": "1234 MARKET ST",
            "Service Request Number": "99999998",
            "Description": "Inserted as a test.",
            "Tree Inside Park or at Rec Center": None,
            "ZipCode": "19107",
            "Status": "New",
            "Notes": None
        },
                {
            "Date/Time Opened": str(datetime.now()),
            "Address/Intersection": "1234 MARKET ST",
            "Service Request Number": "99999997",
            "Description": "Inserted as a test.",
            "Tree Inside Park or at Rec Center": None,
            "ZipCode": "19107",
            "Status": "New",
            "Notes": None
        }
    ]

    operations = [AddRow(data) for data in data_to_add]

    response = sp_list.batch(operations)

    print(response.model_dump())

def test_edit_rows(sp_list: SharePointList):
    key_col = "Service Request Number"
    values = ["99999999", "99999998", "99999997"]

    data_to_edit = [
        {
            "Date/Time Opened": str(datetime.now()),
            "Address/Intersection": "1234 MARKET ST",
            "Service Request Number": "99999999",
            "Description": "Inserted as a test.",
            "Tree Inside Park or at Rec Center": None,
            "ZipCode": "19107",
            "Status": "In-Progress",
            "Notes": None
        },
                {
            "Date/Time Opened": str(datetime.now()),
            "Address/Intersection": "1234 MARKET ST",
            "Service Request Number": "99999998",
            "Description": "Inserted as a test.",
            "Tree Inside Park or at Rec Center": None,
            "ZipCode": "19107",
            "Status": "In-Progress",
            "Notes": None
        },
                {
            "Date/Time Opened": str(datetime.now()),
            "Address/Intersection": "1234 MARKET ST",
            "Service Request Number": "99999997",
            "Description": "Inserted as a test.",
            "Tree Inside Park or at Rec Center": None,
            "ZipCode": "19107",
            "Status": "In-Progress",
            "Notes": None
        }
    ]

    get_operations = [GetRow(key_col=key_col, value=value) for value in values]

    canonical_key = sp_list.display_to_canonical[key_col]

    batch = sp_list.batch(get_operations)

    id_key_map = {}

    for response in batch.responses:
        match response.body:

            case BatchResponseBody():
                row_id = response.body.value[0].fields['id']
                pk_val = response.body.value[0].fields[canonical_key]

            case SharePointListRow():
                row_id = response.body.fields['id']
                pk_val = response.body.fields[canonical_key]

            case ErrorResult():
                error_code = ErrorResult.error.code
                error_message = ErrorResult.error.message

                raise HTTPError(f"{error_code}: {error_message}")

            case _:
                raise ValueError(f"""Batch endpoint returned 
                                    unexpected data format.""")

        id_key_map[pk_val] = row_id

    edit_operations = []

    for data in data_to_edit:
        pk_val = data[key_col]
        row_id = id_key_map[pk_val]
        edit_operations.append(EditRow(row_id, data))


    response = sp_list.batch(edit_operations)

    print(response.model_dump())

def test_delete_rows(sp_list: SharePointList):
    key_col = "Service Request Number"
    values = ["99999999", "99999998", "99999997"]

    get_operations = [GetRow(key_col=key_col, value=value) for value in values]

    canonical_key = sp_list.display_to_canonical[key_col]

    batch = sp_list.batch(get_operations)

    id_key_map = {}

    for response in batch.responses:
        match response.body:

            case BatchResponseBody():
                row_id = response.body.value[0].fields['id']
                pk_val = response.body.value[0].fields[canonical_key]

            case SharePointListRow():
                row_id = response.body.fields['id']
                pk_val = response.body.fields[canonical_key]

            case ErrorResult():
                error_code = ErrorResult.error.code
                error_message = ErrorResult.error.message

                raise HTTPError(f"{error_code}: {error_message}")

            case _:
                raise ValueError(f"""Batch endpoint returned 
                                    unexpected data format.""")

        id_key_map[pk_val] = row_id

    delete_operations = []

    for value in values:
        row_id = id_key_map[value]
        delete_operations.append(DeleteRow(row_id))


    response = sp_list.batch(delete_operations)

    print(response.model_dump())

def main():
    sp_list = SharePointList.setup(site_name=SITE_NAME, list_name=LIST_NAME)
    #test_get_rows(sp_list)
    #test_add_rows(sp_list)
    #test_edit_rows(sp_list)
    test_delete_rows(sp_list)

if __name__ == "__main__":
    main()
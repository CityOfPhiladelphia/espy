from datetime import datetime

from espy.list import SharePointList

SITE_NAME = "ps360-metrics-share"
LIST_NAME = "PPR 311 Test List"


def test_get_rows(sp_list: SharePointList):

    idx_col = "Service Request Number"
    values = ["19914185"]

    results = sp_list.get_rows(idx_col, values)

    for result in results:
        print(result)

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
        }
    ]

    results = sp_list.add_rows(data_to_add)

    for result in results:
        print(result)

def main():
    sp_list = SharePointList.setup(site_name=SITE_NAME, list_name=LIST_NAME)
    test_get_rows(sp_list)
    test_add_rows(sp_list)

if __name__ == "__main__":
    main()
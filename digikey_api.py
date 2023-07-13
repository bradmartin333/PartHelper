import os
import digikey
from pathlib import Path
from digikey.v3.productinformation import KeywordSearchRequest
from digikey.v3.batchproductdetails import BatchProductDetailsRequest


def setup():
    # Read all lines from key file
    key_file = Path("digikey_api_key.txt")
    if key_file.is_file():
        with open(key_file, 'r') as f:
            lines = f.readlines()
            os.environ['DIGIKEY_CLIENT_ID'] = lines[0].strip()
            os.environ['DIGIKEY_CLIENT_SECRET'] = lines[1].strip()

    else:
        print("Error: digikey_api_key.txt not found")
        print("Create a file called digikey_api_key.txt with 2 lines")
        print("Line 1: Client ID")
        print("Line 2: Client Secret")
        exit(1)

    # Set environment variables
    os.environ['DIGIKEY_CLIENT_SANDBOX'] = 'True'
    os.environ['DIGIKEY_STORAGE_PATH'] = 'digikey_cache'


def get_product(manufacturer_pn):
    search_request = KeywordSearchRequest(
        keywords=ziuz.manufacturer_pn, record_count=1)
    result = digikey.keyword_search(body=search_request)
    return result.to_dict()['products'][0]

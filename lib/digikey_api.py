import os
import digikey
from pathlib import Path
from digikey.v3.productinformation import KeywordSearchRequest
from digikey.v3.batchproductdetails import BatchProductDetailsRequest


def setup():
    # Read all lines from key file
    key_file = Path(os.path.join('digikey_cache', 'digikey_api_key.txt'))
    if key_file.is_file():
        with open(key_file, 'r') as f:
            lines = f.readlines()
            os.environ['DIGIKEY_CLIENT_ID'] = lines[0].strip()
            os.environ['DIGIKEY_CLIENT_SECRET'] = lines[1].strip()

    else:
        print("Error: digikey_api_key.txt not found")
        print("Follow the instructions in digikey_cache/README.md")
        exit(1)

    # Set environment variables
    os.environ['DIGIKEY_CLIENT_SANDBOX'] = 'False'
    os.environ['DIGIKEY_STORAGE_PATH'] = 'digikey_cache'


def get_product(mfn):
    search_request = KeywordSearchRequest(keywords=mfn, record_count=1)
    result = digikey.keyword_search(body=search_request)
    return result.to_dict()


def get_cost_1(product):
    return '$' + str(product['unit_price'])


def get_cost_1000(product):
    full_str = str(product)
    break_quantity_index = full_str.find("'break_quantity': 1000,")
    if break_quantity_index == -1:
        return 'n/a'
    unit_price_index = full_str.find("'unit_price': ", break_quantity_index)
    end_index = full_str.find(",", unit_price_index)
    return '$' + str(float(full_str[unit_price_index + 14:end_index]))

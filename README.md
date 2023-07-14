# PartHelper

Scripts that help fill in PCBA parameters using the Digikey API and an offline .xlsx file.

This will really only be useful as a starting point for other projects since my use case is very specific and the .xlsx file is not included.

### Based on
- [peeter123's digikey_api](https://github.com/peeter123/digikey-api)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## Setup 

### Python and Filesystem

1. Install Python 3.10.9 or higher
1. Install pip
1. Install dependencies: `pip install -r requirements.txt`
1. Create `digikey_api_key.txt` in the `digikey_cache` directory
1. Copy your offline .xlsx into the root directory

### Digikey

1. Create a digikey account
1. Login as a [developer](https://developer.digikey.com/)
    - Note that a Sandbox app will not work! It must be a Production app. Even if you specify Sandbox in `digikey_api.py`, the results will be fixed to the demo data.
1. Click `Organizations` in the top right
1. Click `Create Organization` and follow the prompts
1. Click the newly created organization
1. Go to the `Production Apps` tab
1. Click `Create Production App` and follow the prompts
    - The callback URL must be `https://localhost:8139/digikey_callback`
    - Turn on the radio button for `Product Information` (The first one)
1. Click the newly created app
1. Show the `Client ID` and `Client Secret`
1. Copy them (Each on their own line and in that order) into `digikey_api_key.txt` and save

## Usage

1. Update the `this_sheet_name` in `alitum_helper.py` to match the name of the KiCad BOM sheet in your .xlsx file
1. Run `python alitum_helper.py`
1. If prompted to provide a new mfn, lookup the problematic mfn and provide a more suitable Digikey mfn
1. Note if a `Manual check required for` message is displayed at the end
1. Open the generated .xlsx and verify its contents
1. Copy and paste the contents into its future home

## Notes

- There is a daily rate limit of 1000 requests per day for the Digikey API

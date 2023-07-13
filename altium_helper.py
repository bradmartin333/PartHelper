import os
import openpyxl
import eda_parameters as eda
import digikey_api as dk

dk.setup()

for file in os.listdir(os.getcwd()):
    if file.endswith(".xlsx"):
        wb = openpyxl.load_workbook(file)
        print("Loaded file: " + file)
        break

output = wb["Output"]  # Copy vals from Altium into this sheet
combined = wb["Combined"]  # This is an aggregate of all internal part numbers
board = wb["TOF Module"]  # Update this to the board you're working on

alitum_params = eda.sheet_to_param_list(output, eda.AltiumParameter)
kicad_params = eda.sheet_to_param_list(board, eda.KiCadParameter)

for altium_param in alitum_params:
    skip = False
    if altium_param.notes == 'DNP':
        skip = True
    for ref in ['FID', 'LOGO', 'TP', 'MP', 'H', 'JP']:
        if ref in altium_param.identifier:
            skip = True
            continue
    if skip:
        continue
    found = False

    mfn = ''
    for kicad_param in kicad_params:
        if (altium_param.identifier + ',') in kicad_param.references:
            found = True
            mfn = kicad_param.manufacturer_pn
    if not found:
        for kicad_param in kicad_params:
            if altium_param.identifier in kicad_param.references:
                found = True
                mfn = kicad_param.manufacturer_pn
    if not found:
        print("Could not find " + altium_param.identifier)
        continue

    print(mfn)
    product = dk.get_product(mfn)
    print(str(product))
    exit(0)

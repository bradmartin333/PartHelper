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
board = wb["TOF Module"]  # Update this to the board you're working on

alitum_params = eda.sheet_to_param_list(output, eda.AltiumParameter)
kicad_params = eda.sheet_to_param_list(board, eda.KiCadParameter)

combined = wb["Combined"]  # This is an aggregate of all internal part numbers
combined_mfns = []
combined_internal_pns = []
combined_mfn_col = 3
for row in combined.iter_rows(min_row=2):
    combined_mfns.append(row[3].value)
    internal_1 = row[5].value
    internal_2 = row[6].value
    if internal_1 is None:
        combined_internal_pns.append(internal_2)
    else:
        combined_internal_pns.append(internal_1)

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
    original_mfn = mfn
    while product['exact_manufacturer_products_count'] == 0:
        print("Could not find " + mfn)
        mfn = input("Enter a new mfn: ")
        if len(mfn) == 0 or len(mfn) > 100:
            print("Invalid mfn")
            continue
        for param in kicad_params:
            if param.manufacturer_pn == original_mfn:
                param.manufacturer_pn = mfn
        combined_mfns[combined_mfns.index(original_mfn)] = mfn
        product = dk.get_product(mfn)
    exact_mfg_product = product['exact_manufacturer_products'][0]
    altium_param.cost_1 = dk.get_cost_1(exact_mfg_product)
    altium_param.cost_1000 = dk.get_cost_1000(product)
    if altium_param.current != '*':
        for param in exact_mfg_product['parameters']:
            if 'Current' in param['parameter']:
                altium_param.current = param['value']
    altium_param.distributor = 'Digikey'
    altium_param.distributor_pn = exact_mfg_product['digi_key_part_number']
    altium_param.help_url = exact_mfg_product['primary_datasheet']
    altium_param.manufacturer = exact_mfg_product['manufacturer']['value']
    altium_param.manufacturer_pn = exact_mfg_product['manufacturer_part_number']
    altium_param.parata_pn = combined_internal_pns[combined_mfns.index(mfn)]
    if altium_param.rohs != '*':
        altium_param.rohs = 'Yes' if 'ROHS3 Compliant' in str(product) else 'No'
    if altium_param.tolerance != '*':
        found = False
        for param in exact_mfg_product['parameters']:
            if 'Tolerance' in param['parameter']:
                found = True
                altium_param.tolerance = param['value']
        if not found:
            altium_param.tolerance = 'n/a'
    if altium_param.value != '*':
        altium_param.value = exact_mfg_product['parameters'][-1]['value']
    if altium_param.voltage != '*':
        for param in exact_mfg_product['parameters']:
            if 'Voltage' in param['parameter']:
                altium_param.voltage = param['value']
    if altium_param.wattage != '*':
        for param in exact_mfg_product['parameters']:
            if 'Power' in param['parameter']:
                altium_param.wattage = param['value']

eda.param_list_to_sheet(alitum_params, output)
wb.save("test.xlsx")

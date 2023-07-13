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

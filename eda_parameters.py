class AltiumParameter:
    def __init__(self, cols):
        self.object_type = cols[0]
        self.document = cols[1]
        self.identifier = cols[2]
        self.color = cols[3]
        self.cost_1 = cols[4]
        self.cost_1000 = cols[5]
        self.current = cols[6]
        self.distributor = cols[7]
        self.distributor_pn = cols[8]
        self.help_url = cols[9]
        self.manufacturer = cols[10]
        self.manufacturer_pn = cols[11]
        self.notes = cols[12]
        self.parata_pn = cols[13]
        # Everything is RoHS these days
        self.rohs = 'Yes' if cols[14] != '*' else '*'
        self.tolerance = cols[15]
        self.value = cols[16]
        self.voltage = cols[17]
        self.wattage = cols[18]

    def __str__(self):
        return f"Object Type: {self.object_type}\nDocument: {self.document}\nIdentifier: {self.identifier}\nColor: {self.color}\nCost 1: {self.cost_1}\nCost 1000: {self.cost_1000}\nCurrent: {self.current}\nDistributor: {self.distributor}\nDistributor PN: {self.distributor_pn}\nHelp URL: {self.help_url}\nManufacturer: {self.manufacturer}\nManufacturer PN: {self.manufacturer_pn}\nNotes: {self.notes}\nParata PN: {self.parata_pn}\nRoHS: {self.rohs}\nTolerance: {self.tolerance}\nValue: {self.value}\nVoltage: {self.voltage}\nWattage: {self.wattage}\n"


class KiCadParameter:
    def __init__(self, cols):
        self.item_number = cols[0]
        self.sourced = cols[1]
        self.placed = cols[2]
        self.references = cols[3]
        self.amps = cols[4]
        self.manufacturer_pn = cols[5]
        self.manufacturer = cols[6]
        self.tolerance = cols[7]
        self.type = cols[8]
        self.voltage = cols[9]
        self.value = cols[10]
        self.footprint = cols[11]
        self.quantity = cols[12]

    def __str__(self):
        return f"References: {self.references}\nAmps: {self.amps}\nManufacturer: {self.manufacturer}\nManufacturer PN: {self.manufacturer_pn}\nTolerance: {self.tolerance}\nType: {self.type}\nVoltage: {self.voltage}\nValue: {self.value}\nFootprint: {self.footprint}\n"


def sheet_to_param_list(sheet, param_class):
    param_list = []
    for row in sheet.iter_rows(min_row=2):
        cols = []
        for cell in row:
            if cell.value is None:
                cols.append("")
                continue
            cols.append(cell.value)
        param_list.append(param_class(cols))
    return param_list

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
        self.internal_pn = cols[13]
        self.rohs = cols[14]
        self.tolerance = cols[15]
        self.value = cols[16]
        self.voltage = cols[17]
        self.wattage = cols[18]

    def __str__(self):
        return f"Object Type: {self.object_type}\nDocument: {self.document}\nIdentifier: {self.identifier}\nColor: {self.color}\nCost 1: {self.cost_1}\nCost 1000: {self.cost_1000}\nCurrent: {self.current}\nDistributor: {self.distributor}\nDistributor PN: {self.distributor_pn}\nHelp URL: {self.help_url}\nManufacturer: {self.manufacturer}\nManufacturer PN: {self.manufacturer_pn}\nNotes: {self.notes}\nInternal PN: {self.internal_pn}\nRoHS: {self.rohs}\nTolerance: {self.tolerance}\nValue: {self.value}\nVoltage: {self.voltage}\nWattage: {self.wattage}\n"

    def fill_empty(self):
        '''Fill all fields except `notes` with `n/a`'''
        if self.color == '':
            self.color = 'n/a'
        if self.cost_1 == '':
            self.cost_1 = 'n/a'
        if self.cost_1000 == '':
            self.cost_1000 = 'n/a'
        if self.current == '':
            self.current = 'n/a'
        if self.distributor == '':
            self.distributor = 'n/a'
        if self.distributor_pn == '':
            self.distributor_pn = 'n/a'
        if self.help_url == '':
            self.help_url = 'n/a'
        if self.manufacturer == '':
            self.manufacturer = 'n/a'
        if self.manufacturer_pn == '':
            self.manufacturer_pn = 'n/a'
        if self.internal_pn == '':
            self.internal_pn = 'n/a'
        if self.rohs == '':
            self.rohs = 'n/a'
        if self.tolerance == '':
            self.tolerance = 'n/a'
        if self.value == '':
            self.value = 'n/a'
        if self.voltage == '':
            self.voltage = 'n/a'
        if self.wattage == '':
            self.wattage = 'n/a'


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


def sheet_to_param_list(sheet, param_class, has_header_row=False):
    '''Turns an excel sheet into a list of defined parameters'''
    param_list = []
    min_row = 2 if has_header_row else 1
    print("Reading sheet " + str(sheet) + " from row " + str(min_row))
    for row in sheet.iter_rows(min_row=min_row):
        cols = []
        for cell in row:
            if cell.value is None:
                cols.append("")
                continue
            cols.append(cell.value)
        param_list.append(param_class(cols))
    return param_list


def param_list_to_sheet(param_list, sheet, has_header_row=False):
    '''Turns a list of defined parameters into an excel sheet'''
    min_row = 2 if has_header_row else 1
    print("Writing sheet " + str(sheet) + " from row " + str(min_row))
    for i in range(len(param_list)):
        r = i + min_row
        sheet.cell(row=r, column=1).value = param_list[i].object_type
        sheet.cell(row=r, column=2).value = param_list[i].document
        sheet.cell(row=r, column=3).value = param_list[i].identifier
        sheet.cell(row=r, column=4).value = param_list[i].color
        sheet.cell(row=r, column=5).value = param_list[i].cost_1
        sheet.cell(row=r, column=6).value = param_list[i].cost_1000
        sheet.cell(row=r, column=7).value = param_list[i].current
        sheet.cell(row=r, column=8).value = param_list[i].distributor
        sheet.cell(row=r, column=9).value = param_list[i].distributor_pn
        sheet.cell(row=r, column=10).value = param_list[i].help_url
        sheet.cell(row=r, column=11).value = param_list[i].manufacturer
        sheet.cell(row=r, column=12).value = param_list[i].manufacturer_pn
        sheet.cell(row=r, column=13).value = param_list[i].notes
        sheet.cell(row=r, column=14).value = param_list[i].internal_pn
        sheet.cell(row=r, column=15).value = param_list[i].rohs
        sheet.cell(row=r, column=16).value = param_list[i].tolerance
        sheet.cell(row=r, column=17).value = param_list[i].value
        sheet.cell(row=r, column=18).value = param_list[i].voltage
        sheet.cell(row=r, column=19).value = param_list[i].wattage

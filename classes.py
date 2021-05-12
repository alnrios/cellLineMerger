import openpyxl


def col_scan_bool(wkst, desired_value, col_num):
    for col in wkst.iter_cols(col_num):
        for cell in col:
            if cell.value is not None and desired_value in str(cell.value):
                return True
    return False


def col_scan_value(wkst, desired_value, col_num):
    for col in wkst.iter_cols(col_num):
        for cell in col:
            if cell.value is not None and desired_value in str(cell.value):
                return cell


def inc_by_column(cell_name, quantity):
    return chr(ord(cell_name[0]) + quantity) + cell_name[1:]


def string_extract(cell, split_char):
    cell_string = cell.value
    if split_char in cell_string:
        return_val = cell_string.split(split_char)
        return "".join(return_val).strip()
    else:
        print(split_char + " not found in cell A1\n")


def clean_cell_name(cell):
    return str(cell).split('.')[1].split('>')[0]


def drug_check(wkst, col_number):
    for j in reversed(range(4)):
        if col_scan_bool(wkst, "Drug {}".format(str(j + 1)), col_number):
            return j + 1


class ExcelWkbk:
    def __init__(self, wkbk):
        self.imported_wkbk = wkbk
        self.sheet_names = wkbk.sheetnames
        self.sheets = []
        self.technician = ""

    def get_worksheets(self):
        for sheet in range(len(self.sheet_names)):
            ws = self.imported_wkbk[self.sheet_names[sheet]]
            self.technician = string_extract(ws['A1'], 'Name:')
            self.sheets.append(
                ExcelSheet(
                    date=string_extract(ws['A2'], 'SET UP DATE:'),
                    num_of_drugs=drug_check(ws, 1)
                )
            )
            num_drugs = self.sheets[sheet].num_of_drugs
            self.get_and_set_drugs(num_drugs, ws, sheet)

    def get_and_set_drugs(self, num_drugs, wkst, sheet_index):
        for i in range(num_drugs):
            if col_scan_bool(wkst, 'Drug {}'.format(i + 1), 1):
                cell_val = col_scan_value(wkst, 'Drug {}'.format(i + 1), 1)
                drug_name = cell_val.value
                cell_val = clean_cell_name(cell_val)
                conc_cell = inc_by_column(cell_val, 1)
                dilut_cell = inc_by_column(cell_val, 2)
                for col in wkst.iter_cols(2, 3):
                    for cell in col:
                        if conc_cell in str(cell):
                            conc_cell = cell.value
                            break
                        elif dilut_cell in str(cell):
                            dilut_cell = cell.value
                            break
                self.sheets[sheet_index].append_to_drug_list(Drug(drug_name, conc_cell, dilut_cell))

    def get_num_of_sheets(self):
        return len(self.sheets)

    def display_sheet_info(self):
        print(str(self.get_num_of_sheets()) + " sheets in " + self.technician + "'s workbook")
        print('*' * 35)
        for i in range(len(self.sheets)):
            self.sheets[i].display_info()


class ExcelSheet:
    def __init__(self, date, num_of_drugs):
        self.date = date
        self.num_of_drugs = num_of_drugs
        self.drug_list = []
        self.day1_list = []
        self.day7_list = []

    def display_info(self):
        print(str(self.num_of_drugs) + " drugs used on " + self.date)
        for i in range(self.num_of_drugs):
            drug_info = self.drug_list[i]
            print(drug_info.name + "\n\tConcentration:" +
                  str(drug_info.concentration) + "\n\tDilution: " + str(drug_info.dilution))
        print('-' * 35)

    def append_to_drug_list(self, drug):
        self.drug_list.append(drug)


class Plate:
    def __init__(self, barcode):
        self.barcode = barcode
        self.num_of_cell_lines = 0
        self.cell_line_list = []

    def append_cell_lines(self, cell_line):
        self.cell_line_list.append(cell_line)


class ControlPlate(Plate):
    def __init__(self, control_type):
        super().__init__("ControlPlate")
        self.control_type = control_type


class TreatmentPlate(Plate):
    def __init__(self):
        super().__init__("TreatmentPlate")
        self.day1_link = ""
        self.day7_link = ""


class Drug:
    def __init__(self, name, concentration, dilution):
        self.name = name
        self.concentration = concentration
        self.dilution = dilution

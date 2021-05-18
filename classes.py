import openpyxl


def cell_scan(wkst, desired_value, col_num, type_of_scan):
    for col in wkst.iter_cols(col_num):
        for cell in col:
            if cell.value is not None:
                if type_of_scan == 'value return bool':
                    if desired_value in str(cell.value):
                        return True
                elif type_of_scan == 'value return value':
                    if desired_value in str(cell.value):
                        return cell
                elif type_of_scan == 'cell return cell':
                    if desired_value in str(cell):
                        return cell
                elif type_of_scan == 'cell return bool':
                    if desired_value in str(cell):
                        return True
                else:
                    print("Incorrect type of scan. Exit and try again...")


def control_scan(wkst, desired_value, col_num):
    control_list = []
    for col in wkst.iter_cols(col_num):
        for cell in col:
            if cell.value is not None:
                if isinstance(desired_value, list):
                    for i in range(len(desired_value)):
                        if desired_value[i] in str(cell.value):
                            control_list.append(clean_cell_name(cell))
                else:
                    if desired_value in str(cell.value):
                        control_list.append(clean_cell_name(cell))
    return control_list


def inc_by_column(cell_name, quantity):
    return chr(ord(cell_name[0]) + quantity) + cell_name[1:]


# def insert_and_pop(my_list, )

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
    for j in reversed(range(10)):
        if cell_scan(wkst, "Drug {}".format(str(j + 1)), col_number, "value return bool"):
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
            self.get_controls(ws, sheet)

    def get_and_set_drugs(self, num_drugs, wkst, sheet_index):
        for i in range(num_drugs):
            if cell_scan(wkst, 'Drug {}'.format(i + 1), 1, 'value return bool'):
                cell_val = cell_scan(wkst, 'Drug {}'.format(i + 1), 1, 'value return value')
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

    def get_controls(self, wkst, sheet_index):
        control_xl_dict = {
            'd1_xl_cells': control_scan(wkst, ['Day1', 'Day 1'], 1),
            'd7_xl_cells': control_scan(wkst, ['Day7', 'Day 7'], 1)
        }
        for i in control_xl_dict:
            for xl_iter in range(len(control_xl_dict[i])):  # iterator for the control_xl_dict items in lists
                barcode = inc_by_column(control_xl_dict[i][xl_iter], 1)   # first assigns barcode var to the cell number
                col_count = 2  # col_count refers to the column number that'll be used, which will always be 2 for bc's
                cell_line_list = self.get_cell_lines(wkst, barcode, col_count)
                barcode = cell_scan(wkst, barcode, col_count, 'cell return cell').value
                if i == 'd1_xl_cells':
                    self.sheets[sheet_index].day1_list.append(
                        ControlPlate(barcode, len(cell_line_list), cell_line_list, 'd1'))
                else:
                    self.sheets[sheet_index].day1_list.append(
                        ControlPlate(barcode, len(cell_line_list), cell_line_list, 'd7'))

#######################################################################################################################

    def get_cell_lines(self, wkst, barcode_cell, count):
        cell_line_list = []
        pos = 0     # records the position of the cell line
        while True:
            pos += 1
            barcode_cell = inc_by_column(barcode_cell, 1)
            count += 1
            cell_line = cell_scan(wkst, barcode_cell, count, 'cell return cell')
            if cell_line is not None:
                cell_line_list.append(CellLine(cell_line.value, pos))
            else:
                break
        return cell_line_list






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
        print("Control Plates:")
        for i in range(len(self.day1_list)):
            print("\t" + str(self.day1_list[i]))

    def append_to_drug_list(self, drug):
        self.drug_list.append(drug)

    # def append_to_control_list(self, control_type):
    #     if control_type == 'd1':



class Plate:
    def __init__(self, barcode, num_of_cell_lines, cell_line_list):
        self.barcode = barcode
        self.num_of_cell_lines = num_of_cell_lines
        self.cell_line_list = cell_line_list

    # def append_cell_lines(self, cell_line):
    #     self.cell_line_list.append(cell_line)


class ControlPlate(Plate):
    def __init__(self, barcode, num_of_cell_lines, cell_line_list, control_type):
        super().__init__(barcode, num_of_cell_lines, cell_line_list)
        self.control_type = control_type

    def __str__(self):
        return str(self.num_of_cell_lines) + " cell lines in " + self.control_type + " plate " + \
               self.barcode + ": " + self.display_cell_lines()

    def display_cell_lines(self):
        return_string = "["
        for i in range(self.num_of_cell_lines):
            if i == self.num_of_cell_lines - 1:
                return_string += (self.cell_line_list[i].name + ": " + str(self.cell_line_list[i].position))
            else:
                return_string += (self.cell_line_list[i].name + ": " + str(self.cell_line_list[i].position) + ", ")
        return_string += "]"
        return return_string


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


class CellLine:
    def __init__(self, name, position):
        self.name = name
        self.position = position

    def __str__(self):
        return self.name

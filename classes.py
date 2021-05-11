import openpyxl


def col_scan(wkst, desired_value, col_num):
    for col in wkst.iter_cols(col_num):
        for cell in col:
            if cell.value is not None and desired_value in str(cell.value):
                return True
    return False


def string_extract(cell, string):
    cell_string = cell.value
    if string in cell_string:
        return_val = cell_string.split(string)
        return "".join(return_val).strip()
    else:
        print(string + " not found in cell A1\n")


def drug_check(wkst, col_number):
    for j in reversed(range(4)):
        if col_scan(wkst, "Drug {}".format(str(j + 1)), col_number):
            return j + 1


class ExcelWkbk:
    def __init__(self, wkbk):
        self.wkbk = wkbk
        self.sheet_names = wkbk.sheetnames
        self.sheets = []
        self.technician = ""

    def get_worksheets(self):
        for sheet in range(len(self.sheet_names)):
            ws = self.wkbk[self.sheet_names[sheet]]
            self.technician = string_extract(ws['A1'], 'Name:')
            self.sheets.append(
                ExcelSheet(
                    date=string_extract(ws['A2'], 'SET UP DATE:'),
                    num_of_drugs=drug_check(ws, 1)
                )
            )

    def get_num_of_sheets(self):
        return len(self.sheets)

    def print_sheet_info(self):
        print(str(self.get_num_of_sheets()) + " sheets in " + self.technician + "'s workbook")
        for i in range(len(self.sheets)):
            self.sheets[i].display_info()


class ExcelSheet:
    def __init__(self, date, num_of_drugs):
        self.date = date
        self.num_of_drugs = num_of_drugs

    def display_info(self):
        print(str(self.num_of_drugs) + " drugs used on " + self.date + "\n")


class ControlPlate:
    def __init__(self, num_of_cell_lines, cell_line_names):
        if num_of_cell_lines != len(cell_line_names):
            print("The amount of cell lines listed "
                  "does not match the amount given\n"
                  "Try again\n")
        else:
            self.num_of_cell_lines = num_of_cell_lines
            self.cell_line_names = cell_line_names
            print("Success!\n")

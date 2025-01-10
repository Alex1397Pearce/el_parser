import os
import pandas as pd
import openpyxl


class Files:
    def __init__(self, filepath):
        self.filepath = filepath
        self.status = self.file_exist()

    def file_exist(self):
        if os.path.exists(self.filepath):
            return True
        else:
            return False


class Reader(Files):
    def get_list(self, column_name="Артикул"):
        # check file exist
        try:
            df = pd.read_excel(self.filepath)
            column_values = df[column_name].tolist()
            return column_values
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return []


class Excel(Files):
    def __init__(self, filepath):
        super().__init__(filepath)
        if not self.status:
            workbook = openpyxl.Workbook()
            workbook.save(self.filepath)

    def clean(self):
        if not self.status:
            pass
        os.remove(self.filepath)

    def list_to_excel(self, value_articul, url=""):
        workbook = openpyxl.load_workbook(self.filepath)
        while len(workbook.sheetnames) > 1:
            workbook.remove(workbook[workbook.sheetnames[1]])

        if not workbook.sheetnames:
            workbook.create_sheet()

        sheet = workbook.active
        first_empty_row = 1
        while sheet.cell(row=first_empty_row, column=1).value is not None:
            first_empty_row += 1

        sheet.cell(row=first_empty_row, column=1, value=value_articul)
        sheet.cell(row=first_empty_row, column=2, value=url)
        workbook.save(self.filepath)

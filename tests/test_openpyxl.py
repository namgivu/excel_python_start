import os
from shutil import copyfile

from openpyxl import load_workbook, Workbook

PWD = os.path.abspath(os.path.dirname(__file__))

INPUT_FILE  = f'{PWD}/fixture/input.xlsx'
OUTPUT_FILE = f'{PWD}/tmp/output.xlsx'

#region read+write xlsx util
"""
EXP aka  EXPECTED
ACT aka  ACTUAL

wb  aka  workbook
ws  aka  worksheet
"""


class ER:  # ER aka ExcelRead

    @classmethod
    def openpyxl_read_excel(self, path_file):
        wb = load_workbook(path_file)
        ws = wb.active

        max_row = ws.max_row
        max_column = ws.max_column

        data_return = ()
        for i in range(1, max_row + 1):  # iterate over all cells
            row = []
            for j in range(1, max_column + 1):
                cell_obj = ws.cell(row=i, column=j)
                row.append(cell_obj.value)
            data_return += (row,)
        return data_return


class EW:  # EW aka ExcelWrite

    @staticmethod
    def openpyxl_write_excel_file(file_name, data_to_write=()):
        wb = Workbook()
        ws = wb.active

        # append all rows
        for row in data_to_write:
            ws.append(tuple(row))

        # save file
        wb.save(file_name)


    @staticmethod
    def openpyxl_update_excel_file(file_name, data_to_write=()):
        wb = load_workbook(file_name)
        ws = wb.active
        max_row = ws.max_row
        for row in data_to_write:
            for i in range(1, len(row) + 1):  ## first start column and row are 1
                cell = ws.cell(row=max_row + 1, column=i)
                cell.value = row[i - 1]
            max_row += 1 ## add 1 to write new row
        wb.save(file_name)

#endregion read+write xlsx util


class Test:

    def test_openpyxl_read(self):
        EXP = (['abb', 122], ['xxx', 333])

        # testee code
        ACT = ER.openpyxl_read_excel(INPUT_FILE)

        assert ACT == EXP


    def test_openpyxl_write_new_file(self):
        EXP = (['abb', 122], ['xxx', 333])

        # testee code
        EW.openpyxl_write_excel_file(OUTPUT_FILE, EXP)

        # check by reread
        reread = ER.openpyxl_read_excel(OUTPUT_FILE)
        assert reread == EXP


    def test_openpyxl_update_current_file(self):
        # create fixture
        current_rows = ER.openpyxl_read_excel(INPUT_FILE)
        new_rows     = (['new1', 11], ['new2', 22])
        EXP = current_rows + new_rows

        # testee code - create new file from :INPUT and insert :new_rows to :OUTPUT
        copyfile(INPUT_FILE, OUTPUT_FILE)
        EW.openpyxl_update_excel_file(OUTPUT_FILE, new_rows)

        ACT = ER.openpyxl_read_excel(OUTPUT_FILE)
        assert ACT == EXP

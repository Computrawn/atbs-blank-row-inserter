#! python3
# blankRowInserter.py — An exercise in manipulating Excel files.

import sys
import openpyxl

row_location = int(sys.argv[1])
number_of_rows = int(sys.argv[2])
file_name = sys.argv[3]


def insert_rows(location, number, name):
    """Inserts blank rows into excel file and saves it as a new file."""
    wb = openpyxl.load_workbook(f"{name}.xlsx")
    sheet = wb.active
    sheet.insert_rows(location, number)
    wb.save(f"{name}_plus_{number}_rows.xlsx")


insert_rows(row_location, number_of_rows, file_name)

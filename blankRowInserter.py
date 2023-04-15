#! python3
# blankRowInserter.py — An exercise in manipulating Excel files.

import openpyxl

# TODO: Change these variables to command line arguments.

row_location = int(input("Please enter row number: "))
number_of_rows = int(input("Please enter number of rows to add: "))
file_name = input("Please enter file name: ")


def insert_rows(location, number, name):
    """Inserts blank rows into excel file and saves it as a new file."""
    wb = openpyxl.load_workbook(f"{name}.xlsx")
    sheet = wb.active
    sheet.insert_rows(location, number)
    wb.save(f"{name}_plus_{number}_rows.xlsx")


insert_rows(row_location, number_of_rows, file_name)

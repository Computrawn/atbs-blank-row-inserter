#!/usr/bin/env python3
# blankRowInserter.py — An exercise in manipulating Excel files.
# For more information, see README.md

import logging
import sys
import openpyxl

logging.basicConfig(
    level=logging.DEBUG,
    filename="logging.txt",
    format="%(asctime)s -  %(levelname)s -  %(message)s",
)
logging.disable(logging.CRITICAL)  # Note out to enable logging.


def main():
    """Call insert_rows function on command line arguments."""
    try:
        insert_rows(int(sys.argv[1]), int(sys.argv[2]), sys.argv[3])
    except IndexError:
        print(
            """Please run script from the command line with the following arguments:
* row location
* number of rows
* file name"""
        )


def insert_rows(location, number, name):
    """Inserts blank rows into excel file and saves it as a new file."""
    wb = openpyxl.load_workbook(f"{name}.xlsx")
    sheet = wb.active
    sheet.insert_rows(location, number)
    wb.save(f"{name}_plus_{number}_rows.xlsx")


if __name__ == "__main__":
    main()

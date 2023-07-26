import os
import csv
import argparse
from xlsxwriter.workbook import Workbook



if __name__ == "__main__":
    # Don't run if module being imported. Only if script being run directly.
    parser = argparse.ArgumentParser(description="Program to automatically "
                            "convert CPF exports from fake .XLS to real .XLSX")
    parser.add_argument("-d", "--dir", help="Specify dir containing exports "
                                        "to reformat", type=str, default=".")
    # https://www.programcreek.com/python/example/748/argparse.ArgumentParser
    args = parser.parse_args()

    dir_path = os.path.abspath(os.path.normpath(args.dir))

    for tsv_item in sorted(os.listdir(dir_path)):
        xls_item_path = os.path.join(dir_path, tsv_item)
        if os.path.isdir(xls_item_path):
            # print("%s not a file." % item)
            continue
        if os.path.splitext(tsv_item)[-1].upper() != ".XLS":
            continue

        # Rename .XLSX to .tsv
        # print("\nRenaming %s" % tsv_item, end="")
        # tsv_item_path = os.path.splitext(xls_item_path)[0] + ".tsv"
        # os.rename(xls_item_path, tsv_item_path)
        tsv_item_path = xls_item_path
        # print("...done")

        # Generate a .xlsx to populate w/ .tsv data.
        xlsx_file = os.path.splitext(tsv_item_path)[0] + ".xlsx"

        print("Creating %s" % os.path.basename(xlsx_file), end="")
        with Workbook(xlsx_file) as workbook:
            worksheet = workbook.add_worksheet()

            # Create a TSV file reader.
            tsv_reader = csv.reader(open(tsv_item_path, 'r'), delimiter='\t')

            # Read the row data from the TSV file and write it to the XLSX file.
            for row, data in enumerate(tsv_reader):
                worksheet.write_row(row, 0, data)

            # Borrowed from here
            # https://stackoverflow.com/questions/16852655/convert-a-tsv-file-to-xls-xlsx-using-python
        print("...done")


# Created .exe from this after testing script in cmd.exe
# Installed PyInstaller with this in cmd.exe
# python -m pip install pyinstaller
# Made .exe with this command:
# pyinstaller fix_cpf_export_format.py
# Then made single file w/ this (not sure if it needs to be run a second time or if using --):
# python -m pip install pyinstaller
# pyinstaller fix_cpf_export_format.py --onefile
# fix_cpf_export_format.py is/was this script's name 
# https://www.blog.pythonlibrary.org/2021/05/27/pyinstaller-how-to-turn-your-python-code-into-an-exe-on-windows/
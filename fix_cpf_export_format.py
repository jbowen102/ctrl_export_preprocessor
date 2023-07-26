import os
import csv
import argparse

from xlsxwriter.workbook import Workbook


def wait_for_input():
    input("\nEnd of Script. Press Enter to finish and close.")



if __name__ == "__main__":
    try:
        # Don't run if module being imported. Only if script being run directly.
        parser = argparse.ArgumentParser(description="Program to automatically "
                                "collect CPF exports into one .xlsx file.")
        parser.add_argument("-d", "--dir", help="Specify dir containing exports "
                                            "to reformat", type=str, default=".")
        # https://www.programcreek.com/python/example/748/argparse.ArgumentParser
        args = parser.parse_args()

        dir_path = os.path.abspath(os.path.normpath(args.dir))

        # Generate a .xlsx to populate w/ the TSV data.
        xlsx_file = os.path.join(dir_path, "CPF_exports.xlsx")
        with Workbook(xlsx_file) as workbook:
            print("Creating %s" % os.path.basename(xlsx_file))

            for tsv_item in sorted(os.listdir(dir_path)):
                tsv_item_path = os.path.join(dir_path, tsv_item)
                if os.path.isdir(tsv_item_path):
                    continue
                if os.path.splitext(tsv_item)[-1].upper() != ".XLS":
                    continue

                print("\tReading from %s..." % tsv_item, end="")
                # Make a new tab in the output worksheet w/ the same name as the XLS/TSV file.
                worksheet = workbook.add_worksheet(os.path.splitext(tsv_item)[0])

                # Read the row data from the TSV file and write it to the XLSX file.
                tsv_reader = csv.reader(open(tsv_item_path, 'r'), delimiter='\t')
                for row, data in enumerate(tsv_reader):
                    worksheet.write_row(row, 0, data)

                # Original TSV reader code borrowed from here
                # https://stackoverflow.com/questions/16852655/convert-a-tsv-file-to-xls-xlsx-using-python
                print("done")
            print("...done")
        wait_for_input()

    except Exception as exception_text:
        print("\n")
        print(exception_text)
        print("\n" + "*"*10 + "\nException encountered\n" + "*"*10)
        wait_for_input()

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
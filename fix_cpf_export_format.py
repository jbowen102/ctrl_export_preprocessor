import os
import csv
import xlrd
import argparse
from datetime import datetime

import magic
from xlsxwriter.workbook import Workbook


def wait_for_input():
    input("\nEnd of Script. Press Enter to finish and close.")



if __name__ == "__main__":
    # Don't run if module being imported. Only if script being run directly.
    try:
        parser = argparse.ArgumentParser(description="Program to automatically "
                                "collect CPF exports into one .xlsx file.")
        parser.add_argument("-d", "--dir", help="Specify dir containing exports "
                                            "to reformat", type=str, default=".")
        # https://www.programcreek.com/python/example/748/argparse.ArgumentParser
        args = parser.parse_args()

        dir_path = os.path.abspath(os.path.normpath(args.dir))

        # Generate a .xlsx to populate w/ the TSV data.
        timestamp = datetime.now().strftime("%Y-%m-%dT%H%M%S")
        xlsx_file = os.path.join(dir_path, "CPF_exports_%s.xlsx" % timestamp)
        with Workbook(xlsx_file) as workbook:
            print("Creating %s" % os.path.basename(xlsx_file))

            # Loop through all CPF exports in directory.
            for cpf_export in sorted(os.listdir(dir_path)):
                cpf_export_path = os.path.join(dir_path, cpf_export)
                if os.path.isdir(cpf_export_path):
                    continue
                if os.path.splitext(cpf_export)[-1].upper() != ".XLS":
                    continue

                # Make a new tab in the output worksheet w/ the same name as the CPF export.
                worksheet = workbook.add_worksheet(os.path.splitext(cpf_export)[0])
                # Read the row data from the CPF export and write it to the XLSX file.

                # Determine if CPF export is a real XLS or TSV.
                # https://stackoverflow.com/questions/43580/how-to-find-the-mime-type-of-a-file-in-python
                # mime_type = mimetypes.guess_type(cpf_export_path)[0] # This only uses extension to determine.
                MagicObj = magic.detect_from_filename(cpf_export_path)
                # Not based on extension, despite function name seeming to indicate that.

                print("\tReading from %s..." % cpf_export, end="")
                if MagicObj.mime_type == "text/plain":
                    # TSV masquerading as XLS
                    tsv_reader = csv.reader(open(cpf_export_path, 'r'), delimiter='\t')
                    for row, data in enumerate(tsv_reader):
                        worksheet.write_row(row, 0, data)

                elif MagicObj.mime_type == "application/vnd.ms-excel":
                    # Real XLS
                    with xlrd.open_workbook(cpf_export_path) as xls_reader:
                        xls_sheet = xls_reader.sheet_by_index(0)
                        for row in range(xls_sheet.nrows):
                            data = xls_sheet.row_values(row)
                            worksheet.write_row(row, 0, data)
                else:
                    raise Exception('"%s" - Filetype not recognized (should be '
                                    'CPF export w/ .XLS extension)' % cpf_export)
                # Reference: If wrong file is passed to wrong reader, returns UnicodeDecodeError w/ TSV-read attempt.
                # xlrd.biffh.XLRDError w/ XLS-read attempt (that encounters a TSV)

                print("done")
            print("...done")
        wait_for_input()

    except Exception as exception_text:
        print("\n")
        print(exception_text)
        print("\n" + "*"*10 + "\nException encountered\n" + "*"*10)
        wait_for_input()


# Created .exe from this after testing script in cmd.exe
# Installed PyInstaller with this in cmd.exe:
    # python -m pip install pyinstaller
# Made .exe with this command:
    # pyinstaller fix_cpf_export_format.py --onefile
# https://www.blog.pythonlibrary.org/2021/05/27/pyinstaller-how-to-turn-your-python-code-into-an-exe-on-windows/
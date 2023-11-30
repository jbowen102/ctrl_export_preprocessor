import os
import csv
import xlrd
import argparse
from datetime import datetime

from tqdm import tqdm
import magic
from xlsxwriter.workbook import Workbook


def wait_for_input():
    input("\nEnd of Script. Press Enter to finish and close.")


def convert_export(tsv_path, new_filename, check_for_xls=True, replace=True):
    if not os.path.exists(tsv_path):
        print('Can\'t find source file "%s".' % os.path.basename(tsv_path))
        return
    if os.path.splitext(tsv_path)[-1].upper() != ".XLS":
        raise Exception('"%s" - Filetype not recognized (should be '
                        'CPF export w/ .XLS extension)' % os.path.basename(tsv_path))

    new_filepath = os.path.join(os.path.dirname(tsv_path), new_filename)
    if os.path.exists(new_filepath):
        # Don't overwrite
        print('"%s" already exists. Skipping.' % new_filename)
        return new_filepath

    tsv_mime_type = "text/plain"
    xls_mime_type = "application/vnd.ms-excel"
    xlsx_mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" # Reference
    if check_for_xls:
        # Determine if CPF export is a real XLS or TSV.
        # https://stackoverflow.com/questions/43580/how-to-find-the-mime-type-of-a-file-in-python
        MagicObj = magic.detect_from_filename(tsv_path)
        # Not based on extension, despite function name seeming to indicate that.

        mime_type = MagicObj.mime_type
        if mime_type not in [tsv_mime_type, xls_mime_type]:
            raise Exception('"%s" - Filetype not recognized (should be '
                            'CPF export w/ .XLS extension)' % os.path.basename(tsv_path))
    else:
        mime_type = tsv_mime_type
        # Used in cases where this function gets called right after exporting
        # from program, so we can be sure it's the raw export.

    with Workbook(new_filepath) as workbook:
        worksheet = workbook.add_worksheet("Parameters")

        if mime_type == tsv_mime_type:
            # TSV masquerading as XLS
            tsv_reader = csv.reader(open(tsv_path, 'r'), delimiter='\t')
            for row, data in enumerate(tsv_reader):
                worksheet.write_row(row, 0, data)
            # Borrowed from here
            # https://stackoverflow.com/questions/16852655/convert-a-tsv-file-to-xls-xlsx-using-python

        elif mime_type == xls_mime_type:
            # Real XLS
            xls_path = tsv_path
            with xlrd.open_workbook(xls_path) as xls_reader:
                xls_sheet = xls_reader.sheet_by_index(0)
                for row in range(xls_sheet.nrows):
                    data = xls_sheet.row_values(row)
                    worksheet.write_row(row, 0, data)

    if os.path.exists(new_filepath): # Confirm
        if replace:
            os.remove(tsv_path) # Delete tsv file

        return new_filepath
    else:
        raise Exception("Can't find export %s converted from %s" % (new_filepath, tsv_path))


def convert_all_exports(dir_path):
    for tsv_item in tqdm(sorted(os.listdir(dir_path)), colour="yellow"):
        tsv_item_path = os.path.join(dir_path, tsv_item)
        if os.path.isdir(tsv_item_path):
            # print("%s not a file." % item)
            continue
        if os.path.splitext(tsv_item)[-1].upper() != ".XLS":
            continue

        new_xlsx_filename = os.path.splitext(tsv_item_path)[0] + ".xlsx"
        convert_export(tsv_item_path, new_xlsx_filename)


def convert_and_aggregate_exports(dir_path):
    """Program to automatically collect CPF exports into one .xlsx file."""
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
                # XLSX is this: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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



if __name__ == "__main__":
    # Don't run if module being imported. Only if script being run directly.
    try:
        parser = argparse.ArgumentParser(description="Program to fix CPF-export"
                                                            "file format.")
        parser.add_argument("-d", "--dir", help="Specify dir containing exports "
                                            "to reformat", type=str, default=".")
        # https://www.programcreek.com/python/example/748/argparse.ArgumentParser
        args = parser.parse_args()

        dirpath = os.path.abspath(os.path.normpath(args.dir))
        convert_and_aggregate_exports(dirpath)
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
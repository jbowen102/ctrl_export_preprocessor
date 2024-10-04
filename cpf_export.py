import os
import time
import re
import subprocess
import shutil

import argparse
import pandas as pd
from tqdm import tqdm
import colorama
from xlsxwriter.workbook import Workbook
import xlwings as xw
import openpyxl as pyxl
if os.name == "nt":
    # Allows me to test other (non-GUI) features in WSL where pyautogui import fails
    import pyautogui as gui
    gui.FAILSAFE = True
    # Allows moving mouse to upper-left corner of screen to abort execution.
    gui.PAUSE = 0.5 # 500 ms pause after each command.
    # https://pyautogui.readthedocs.io/en/latest/quickstart.html


import fix_cpf_export_format as fixcpf
from dir_names import DIR_REMOTE_SRC, \
                      DIR_FIELD_DATA, \
                        DIR_IMPORT_ROOT, DIR_REMOTE_BU, DIR_IMPORT, \
                        DIR_EXPORT, DIR_EXPORT_BUFFER, \
                      DIR_REMOTE_SHARE_CTRL, DIR_REMOTE_SHARE_BATT, \
                      AZ_BLOB_ADDR_CTRL, AZ_BLOB_ADDR_BATT, \
                      ERROR_HISTORY_SAVE_IMG, ERROR_HISTORY_BLANK


DATE_FORMAT = "%Y%m%d"
SN_REGEX = r"(3\d{6}|5\d{6}|8\d{6})"
# Any "3" or "5" or "8" followed by six more digits
DATE_REGEX = r"(20\d{2}[0-1]\d[0-3]\d)"
# Any "20" followed by two digits,
    # followed by either "0" or "1" and any digit (months)
        # followed by either "0", "1", "2", or "3" paired with any digit (days 00-39)
# Could catch some invalid dates like 20231131. Further validated below in find_in_string()


CDF_EXPORT_SUFFIX = "_CDF.xlsx"
CPF_PARAM_EXPORT_SUFFIX = "_cpf-params.tsv"
CPF_FAULT_EXPORT_SUFFIX = "_cpf-faults.tsv"
CPF_COMBINED_EXPORT_SUFFIX = "_cpf.xlsx"

ERROR_HISTORY_SAVE_BUTTON_LOC = None # Will be modified below


class UserCancel(Exception):
    pass


def find_in_string(regex_pattern, string_to_search, prompt, date_target=False, allow_none=False):
    """Finds single match in string_to_search or presents prompt to user.
    If allow_none set to True, prompt only given upon multiple matches.
    date_target=True adds date validation.
    """
    found = None # Initialize variable for loop
    while not found:
        matches = re.findall(regex_pattern, string_to_search, flags=re.IGNORECASE)
        if len(matches) == 1 and date_target:
            # If looking for a date, check for valid date value (regex doesn't fully validate)
            # print("\t\tmatches[0]: " + matches[0]) # DEBUG
            try:
                time.strptime(matches[0], DATE_FORMAT)
            except ValueError:
                # Fall through to prompt user for manual entry.
                pass
            else:
                # Valid date
                return matches[0]
        elif len(matches) == 1:
            return matches[0]
            # loop exits
        elif len(matches) == 0 and allow_none:
            # print("\t\t%s: no matches; returning None" % string_to_search) # DEBUG
            return None

        # No matches, multiple matches, or invalid date found:
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + prompt)
        string_to_search = input(">" + colorama.Style.RESET_ALL)


def datestamp_remote(remote=DIR_REMOTE_SRC):
    while not os.path.exists(remote):
        # Prompt user to mount network drives if not found.
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\n"%s" not found. Mount '
                                'and press Enter to try again.' % src)
        input("> " + colorama.Style.RESET_ALL)

    # Keep track of renames for display later.
    old_names = []
    new_names = []

    file_count = sum(len(files) for _, _, files in os.walk(remote))
    # https://stackoverflow.com/questions/35969433/using-tqdm-on-a-for-loop-inside-a-function-to-check-progress
    with tqdm(total=file_count, colour="#05e4ab") as pbar:
        for dirpath, dirnames, filenames in os.walk(remote):
            for file_name in sorted(filenames):
                pbar.update(1)
                filepath = os.path.join(dirpath, file_name)
                item_name = os.path.splitext(file_name)[0]
                ext = os.path.splitext(file_name)[-1]

                if ext.lower() in (".cpf", ".cdf"):
                    # Find S/N in filename
                    prompt_str = "Can't parse S/N from import filename \"%s\".\n" \
                                                    "Type S/N manually: " % file_name
                    # print("\n\tS/N:") # DEBUG
                    serial_num = find_in_string(SN_REGEX, item_name, prompt_str)
                    # Now look for date in remaining string. Will add later if not present.
                    # print("\tReceived %s as S/N back from find_in_string()" % serial_num) # DEBUG
                    remaining_str = item_name.split(serial_num)
                    date_found = False
                    for substring in remaining_str:
                        prompt_str = "Can't find single valid date match in import " \
                                                            "filename \"%s\".\n" \
                                "Type manually (YYYYMMDD format): " % file_name
                        # print("\tDate:") # DEBUG
                        date_match = find_in_string(DATE_REGEX, substring,
                                    prompt_str, date_target=True, allow_none=True)
                        # print("\tReceived %s as date back from find_in_string()" % date_match) # DEBUG

                        if date_match is not None:
                            date_found = True
                            existing_datestamp = date_match
                            break

                    if date_found:
                        datestamp = existing_datestamp
                    else:
                        # Find file last-modified time. Precise enough for our needs.
                        mod_date = time.localtime(os.path.getmtime(filepath))

                        # Some files (CDF at least) have bogus mod dates - usually in 1999 or 2000.
                        # In that case, use today's date.
                        if mod_date < time.strptime("20200101", DATE_FORMAT):
                            # Substitute in today's date
                            date_to_use = time.localtime()
                        else:
                            date_to_use = mod_date

                        datestamp = time.strftime(DATE_FORMAT, date_to_use)

                    new_filename = "%s_sn%s%s" % (datestamp, serial_num, ext)
                    new_filepath = os.path.join(dirpath, new_filename)

                    if file_name != new_filename:
                        os.rename(filepath, new_filepath)

                        old_names.append(file_name)
                        new_names.append(new_filename)
                # input("> ") # DEBUG

    print("Renames:")
    if len(old_names) > 0:
        for i, name in enumerate(old_names):
            print(colorama.Fore.MAGENTA + "\t%s\t->\t%s" % (old_names[i], new_names[i]))
        input(colorama.Fore.GREEN + colorama.Style.BRIGHT + "\nPress Enter to continue"
                                                    + colorama.Style.RESET_ALL)
    else:
        print(colorama.Fore.MAGENTA + "\t[None]" + colorama.Style.RESET_ALL)
        time.sleep(2) # Pause for user to see that no files were renamed.


def sync_remote(src, dest, multilevel=True, purge=False, silent=False):
    if not os.path.exists(src):
        raise Exception("Can't find src dir '%s'" % src)
    if not os.path.exists(dest):
        raise Exception("Can't find dest dir '%s'" % dest)

    flags = []
    if multilevel and os.name=="nt":
        flags.append("/s")
    elif not multilevel and os.name=="posix":
        flags.extend(["-f", "- /*/"])
        # https://superuser.com/questions/436070/rsync-copying-directory-contents-non-recursively

    if purge and os.name=="nt":
        flags.append("/purge")
    elif purge and os.name=="posix":
        flags.append("--delete-before")

    if silent and os.name=="nt":
        flags.extend(["/NFL", "/NDL", "/NJH", "/NJS", "/nc", "/ns", "/np"])
        # https://stackoverflow.com/questions/3898127/how-can-i-make-robocopy-silent-in-the-command-line-except-for-progress
    elif silent and os.name=="posix":
        flags.append("-q")
        # https://serverfault.com/questions/547106/run-totally-silent-rsync

    if os.name=="nt":
        if not silent:
            print("Attempting to run robocopy..." + colorama.Fore.YELLOW)
        returncode = subprocess.call(["robocopy", src, dest, "/compress"] + flags)
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        # https://stackoverflow.com/questions/13161659/how-can-i-call-robocopy-within-a-python-script-to-bulk-copy-multiple-folders

    elif os.name=="posix":
        if not silent:
            print("Attempting to run rsync..." + colorama.Fore.YELLOW)
        CompProc = subprocess.run(["rsync", "-azivh"] + flags + ["%s/" % src,
                                        "%s/" % dest], stderr=subprocess.STDOUT)

    if not silent:
        print(colorama.Style.RESET_ALL)

    # Check for success
    if (os.name=="nt" and returncode < 8) or (os.name=="posix" and CompProc.returncode == 0):
        # https://superuser.com/questions/280425/getting-robocopy-to-return-a-proper-exit-code
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        if not silent:
            print("Sync to '%s' successful\n" % os.path.basename(dest))
    else:
        raise Exception("SYNC to '%s' FAILED" % os.path.basename(dest))


def back_up_remote(src=DIR_REMOTE_SRC, dest_root=DIR_REMOTE_BU):
    if not os.path.exists(src):
        raise Exception("Can't find src dir '%s'" % src)
    if not os.path.exists(dest_root):
        raise Exception("Can't find dest_root dir '%s'" % dest_root)

    # Back up remote source contents before datestamping files on remote.
    sync_remote(src, os.path.join(dest_root, "mirror"), purge=True, silent=True)
    # Removes any extraneous files from local import folder that don't exist in remote.

    sync_remote(src, os.path.join(dest_root, "union"), silent=True)
    # Leaves all in place


def remote_updates(src=DIR_REMOTE_SRC, dest=DIR_IMPORT):
    """1. Back up remote files locally.
       2. Update filenames in remote source where new raw files appear.
       3. Sync renamed remote source files to shared folder.
       4. Sync renamed remote source files locally.
    """
    if not os.path.exists(dest):
        raise Exception("Can't find dest dir '%s'" % dest)

    # Pull from remote dir.
    print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate filenames '
                    'in remote directory (datestamp) "%s"? [Y / N]' % src)
    update_remote_filenames = input("> " + colorama.Style.RESET_ALL)
    if update_remote_filenames.upper() == "Y":
        try:
            print("Backing up remote files...")
            back_up_remote()
            print("...done")

            print("Updating remote filenames...")
            datestamp_remote()
        except KeyboardInterrupt:
            print("User aborted.\n")
        else:
            print("...done\n")

        # Also back up to shared folder for reference.
        print("Syncing source files to shared folder...")
        try:
            sync_remote(DIR_REMOTE_SRC, os.path.join(DIR_REMOTE_SHARE_CTRL, "Raw"), purge=True)
        except KeyboardInterrupt:
            print("User aborted.\n")
        else:
            print("...done")
    else:
        # Accept any answer other than Y/y as negative.
        print("Skipping remote-dir filename updates.")

    # Sync renamed remote source files locally.
    print(colorama.Fore.GREEN + colorama.Style.BRIGHT + '\nUpdate local import '
                                                    'dir from remote? [Y / N]')
    update_import_dir = input("> " + colorama.Style.RESET_ALL)
    if update_import_dir.upper() == "Y":
        print("Updating local files from remote dir...")
        sync_remote(os.path.join(DIR_REMOTE_SRC, "CDF Files/"), DIR_IMPORT, purge=True)
        sync_remote(os.path.join(DIR_REMOTE_SRC, "CPF Files/"), DIR_IMPORT)
        print("...done")
    else:
        print("Skipping import-dir update from remote.\n")


def convert_file(cxf_path, target_dir, tmp_dir=None, check_sn=False, gui_in_focus=False):
    """
    Converts either a CPF or CDF to Excel format.
    tmp_dir path only required for processing CPFs.
    check_sn indicates whether to validate vehicle S/N in filename.
    """
    if not os.path.exists(cxf_path):
        raise Exception("Can't find src file '%s'" % cxf_path)
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    file_type = os.path.splitext(cxf_path)[-1]
    cxf_name = os.path.basename(cxf_path)

    if file_type.lower() == ".cpf" and tmp_dir is None:
        tmp_dir_try = os.path.join(target_dir, "tmp")
        if os.path.exists(tmp_dir_try):
            print(colorama.Fore.CYAN + colorama.Style.BRIGHT)
            print("Use %s as temp dir for CPF processing? [Y/N]" % tmp_dir_try)
            print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
            answer = input("> " + colorama.Style.RESET_ALL)
            if answer.lower() != "y":
                raise Exception("Need temp dir to process CPF files.")
            # Fall through to assignment below
        else:
            os.mkdir(tmp_dir_try) # Will leave in place after processing finished.
        tmp_dir = tmp_dir_try

    if not gui_in_focus:
        select_program(os.path.splitext(cxf_path)[-1][1:])

    if file_type.lower() == ".cpf":
        cpf_open = False
        # Open CPF in GUI and export parameters if export doesn't exist already.
        cpf_param_export_filename = os.path.splitext(cxf_name)[0] + CPF_PARAM_EXPORT_SUFFIX
        if not os.path.exists(os.path.join(tmp_dir, cpf_param_export_filename)):
            cpf_open = open_cpf(cxf_path)
            cpf_params_path = export_cpf_params(tmp_dir, cpf_param_export_filename,
                                                            validate_sn=check_sn)
        else:
            cpf_params_path = os.path.join(tmp_dir, cpf_param_export_filename)

        # Open CPF in GUI and export faults if export doesn't exist already.
        cpf_fault_export_filename = os.path.splitext(cxf_name)[0] + CPF_FAULT_EXPORT_SUFFIX
        if not os.path.exists(os.path.join(tmp_dir, cpf_fault_export_filename)):
            if not cpf_open:
                cpf_open = open_cpf(cxf_path)

            # Export faults
            cpf_faults_path = export_cpf_faults(tmp_dir, cpf_fault_export_filename)

        else:
            # If it already exists in temp dir from previous processing.
            cpf_faults_path = os.path.join(tmp_dir, cpf_fault_export_filename)

        # Combine both tsvs to single export file.
        cpf_combined_export_filename = os.path.splitext(cxf_name)[0] + CPF_COMBINED_EXPORT_SUFFIX
        cpf_combined_export_path = os.path.join(target_dir, cpf_combined_export_filename)
        fixcpf.combine_param_and_fault_export(cpf_params_path, cpf_faults_path, cpf_combined_export_path)
        return True

    elif file_type.lower() == ".cdf":
        valid_cdf = open_cdf(cxf_path)

        if valid_cdf:
            cdf_export_filename = os.path.splitext(cxf_name)[0] + CDF_EXPORT_SUFFIX
            export_path = export_cdf(target_dir, cdf_export_filename, validate_sn=check_sn)
            # select_program("cdf") # Inconsistent Excel behavior - sometimes steals focus and sometimes doesn't
            return True
        else:
            print("\n\tSkipping %s (empty file)." % os.path.basename(cxf_path))
            return False


def select_program(filetype):
    # Brings conversion program into focus.
    answer = gui.confirm("Bring %s-conversion GUI into focus, make sure CAPSLOCK "
                                    "is off, then click OK." % filetype.upper())
    if answer == "OK":
        print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + "\nGUI interaction "
                    "commencing (%s). Move mouse pointer to upper left of "
                    "screen to abort." % filetype.upper() + colorama.Style.RESET_ALL)
    else:
        raise UserCancel()


def open_cpf(file_path):
    if not os.path.exists(file_path):
        raise Exception("Can't find file_path '%s'" % file_path)

    # Assumes 1314 program already in focus.
    # Get to import folder
    gui.hotkey("ctrl", "o")

    gui.hotkey("ctrl", "l") # Select address bar

    gui.typewrite(os.path.dirname(file_path)) # Navigate to import folder.
    gui.press(["enter"])

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(os.path.basename(file_path))
    gui.press(["enter"]) # Confirm filename to open.
    time.sleep(1) # Allow time for to open.

    return True


def export_cpf_params(target_dir, output_filename, validate_sn):
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    # Assumes 1314 program already in focus.
    gui.hotkey("alt", "f") # Open File menu (toolbar).
    gui.press(["e"]) # Select Export from File menu.

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(output_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.2)

    # Check if new file exists in exported location as expected after conversion.
    export_path = os.path.join(target_dir, output_filename)
    assert os.path.exists(export_path), "Can't confirm output file existence."

    if validate_sn:
        match = check_cpf_vehicle_sn(export_path)
        if not match:
            # check_cpf_vehicle_sn() may prompt user to ack. Re-focus CPF program after.
            select_program("cpf")

    return export_path


def check_cpf_vehicle_sn(cpf_param_path):
    cpf_param_filename = os.path.basename(cpf_param_path)

    vehicle_sn_stored = fixcpf.parse_cpf_vehicle_sn(cpf_param_path)
    prompt_str = "Can\'t parse S/N from cpf_param_filename \"%s\".\n" \
                                                "Type S/N manually: " % cpf_param_filename
    vehicle_sn_from_filename = find_in_string(SN_REGEX, cpf_param_filename, prompt_str)
    # print("%s\tfrom filename." % vehicle_sn_from_filename) # DEBUG
    # print("%s\tstored in CPF." % vehicle_sn_stored) # DEBUG

    if vehicle_sn_stored is None:
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("No S/N found in \"%s\". Press Enter to continue." % cpf_param_filename + colorama.Style.RESET_ALL)
        return False
    elif hex(int(vehicle_sn_stored)) == "0xffffffff":
        # If vehicle S/N was not written to controller, S/N value in CPF export
        # will be "4294967295", which translates to "0xFFFFFFFF" in hex.
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("S/N not stored in controller: Found %s in \"%s\".\nPress Enter to continue."
                % (hex(int(vehicle_sn_stored)), cpf_param_filename) + colorama.Style.RESET_ALL)
        return False
    elif vehicle_sn_stored != vehicle_sn_from_filename:
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("S/N mismatch: %s in \"%s\".\nEvaluate and fix filenames if needed "
                                "(import and export).\nPress Enter to continue."
                % (vehicle_sn_stored, cpf_param_filename) + colorama.Style.RESET_ALL)
        return False
    else:
        return True


def export_cpf_faults(target_dir, output_filename):
    global ERROR_HISTORY_SAVE_BUTTON_LOC # Allow modification of global variable

    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    # Assumes 1314 program already in focus.

    gui.hotkey("ctrl", "4") # Diagnostics tab

    # Click on Save button inside Error History tab (different than Ctrl+S save)
    # Mouse hovering over Save icon from previous export changes button appearance.
    if ERROR_HISTORY_SAVE_BUTTON_LOC is None:
        loc_tuple = gui.locateCenterOnScreen(ERROR_HISTORY_SAVE_IMG)
        if loc_tuple is None:
            raise Exception("Can't find Error History save button.")
        else:
            ERROR_HISTORY_SAVE_BUTTON_LOC = loc_tuple # Update global variable.
    else:
        # Use previously-found button if coordinates stored already.
        loc_tuple = ERROR_HISTORY_SAVE_BUTTON_LOC

    # If no faults present in CPF, Save button will be absent. Will fail to export cpf-faults file.

    # loc_tuple = gui.locateCenterOnScreen(ERROR_HISTORY_SAVE_IMG)
    # if loc_tuple is None:
    #     # If no faults present in CPF, Save button will be absent.
    #     # Look for greyed-out Save icon.
    #     loc_tuple = gui.locateCenterOnScreen(ERROR_HISTORY_BLANK)
    #     if loc_tuple is None:
    #         raise Exception("Can't find Error History save button.")
    #     else:
    #         print("\nNo faults present in %s" % output_filename)
    #         return None

    x, y = loc_tuple
    gui.click(x, y)

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(output_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.2)

    # Check if new file exists in exported location as expected after conversion.
    export_path = os.path.join(target_dir, output_filename)
    if not os.path.exists(export_path):
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
        print("\nCan't confirm output file existence (\"%s\").\nEmpty fault history [Y/N]? " % output_filename)
        answer = input("> " + colorama.Style.RESET_ALL)
        if answer.upper() == "Y":
            select_program("cpf")
            export_path = None
        else:
            # Accept anything other than a blank input or 'Y' as a No.
            raise Exception("Can't find cpf_faults file '%s'" % output_filename)

    gui.hotkey("ctrl", "f4") # Close CPF file.
    return export_path


def open_cdf(file_path):
    if not os.path.exists(file_path):
        raise Exception("Can't find file_path '%s'" % file_path)

    # Assumes CIT program already in focus.
    # Ensure file is nonzero size. CIT gives error window for empty file.
    if not os.path.getsize(file_path):
        # Skip empty file
        return False

    # Assumes CIT project open and Programmer window open, in focus.
    # Import file
    gui.press(["alt"])
    gui.press(["f"])
    gui.press(["i"])
    gui.press(["c"])

    gui.press(["enter"]) # Confirm node to use.

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(os.path.dirname(file_path)) # Navigate to import folder.
    gui.press(["enter"])

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(os.path.basename(file_path))
    gui.press(["enter"]) # Confirm filename to open.
    time.sleep(1) # Allow time for file to open.

    return True


def export_cdf(target_dir, output_filename, validate_sn):
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    # Assumes CIT project open and Programmer window open, in focus.
    # Export spreadsheet
    gui.press(["alt"])
    gui.press(["f"])
    gui.press(["e"])
    gui.press(["s"])

    gui.hotkey("alt", "n") # Select filename field
    gui.typewrite(output_filename)

    gui.hotkey("ctrl", "l") # Select address bar
    gui.typewrite(target_dir) # Navigate to target export folder.
    gui.press(["enter"])
    gui.hotkey("alt", "s") # Save
    time.sleep(0.75)

    gui.press(["enter"]) # Click through error

    time.sleep(20) # Allow time for it to write and open Excel file.
    # CIT opens .xlsx export automatically.
    # Close export (doesn't always work):
    export_path = os.path.join(target_dir, output_filename)
    book = xw.Book(export_path)
    book.close()

    if validate_sn:
        match = check_cdf_vehicle_sn(export_path)
        if not match:
            # check_cdf_vehicle_sn() may prompt user to ack. Re-focus CDF program after.
            select_program("cdf")

    # Re-focus on CIT.
    gui.click(1477, 17) # Temporary workaround to click title bar of CIT.
    return os.path.join(target_dir, output_filename)


def extract_cdf_vehicle_sn(export_filepath):

    param_df = pd.read_excel(export_filepath, sheet_name="Parameters")
    for _, row in param_df.iterrows():
        if row["Variable Name"] == "nvuser4":
            # Check if VCL Alias column available (old CIT versions don't include it.)
            if "VCL Alias" in param_df.columns:
                error_text = ("Expected 'VCL Alias' of 'nvuser4' variable to be "
                "'NV_VehicleSerialNumber', but instead is '%s'." % row["VCL Alias"])
                assert row["VCL Alias"] == "NV_VehicleSerialNumber", error_text

            vehicle_sn_param = row["Application Default"]

    if not vehicle_sn_param or pd.isna(vehicle_sn_param):
        # Empty value
        return None
    elif vehicle_sn_param.isdecimal() and hex(int(vehicle_sn_param)) == "0xffffffff":
        # If vehicle S/N was not written to controller, S/N value in CDF export
        # will be "4294967295", which translates to "0xFFFFFFFF" in hex.
        # https://stackoverflow.com/questions/44891070/whats-the-difference-between-str-isdigit-isnumeric-and-isdecimal-in-pyth
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("S/N not stored in controller: Found '%s' in %s.\nPress Enter to continue."
                % (hex(int(vehicle_sn_param)), os.path.basename(export_filepath)) + colorama.Style.RESET_ALL)
        return None

    # Validate that S/N value conforms to expected format.
    prompt_str = ("Found multiple possible S/N values stored in CDF: '%s'. Press Enter to continue." % vehicle_sn_param)
    valid_sn = find_in_string(SN_REGEX, vehicle_sn_param, prompt_str, allow_none=True)
    if valid_sn is None:
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("Expected 'nvuser4' variable to contain S/N in 7-digit format starting "
                        "with 3, 5, or 8.\nFound '%s' in %s instead."
                        % (vehicle_sn_param, os.path.basename(export_filepath))
                                                    + colorama.Style.RESET_ALL)
        return None
    elif valid_sn != vehicle_sn_param:
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("'nvuser4' value '%s' (in %s) appears to contain S/N with right format but "
                        "may contain additional content."
                        % (vehicle_sn_param, os.path.basename(export_filepath))
                                                     + colorama.Style.RESET_ALL)
        return None
    else:
        return vehicle_sn_param # string


def check_cdf_vehicle_sn(cdf_path):
    cdf_filename = os.path.basename(cdf_path)

    vehicle_sn_stored = extract_cdf_vehicle_sn(cdf_path)

    prompt_str = "Can\'t parse S/N from cdf_filename \"%s\".\n" \
                                                "Type S/N manually: " % cdf_filename
    vehicle_sn_from_filename = find_in_string(SN_REGEX, cdf_filename, prompt_str)

    if vehicle_sn_stored is None:
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("No valid S/N found in \"%s\". Press Enter to continue." % cdf_filename + colorama.Style.RESET_ALL)
        return False
    elif vehicle_sn_stored != vehicle_sn_from_filename:
        print(colorama.Fore.RED + colorama.Style.BRIGHT)
        input("S/N mismatch: %s in \"%s\".\nEvaluate and fix filenames if needed "
                                "(import and export).\nPress Enter to continue."
                % (vehicle_sn_stored, cdf_filename) + colorama.Style.RESET_ALL)
        return False
    else:
        return True


def convert_all(file_type, source_dir, dest_dir, check_SNs=False):
    if not os.path.exists(source_dir):
        raise Exception("Can't find source_dir '%s'" % source_dir)
    if not os.path.exists(dest_dir):
        raise Exception("Can't find dest_dir '%s'" % dest_dir)

    try:
        select_program(file_type)
    except UserCancel:
        return

    file_list = [x for x in sorted(os.listdir(source_dir)) if x.lower().endswith(file_type)]
    for filename in tqdm(file_list, colour="cyan"):
        # Check for existing export
        if file_type == "cpf" and (os.path.exists(os.path.join(DIR_EXPORT,
                            os.path.splitext(filename)[0] + CPF_COMBINED_EXPORT_SUFFIX))):
            # Skip if already processed this file.
            tqdm.write("Already processed %s" % os.path.basename(filename)) # DEBUG
            continue
        elif file_type == "cdf" and (os.path.exists(os.path.join(DIR_EXPORT,
                            os.path.splitext(filename)[0] + CDF_EXPORT_SUFFIX))):
            # Skip if already processed this file.
            tqdm.write("Already processed %s" % os.path.basename(filename)) # DEBUG
            continue

        filepath = os.path.join(source_dir, filename)
        if (os.path.isfile(filepath) and
                    os.path.splitext(filename)[-1].lower() == ".%s" % file_type):
            try:
                success = convert_file(filepath, dest_dir,
                                tmp_dir=DIR_EXPORT_BUFFER, check_sn=check_SNs,
                                                              gui_in_focus=True)
            except Exception as exception_text:
                print(colorama.Fore.CYAN + colorama.Style.BRIGHT)
                print("\nEncountered exception processing %s" % filename + colorama.Style.RESET_ALL)
                print(exception_text)
                print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
                print("Press Enter to continue with other files, 'e' to exit "
                                "file-conversion loop, or 'q' to quit program.")
                answer = input("> " + colorama.Style.RESET_ALL)
                if answer.lower() == "":
                    select_program(file_type)
                    continue
                elif answer.lower() == "e":
                    break
                else:
                    # Accept anything other than a blank input or 'e' as a quit command.
                    quit()
            else:
                if success:
                    tqdm.write("Processed %s" % filename)

        else:
            # Skip directories
            continue


def convert_cpfs_in_export(dir_path):
    """Convert CPF exports (.XLS extension but TSV format) to true Excel format."""
    if not os.path.exists(dir_path):
        raise Exception("Can't find dir_path '%s'" % dir_path)

    print("\nConverting CPF exports from tsv format (named .XLS) to .xslx (in dir "
                                                    "\n\t\"%s\")..." % dir_path)
    try:
        fixcpf.convert_all_param_exports(dir_path, check_xls=False)
        print("...done")
    except PermissionError:
        # Gets a PermissionError if running on PowerShell most of the time.
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
        input("\nEncountered permission error in removing CPF tsv files.\n"
                        "Press Enter to continue to next part of program.")
        print(colorama.Style.RESET_ALL)


def create_file_struct():
    # Make field-data dirs if any don't exist yet.
    if not os.path.exists(DIR_FIELD_DATA):
        os.mkdir(DIR_FIELD_DATA)
        print("Created %s" % DIR_FIELD_DATA)

    if not os.path.exists(DIR_IMPORT_ROOT):
        os.mkdir(DIR_IMPORT_ROOT)
        print("Created %s" % DIR_IMPORT_ROOT)

    if not os.path.exists(DIR_REMOTE_BU):
        os.mkdir(DIR_REMOTE_BU)
        print("Created %s" % DIR_REMOTE_BU)

    if not os.path.exists(DIR_IMPORT):
        os.mkdir(DIR_IMPORT)
        print("Created %s" % DIR_IMPORT)

    if not os.path.exists(DIR_EXPORT):
        os.mkdir(DIR_EXPORT)
        print("Created %s" % DIR_EXPORT)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Program to convert CPF or CDF "
                                    "exports from binary to .xlsx file format.")
    parser.add_argument("-d", "--dir", help="Specify dir containing exports "
                                                        "to convert.", type=str)
    # parser.add_argument("-f", "--file", help="Specify file path of one export  " # maybe implement later
    #                                                     "to reformat.", type=str)
    # parser.add_argument("-a", "--auto", help="Specify to execute entire routine " # maybe implement later
    #             "of downloading new exports from remote drive, converting all, "
    #                     "and uploading to Azure blob.", action="store_true")
    args = parser.parse_args()

    # Default is auto-run, but if user specifies --dir, disable auto-run.

    if args.dir:
        auto_run = False
        check_vehicle_sns = False
        import_dir = args.dir
        export_dir = args.dir
    else:
        auto_run = True
        check_vehicle_sns = True
        import_dir = DIR_IMPORT
        export_dir = DIR_EXPORT
        # Set up directory structure if absent on local machine.
        create_file_struct()
        # Remote source backup, filename updates, sync remote locally and to shared folder
        remote_updates()
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
        print("Press Enter to proceed to file processing or 'q' to quit program.")
        answer = input("> " + colorama.Style.RESET_ALL)
        if answer == "":
            pass
        else:
            # Accept anything other than a blank input as a quit command.
            quit()

    # Convert exports
    if os.name == "nt":
        try:
            convert_all("cpf", import_dir, export_dir, check_SNs=check_vehicle_sns)
            convert_all("cdf", import_dir, export_dir, check_SNs=check_vehicle_sns)
            print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + "\nGUI "
                                                            "interaction done\n")
            print(colorama.Style.RESET_ALL)
        except gui.FailSafeException:
            print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + "\n\nUser "
                                                    "canceled GUI interaction.")
            print(colorama.Style.RESET_ALL)
            time.sleep(3)
            # If user terminates GUI interraction, continue running below.
            pass
    else:
        print(colorama.Fore.MAGENTA + colorama.Style.BRIGHT + "Skipping GUI "
                                    "interaction (requires Windows system).")
        print(colorama.Style.RESET_ALL)


    if auto_run:
        # Sync to shared folder
        print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
        print("\nSync controller export dir to shared folder? Enter to proceed, "
                                            "'s' to skip, or 'q' to quit program.")
        answer = input("> " + colorama.Style.RESET_ALL)
        if answer == "":
            print("Syncing processed files to shared folder...")
            sync_remote(DIR_EXPORT, os.path.join(DIR_REMOTE_SHARE_CTRL, "Converted"),
                                                    purge=True, multilevel=False)
            print("...done")
        elif answer.lower() == "s":
            print("Skipping shared-folder sync.")
        else:
            # Accept anything other than a blank input as a quit command.
            quit()


        # Sync to second remote (Azure blob)
        if os.name=="nt":
            # Controller exports
            print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
            print("\nSync controller export dir to Azure blob? Enter to proceed, "
                                            "'s' to skip, or 'q' to quit program.")
            answer = input("> " + colorama.Style.RESET_ALL)
            if answer == "":
                print("\nRunning AzCopy sync job (controller exports)...")
                print(colorama.Fore.BLUE + colorama.Style.BRIGHT)
                returncode = subprocess.call(["azcopy", "sync",
                                                    "--delete-destination", "true",
                                            DIR_EXPORT + "\\", AZ_BLOB_ADDR_CTRL])
                # https://learn.microsoft.com/en-us/azure/storage/common/storage-ref-azcopy-sync
                print(colorama.Style.RESET_ALL + "...done")
            elif answer.lower() == "s":
                print("Skipping sync from ctrl-export folder to shared folder.")
            else:
                # Accept anything other than a blank input as a quit command.
                quit()

            # Batt export dir
            print(colorama.Fore.GREEN + colorama.Style.BRIGHT)
            print("\nSync battery export dir from shared folder to Azure blob? "
                                        "Enter to proceed or 'q' to quit program.")
            answer = input("> " + colorama.Style.RESET_ALL)
            if answer == "":
                print("\nRunning AzCopy sync job (batt export)...")
                print(colorama.Fore.BLUE + colorama.Style.BRIGHT)
                returncode = subprocess.call(["azcopy", "sync",
                                                    "--delete-destination", "true",
                                    DIR_REMOTE_SHARE_BATT + "\\", AZ_BLOB_ADDR_BATT])
                print(colorama.Style.RESET_ALL + "...done")
            else:
                print("Skipping sync from batt dir to shared folder.")

        else:
            print("Skipping AzCopy jobs (requires Windows system).")

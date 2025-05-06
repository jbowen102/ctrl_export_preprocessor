import os
import time
import re
import subprocess
import shutil

import argparse
import pandas as pd
from tqdm import tqdm
from colorama import Style, Fore, Back
import xlwings as xw
if os.name == "nt":
    # Allows testing other (non-GUI) features in WSL where pyautogui import fails
    import pyautogui as gui
    gui.FAILSAFE = True
    # Allows moving mouse to upper-left corner of screen to abort execution.
    # https://pyautogui.readthedocs.io/en/latest/quickstart.html


try:
    import fix_cpf_export_format as fixcpf
    from sw_rev_mapping import REV_MAP_ALL_F
    from dir_names import DIR_REMOTE_SRC, \
                          DIR_FIELD_DATA, \
                              DIR_IMPORT_ROOT, DIR_REMOTE_BU, DIR_IMPORT, \
                              DIR_EXPORT, \
                          DIR_REMOTE_SHARE_CTRL, DIR_REMOTE_SHARE_BATT, DIR_REMOTE_SHARE_MFG, \
                          AZ_BLOB_ADDR_CTRL, AZ_BLOB_ADDR_BATT, AZ_BLOB_ADDR_MFG, \
                          ERROR_HISTORY_SAVE_IMG, ERROR_HISTORY_BLANK
except ModuleNotFoundError:
    import ctrl_export_preprocessor.fix_cpf_export_format as fixcpf
    from ctrl_export_preprocessor.sw_rev_mapping import REV_MAP_ALL_F
    from ctrl_export_preprocessor.dir_names import DIR_REMOTE_SRC, \
                                                DIR_FIELD_DATA, \
                                                    DIR_IMPORT_ROOT, DIR_REMOTE_BU, DIR_IMPORT, \
                                                    DIR_EXPORT, \
                                                DIR_REMOTE_SHARE_CTRL, DIR_REMOTE_SHARE_BATT, DIR_REMOTE_SHARE_MFG, \
                                                AZ_BLOB_ADDR_CTRL, AZ_BLOB_ADDR_BATT, AZ_BLOB_ADDR_MFG, \
                                                ERROR_HISTORY_SAVE_IMG, ERROR_HISTORY_BLANK


DATE_REGEX_1 = r"(20\d{2}[0-1]\d[0-3]\d)"
DATE_REGEX_2 = r"(20\d{2}-[0-1]\d-[0-3]\d)"
# Could catch some invalid dates like 20231131. Further validated below in find_in_string()
DATE_FORMAT_1 = "%Y%m%d"
DATE_FORMAT_2 = "%Y-%m-%d"
DATE_FORMAT_3 = "%m%d%Y"
DATE_FORMATS = [DATE_FORMAT_1, DATE_FORMAT_2, DATE_FORMAT_3]

SN_REGEX = r"(3\d{6}|5\d{6}|8\d{6})"
# Any "3" or "5" or "8" followed by six more digits
CDF_SW_PN_REGEX = r"\d{6}\.\d{2}|\d{8}\.\d{2}"
SW_PN_REGEX = r"\d{6}G\d{2}|\d{8}G\d{2}"


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
    prompted = False
    while not found:
        matches = re.findall(regex_pattern, string_to_search, flags=re.IGNORECASE)
        if len(matches) == 1 and date_target:
            # If looking for a date, check for valid date value (regex doesn't fully validate)
            # print("\t\tmatches[0]: " + matches[0]) # DEBUG
            for date_format in DATE_FORMATS:
                try:
                    time.strptime(matches[0], date_format)
                except ValueError:
                    continue
                else:
                    # Valid date
                    return matches[0], prompted
        elif len(matches) == 1:
            return matches[0], prompted
            # loop exits
        elif len(matches) == 0 and allow_none:
            # print("\t\t%s: no matches; returning None" % string_to_search) # DEBUG
            return None, prompted

        # No matches, multiple matches, or invalid date found:
        prompted = True
        print(Fore.GREEN + Style.BRIGHT + prompt)
        string_to_search = input(">" + Style.RESET_ALL)


def datestamp_remote(remote=DIR_REMOTE_SRC):
    while not os.path.exists(remote):
        # Prompt user to mount network drives if not found.
        print(Fore.GREEN + Style.BRIGHT + '\n"%s" not found. Mount '
                                'and press Enter to try again.' % src)
        input("> " + Style.RESET_ALL)

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
                    serial_num, _ = find_in_string(SN_REGEX, item_name, prompt_str)
                    # Now look for date in remaining string. Will add later if not present.
                    # print("\tReceived %s as S/N back from find_in_string()" % serial_num) # DEBUG
                    remaining_str = item_name.split(serial_num)
                    date_found = False
                    for substring in remaining_str:
                        prompt_str = "Can't find single valid date match in import " \
                                                            "filename \"%s\".\n" \
                                "Type manually (YYYYMMDD format): " % file_name
                        # print("\tDate:") # DEBUG
                        date_match, _ = find_in_string(DATE_REGEX_1, substring,
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
                        if mod_date < time.strptime("20200101", DATE_FORMAT_1):
                            # Substitute in today's date
                            date_to_use = time.localtime()
                        else:
                            date_to_use = mod_date

                        datestamp = time.strftime(DATE_FORMAT_1, date_to_use)

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
            print(Fore.MAGENTA + "\t%s\t->\t%s" % (old_names[i], new_names[i]))
        input(Fore.GREEN + Style.BRIGHT + "\nPress Enter to continue"
                                                    + Style.RESET_ALL)
    else:
        print(Fore.MAGENTA + "\t[None]" + Style.RESET_ALL)
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
            print("Attempting to run robocopy..." + Fore.YELLOW)
        returncode = subprocess.call(["robocopy", src, dest, "/compress"] + flags)
        # https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
        # https://stackoverflow.com/questions/13161659/how-can-i-call-robocopy-within-a-python-script-to-bulk-copy-multiple-folders

    elif os.name=="posix":
        if not silent:
            print("Attempting to run rsync..." + Fore.YELLOW)
        CompProc = subprocess.run(["rsync", "-azivh"] + flags + ["%s/" % src,
                                        "%s/" % dest], stderr=subprocess.STDOUT)

    if not silent:
        print(Style.RESET_ALL)

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


def update_from_remote_dirs(src=DIR_REMOTE_SRC, dest=DIR_IMPORT):
    """1. Back up remote files locally.
       2. Update filenames in remote source where new raw files appear.
       3. Sync renamed remote source files to shared folder.
       4. Sync renamed remote source files locally.
    """
    if not os.path.exists(dest):
        raise Exception("Can't find dest dir '%s'" % dest)

    # Pull from remote dir.
    print(Fore.GREEN + Style.BRIGHT + '\nUpdate filenames '
                    'in remote directory (datestamp) "%s"? [Y / N]' % src)
    update_remote_filenames = input("> " + Style.RESET_ALL)
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
    print(Fore.GREEN + Style.BRIGHT + '\nUpdate local import '
                                                    'dir from remote? [Y / N]')
    update_import_dir = input("> " + Style.RESET_ALL)
    if update_import_dir.upper() == "Y":
        print("Updating local files from remote dir...")
        sync_remote(os.path.join(DIR_REMOTE_SRC, "CDF Files/"), DIR_IMPORT, purge=True)
        sync_remote(os.path.join(DIR_REMOTE_SRC, "CPF Files/"), DIR_IMPORT)
        print("...done")
    else:
        print("Skipping import-dir update from remote.\n")


def convert_file(cpf_path, target_dir, check_sn=False, gui_in_focus=False):
    """
    Converts a CPF to Excel format.
    temp_dir path required for processing CPFs.
    check_sn indicates whether to validate vehicle S/N in filename.
    """
    if not os.path.exists(cpf_path):
        raise Exception("Can't find src file '%s'" % cpf_path)
    if not os.path.exists(target_dir):
        raise Exception("Can't find target_dir '%s'" % target_dir)

    file_type = os.path.splitext(cpf_path)[-1]
    cpf_name = os.path.basename(cpf_path)

    temp_dir = os.path.join(target_dir, "tmp")
    if not os.path.exists(temp_dir):
        os.mkdir(temp_dir) # Will leave in place after processing finished.

    if not gui_in_focus:
        select_program(os.path.splitext(cpf_path)[-1][1:])

    cpf_open = False
    # Open CPF in GUI and export parameters if export doesn't exist already.
    cpf_param_export_filename = os.path.splitext(cpf_name)[0] + CPF_PARAM_EXPORT_SUFFIX
    if not os.path.exists(os.path.join(temp_dir, cpf_param_export_filename)):
        cpf_open = open_cpf(cpf_path)
        cpf_params_path = export_cpf_params(temp_dir, cpf_param_export_filename,
                                                        validate_sn=check_sn)
    else:
        cpf_params_path = os.path.join(temp_dir, cpf_param_export_filename)

    # Open CPF in GUI and export faults if export doesn't exist already.
    cpf_fault_export_filename = os.path.splitext(cpf_name)[0] + CPF_FAULT_EXPORT_SUFFIX
    if not os.path.exists(os.path.join(temp_dir, cpf_fault_export_filename)):
        if not cpf_open:
            cpf_open = open_cpf(cpf_path)

        # Export faults
        cpf_faults_path = export_cpf_faults(temp_dir, cpf_fault_export_filename)

    else:
        # If it already exists in temp dir from previous processing.
        cpf_faults_path = os.path.join(temp_dir, cpf_fault_export_filename)

    # Combine both tsvs to single export file.
    cpf_combined_export_filename = os.path.splitext(cpf_name)[0] + CPF_COMBINED_EXPORT_SUFFIX
    cpf_combined_export_path = os.path.join(target_dir, cpf_combined_export_filename)
    fixcpf.combine_param_and_fault_export(cpf_params_path, cpf_faults_path, cpf_combined_export_path)
    return True


def select_program(filetype):
    # Brings conversion program into focus.
    proj_file_msg = ""
    answer = gui.confirm("%sBring %s-conversion GUI into focus, make sure CAPSLOCK "
                    "is off, then click OK." % (proj_file_msg, filetype.upper()))
    if answer == "OK":
        print(Fore.MAGENTA + Style.BRIGHT + "\n\nGUI interaction "
                    "commencing (%s). Move mouse pointer to upper left of "
                    "screen to abort." % filetype.upper() + Style.RESET_ALL)
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
    time.sleep(1 * GUI_PAUSE_MULT) # Allow time for to open.

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
    time.sleep(0.2 * GUI_PAUSE_MULT)

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
    vehicle_sn_from_filename, _ = find_in_string(SN_REGEX, cpf_param_filename, prompt_str)
    # print("%s\tfrom filename." % vehicle_sn_from_filename) # DEBUG
    # print("%s\tstored in CPF." % vehicle_sn_stored) # DEBUG

    if vehicle_sn_stored is None:
        print(Fore.RED + Style.BRIGHT)
        print("No S/N found in \"%s\"." % cpf_param_filename + Style.RESET_ALL)
        time.sleep(2)
        return False
    elif hex(int(vehicle_sn_stored)) == "0xffffffff":
        # If vehicle S/N was not written to controller, S/N value in CPF export
        # will be "4294967295", which translates to "0xFFFFFFFF" in hex.
        print(Fore.RED + Style.BRIGHT)
        print("S/N not stored in controller: Found %s in \"%s\"."
                % (hex(int(vehicle_sn_stored)), cpf_param_filename) + Style.RESET_ALL)
        time.sleep(2)
        return False
    elif vehicle_sn_stored != vehicle_sn_from_filename:
        print(Fore.RED + Style.BRIGHT)
        input("S/N mismatch: %s in \"%s\".\nEvaluate and fix filenames if needed "
                                "(import and export).\nPress Enter to continue."
                % (vehicle_sn_stored, cpf_param_filename) + Style.RESET_ALL)
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
            print(Fore.GREEN + Style.BRIGHT)
            print("\nCan't find Error History save button for (\"%s\").\n"
                                "Empty fault history [Y/N]?" % output_filename)
            answer = input("> " + Style.RESET_ALL)
            if answer.upper() == "Y":
                select_program("cpf")
                gui.hotkey("ctrl", "f4") # Close CPF file.
                return None
            else:
                # Accept anything other than a blank input or 'Y' as a No.
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
    time.sleep(0.2 * GUI_PAUSE_MULT)

    # Check if new file exists in exported location as expected after conversion.
    export_path = os.path.join(target_dir, output_filename)
    if not os.path.exists(export_path):
        print(Fore.GREEN + Style.BRIGHT)
        print("\nCan't confirm output file existence (\"%s\").\nEmpty fault history [Y/N]?" % output_filename)
        answer = input("> " + Style.RESET_ALL)
        if answer.upper() == "Y":
            select_program("cpf")
            export_path = None
        else:
            # Accept anything other than a blank input or 'Y' as a No.
            raise Exception("Can't find cpf_faults file '%s'" % output_filename)

    gui.hotkey("ctrl", "f4") # Close CPF file.
    return export_path


class GUI_Driver(object):
    def __init__(self):
        self.gui_in_focus = False

    def gui_is_in_focus(self):
        return self.gui_in_focus

    def lose_focus(self):
        self.gui_in_focus = False

    def select_program(self, filetype):
        # Brings conversion program into focus.
        if filetype.upper() == "CDF":
            proj_file_msg = "Check intended project file is loaded in CIT.\n"
        else:
            proj_file_msg = ""
        answer = gui.confirm("%sBring %s-conversion GUI into focus, make sure CAPSLOCK "
                        "is off, then click OK." % (proj_file_msg, filetype.upper()))
        if answer == "OK":
            print(Fore.MAGENTA + Style.BRIGHT + "\n\nGUI interaction "
                        "commencing (%s). Move mouse pointer to upper left of "
                        "screen to abort." % filetype.upper() + Style.RESET_ALL)
            self.gui_in_focus = True
        else:
            self.gui_in_focus = False
            raise UserCancel()

    def open_cdf(self, file_path):
        if not os.path.exists(file_path):
            raise Exception("Can't find file_path '%s'" % file_path)

        # Requires CIT project open and Programmer window open, in focus.
        if not self.gui_in_focus:
            self.select_program("CDF")

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
        time.sleep(1 * GUI_PAUSE_MULT) # Allow time for file to open.

        return True


    def export_cdf(self, output_filepath):
        target_dir = os.path.dirname(output_filepath)
        output_filename = os.path.basename(output_filepath)

        if not os.path.exists(target_dir):
            raise Exception("Can't find target_dir '%s'" % target_dir)

        # Requires CIT project open and Programmer window open, in focus.
        if not self.gui_in_focus:
            self.select_program("CDF")

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
        time.sleep(0.75 * GUI_PAUSE_MULT)

        gui.press(["enter"]) # Click through error

        time.sleep(20 * GUI_PAUSE_MULT) # Allow time for it to write and open Excel file.
        # CIT opens .xlsx export automatically.
        # Close export (doesn't always work):
        book = xw.Book(output_filepath)
        book.close()

        # Re-focus on CIT.
        # Excel behavior inconsistent.
        # Closing workbook above often leaves a blank instance of Excel anyway.
        # Click title bar of CIT to bring back in focus.
        # Snap CIT to right half of screen and make sure Excel window isn't full-screen.
        gui.click(1477, 17)


class CloneDataFile(object):
    def __init__(self, source_filepath, CDF_DB):
        assert os.path.exists(source_filepath), "Tried to create CDF object w/ invalid filepath: %s" % import_filepath
        self.import_filepath = source_filepath
        self.cdf_filename = os.path.basename(source_filepath)
        self.export_filename = os.path.splitext(self.cdf_filename)[0] + CDF_EXPORT_SUFFIX # Usually doesn't exist yet.
        self.export_path = None       # To be set by convert()

        # Ensure CDF file size is nonzero. CIT gives error window for empty file.
        if os.path.getsize(self.import_filepath):
            self.valid_cdf = True
        else:
            # Empty file
            self.valid_cdf = False

        self.source_ctrl_sw_pn = None # To be set by extract_cdf_source_sw_pn()

        self.vehicle_sn_param = None  # Vehicle S/N stored in controller. Various failure modes can couse this to be wrong.
        self.vehicle_sn = None        # Canonical vehicle S/N after validation. Still may be none if impossible to confidently infer.

        self.ParentDB = CDF_DB

    def is_valid_cdf(self):
        return self.valid_cdf

    def get_cdf_filename(self):
        return self.cdf_filename

    def get_ctrl_sw_rev(self):
        assert self.source_ctrl_sw_pn is not None, "Tried to get %s's ctrl_sw_rev, but it hasn't been set yet" % self.cdf_filename
        return REV_MAP_ALL_F[self.source_ctrl_sw_pn]

    def set_export_dir(self, export_dir_path):
        if self.has_export() and os.path.dirname(self.export_path) != export_dir_path:
            raise Exception("Tried to set different export dir path for %s" % self)
        if not os.path.exists(export_dir_path):
            raise Exception("Can't find target_dir '%s' to export %s" % (export_dir_path, self))

        self.temp_dir = os.path.join(export_dir_path, "tmp")
        if not os.path.exists(self.temp_dir):
            os.mkdir(self.temp_dir) # Will leave in place after processing finished.
        self.export_path = os.path.join(export_dir_path, self.export_filename)

    def has_export(self, export_dir_given=None):
        """Returns True iff CDF object already has a stored, valid filepath in export_path attribute
        or if export_filename exists in given export_dir.
        Does not create persistent export_path attribute using export_dir_given.
        """
        if self.export_path is None and export_dir_given is None:
            return False
        if self.export_path is None:
            # Object hasn't stored an export_path yet, but caller passed in an export_dir to try.
            export_path_try = os.path.join(export_dir_given, self.export_filename)
            if os.path.exists(export_path_try):
                self.export_path = export_path_try
                return True
            else:
                return False
        elif os.path.exists(self.export_path):
            # export_path string stored, and the path is valid.
            return True
        else:
            return False # export_path string stored, but nothing written there yet.

    def remove_export(self):
        assert self.has_export(), "Tried to remove %s export but it doesn't exist." % self
        os.remove(self.export_path)

    def convert(self, target_dir, check_sn=False):
        """
        Converts a CDF to Excel format.
        check_sn indicates whether to validate vehicle S/N in filename.
        """
        assert self.is_valid_cdf(), "Tried to convert an empty file."

        assert not self.has_export(), "Tried to convert %s, but an export already exists here: '%s'" % (self, self.export_path)
        self.set_export_dir(target_dir)

        # See if export exists there already
        assert not self.has_export(), "Tried to process file %s that already has export at %s" % (self.cdf_filename, self.export_path)

        assert self.ParentDB.get_GUI_Driver() is not None, "Tried to convert %s but no GUI driver active." % self
        self.valid_cdf = self.ParentDB.get_GUI_Driver().open_cdf(self.import_filepath)
        if self.valid_cdf:
            self.ParentDB.get_GUI_Driver().export_cdf(self.export_path)
            # select_program("cdf") # Inconsistent Excel behavior - sometimes steals focus and sometimes doesn't
        else:
            print("\n\tSkipping %s (empty file)." % os.path.basename(cdf_path))
            return False

        if check_sn:
            self.check_stored_vehicle_sn()

        valid_alias_mapping = self.check_cprj_rev_match()
        if not valid_alias_mapping:
            self.remove_export()
            return False

        return True

    def check_stored_vehicle_sn(self):
        assert self.has_export(), "Tried to extract vehicle S/N from CDF export, but export doesn't exist.\n\t%s" % self

        prompt_str = "Can\'t parse S/N from cdf_filename \"%s\".\n" \
                                                    "Type S/N manually: " % self.cdf_filename
        vehicle_sn_from_filename, stole_focus = find_in_string(SN_REGEX, self.cdf_filename, prompt_str)
        if stole_focus:
            self.ParentDB.get_GUI_Driver().lose_focus()

        self.extract_stored_vehicle_sn() # Populates self.vehicle_sn_param
        if self.vehicle_sn_param is None:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("No valid S/N found in \"%s\". Press Enter to continue." % self.cdf_filename + Style.RESET_ALL)
        elif self.vehicle_sn_param != vehicle_sn_from_filename:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("S/N mismatch: %s in \"%s\".\nEvaluate and fix filenames if needed "
                                    "(import and export).\nPress Enter to continue."
                    % (self.vehicle_sn_param, self.cdf_filename) + Style.RESET_ALL)
            # TODO: prompt user - should self.vehicle_sn should be set to vehicle_sn_from_filename in this case?
        else:
            self.vehicle_sn = self.vehicle_sn_param

    def extract_stored_vehicle_sn(self):
        CDF_VARIABLE_NAME = "nvuser4"
        assert self.has_export(), "Tried to extract vehicle S/N from CDF export, but export doesn't exist.\n\t%s" % self

        # vehicle S/N in export file
        param_df = pd.read_excel(self.export_path, sheet_name="Parameters")
        for _, row in param_df.iterrows():
            if row["Variable Name"] == CDF_VARIABLE_NAME:
                # Check if VCL Alias column available (old CIT versions don't include it.)
                if "VCL Alias" in param_df.columns:
                    error_text = ("Expected 'VCL Alias' of '%s' variable to be "
                                    "'NV_VehicleSerialNumber', but instead is '%s'."
                                            % (CDF_VARIABLE_NAME, row["VCL Alias"]))
                    assert row["VCL Alias"].lower() == "nv_vehicleserialnumber", error_text

                vehicle_sn_param = row["Application Default"]

        if not vehicle_sn_param or pd.isna(vehicle_sn_param):
            # Empty value
            self.vehicle_sn_param = None
            return
        elif vehicle_sn_param.isdecimal() and hex(int(vehicle_sn_param)) == "0xffffffff":
            # If vehicle S/N was not written to controller, S/N value in CDF export
            # will be "4294967295", which translates to "0xFFFFFFFF" in hex.
            # https://stackoverflow.com/questions/44891070/whats-the-difference-between-str-isdigit-isnumeric-and-isdecimal-in-pyth
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("S/N not stored in controller: Found '%s' in %s.\nPress Enter to continue."
                    % (hex(int(vehicle_sn_param)), self.export_filename) + Style.RESET_ALL)
            self.vehicle_sn_param = None
            return

        # Validate that S/N value conforms to expected format.
        prompt_str = ("Found multiple possible S/N values stored in CDF: '%s'. Press Enter to continue." % vehicle_sn_param)
        valid_sn, stole_focus = find_in_string(SN_REGEX, vehicle_sn_param, prompt_str, allow_none=True)
        if stole_focus:
            self.ParentDB.get_GUI_Driver().lose_focus()

        if valid_sn is None:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("Expected '%s' variable to contain S/N in 7-digit format starting "
                            "with 3, 5, or 8.\nFound '%s' in %s instead."
                        % (CDF_VARIABLE_NAME, vehicle_sn_param, self.export_filename)
                                                        + Style.RESET_ALL)
            self.vehicle_sn_param = None
        elif valid_sn != vehicle_sn_param:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("'%s' value '%s' (in %s) appears to contain S/N with right format but "
                            "may contain additional content."
                        % (CDF_VARIABLE_NAME, vehicle_sn_param, self.export_filename)
                                                        + Style.RESET_ALL)
            self.vehicle_sn_param = None
        else:
            self.vehicle_sn_param = vehicle_sn_param # string

    def check_cprj_rev_match(self, given_export_dir=None):
        if given_export_dir:
            self.set_export_dir(given_export_dir)
        assert self.has_export(), "Tried to check cprj rev mapping in CDF export, but export doesn't exist.\n\t%s" % self

        cdf_cprj_pn = self.extract_cdf_cprj_pn()
        cprj_map_rev = REV_MAP_ALL_F[cdf_cprj_pn]

        self.extract_cdf_source_sw_pn()

        if self.source_ctrl_sw_pn is None:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("No valid SW P/N found in \"%s\". Cannot confirm valid VCL Alias "
                                                "mapping. Press Enter to continue."
                                        % self.cdf_filename + Style.RESET_ALL)
            return False

        ctrl_sw_rev = REV_MAP_ALL_F[self.source_ctrl_sw_pn]
        if cprj_map_rev != ctrl_sw_rev:
            # self.ParentDB.get_GUI_Driver().lose_focus() # DEBUG
            # print(Fore.RED + Style.BRIGHT) # DEBUG
            # print("%s: cprj rev mismatch - Deleting export\n\tCtrl SW %s rev %s, but project file used (%s) was rev %s."
            #                      % (self, self.source_ctrl_sw_pn, ctrl_sw_rev, cdf_cprj_pn, cprj_map_rev) + Style.RESET_ALL) # DEBUG
            # Caller will delete file.
            return False
        else:
            return True

    def extract_cdf_source_sw_pn(self):
        """Takes in CDF export (.xlsx format), locates source-vehicle's
        controller-software P/N, and returns it as a string.
        """
        assert self.has_export(), "Tried to extract CDF source SW P/N from export, but export doesn't exist.\n\t%s" % self
        VSN_CDF_VAR_NAME = "user119"
        param_df = pd.read_excel(self.export_path, sheet_name="Parameters")
        for _, row in param_df.iterrows():
            if row["Variable Name"] == VSN_CDF_VAR_NAME:
                # Check if VCL Alias column available (old CIT versions don't include it.)
                if "VCL Alias" in param_df.columns:
                    error_text = ("Expected 'VCL Alias' of '%s' variable to be "
                                    "'ApplicationNameAsInt32', but instead is '%s'."
                                            % (VSN_CDF_VAR_NAME, row["VCL Alias"]))
                    assert row["VCL Alias"].lower() == "applicationnameasint32", error_text

                vehicle_ctrl_sw_param = row["Application Default"]

        if not vehicle_ctrl_sw_param or pd.isna(vehicle_ctrl_sw_param):
            # Empty value
            self.source_ctrl_sw_pn = None
            return

        # Validate that SW P/N value conforms to expected format.
        prompt_str = ("Found multiple possible ctrl SW P/Ns stored in CDF '%s': '%s'.\n"
                                                        "Press Enter to continue."
                        % (self.export_filename, vehicle_ctrl_sw_param))
        valid_sw_pn, stole_focus = find_in_string(CDF_SW_PN_REGEX, vehicle_ctrl_sw_param, prompt_str, allow_none=True)
        if stole_focus:
            ParentDB.get_GUI_Driver().lose_focus()

        if valid_sw_pn is None:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("Expected '%s' variable to contain SW P/N in ########.## format."
                                                    "\nFound '%s' in %s instead."
                                        % (VSN_CDF_VAR_NAME, vehicle_ctrl_sw_param,
                        self.export_filename) + Style.RESET_ALL)
            self.source_ctrl_sw_pn = None
        elif valid_sw_pn != vehicle_ctrl_sw_param:
            self.ParentDB.get_GUI_Driver().lose_focus()
            print(Fore.RED + Style.BRIGHT)
            input("'%s' value '%s' (in %s) appears to contain SW P/N with right "
                                        "format but may contain additional content."
                    % (VSN_CDF_VAR_NAME, vehicle_ctrl_sw_param, self.export_filename)
                                                        + Style.RESET_ALL)
            self.source_ctrl_sw_pn = None
        else:
            # Replace period with "G" in SW P/N string and return
            self.source_ctrl_sw_pn = "G".join(vehicle_ctrl_sw_param.split(".")) # string

    def extract_cdf_cprj_pn(self, given_export_dir=None):
        """Takes in CDF export (.xlsx format), locates the SW P/N associated with
        the .cprj file that was loaded in CIT when the CDF was converted.
        Returns cprj SW P/N as string.
        """
        assert self.has_export(), "Tried to extract CDF cprj P/N from export, but export doesn't exist.\n\t%s" % self
        worksheet_names = pd.ExcelFile(self.export_path).sheet_names
        # https://stackoverflow.com/a/17977609

        found = False
        for sheet_name in worksheet_names:
            prompt_str = ("Found multiple possible cprj SW P/Ns stored in CDF (%s) worksheet name '%s'.\n"
                                                            "Press Enter to continue."
                                % (self.export_filename, sheet_name))
            # Find worksheet w/ P/N in the name
            sw_pn, stole_focus = find_in_string(SW_PN_REGEX, sheet_name, prompt_str, allow_none=True)
            if stole_focus:
                self.ParentDB.get_GUI_Driver().lose_focus()

            if sw_pn is None:
                continue
            elif found:
                # If a second tab name containing P/N is found, that violates an assumption that only one tab has it.
                raise Exception("Found more than one tab in file '%s' w/ a name "
                                        "including cprj SW P/N." % self.export_path)
            else:
                found = True
                cprj_pn = sw_pn
                continue # Look at rest of tabs to see if another exists w/ P/N, even though that isn't expected.

        assert found, "No cprj SW P/N found in any tab in file '%s'." % self.export_path

        return cprj_pn

    def __str__(self):
        return self.cdf_filename

    def __repr__(self):
        return "CDF object '%s'" % self.__str__()


class CloneDataFileDB(object):
    def __init__(self, cdf_import_dir, conv_export_dir):
        assert os.path.exists(cdf_import_dir), "Can't find source_dir '%s'" % cdf_import_dir
        assert os.path.exists(conv_export_dir), "Can't find export_dir '%s'" % conv_export_dir

        self.file_type = "CDF"
        self.source_dir = cdf_import_dir
        self.export_dir = conv_export_dir

        self.CDF_list = None # To be populated by _build_cdf_list()

        self.cprj_rev_dict = dict() # value: CIT cprj rev; val: list of CDF objects.
        # Populated when mismatch detected b/w SW rev and CIT cprj rev, resulting in incorrect VCL-alias mappings

        self._build_cdf_list()

        self.ActiveGUI_Driver = None    # To be set by convert_all()

    def get_GUI_Driver(self):
        return self.ActiveGUI_Driver

    def _build_cdf_list(self):
        self.CDF_list = []
        for filename in sorted(os.listdir(self.source_dir)):
            if filename.upper().endswith(self.file_type):
                self.CDF_list.append( CloneDataFile(os.path.join(self.source_dir, filename), self) )

    def convert_all(self, ActiveGUI_Driver, check_SNs=False):
        if self.cprj_rev_dict:
            # Tail call will land here to process previously-encounterd CDFs
            # that need a different cprf rev to process correctly.
            cprj_rev, CDF_obj_list = self.cprj_rev_dict.popitem()
            self.ActiveGUI_Driver.lose_focus()
            print(Fore.GREEN + Style.BRIGHT)
            input("Load cprj w/ rev %s into CIT then press Enter to continue." % cprj_rev + Style.RESET_ALL)
        else:
            # First call will land here.
            CDF_obj_list = self.CDF_list

        self.ActiveGUI_Driver = ActiveGUI_Driver
        try:
            self.ActiveGUI_Driver.select_program(self.file_type)
        except UserCancel:
            return

        for CDF_obj in tqdm(CDF_obj_list, colour="#6700ff"):
            # Check for existing export
            if not CDF_obj.is_valid_cdf():
                print(Fore.WHITE, Style.DIM, end="")
                tqdm.write("%s: Skipping empty file" % CDF_obj + Style.RESET_ALL)
                continue
            elif CDF_obj.has_export(self.export_dir):
                # Skip if already processed this file.
                # Will delete and reprocess to remove invalid mappings from exports.
                if not CDF_obj.check_cprj_rev_match(self.export_dir):
                    CDF_obj.remove_export()
                    print(Fore.RED, end="")
                    tqdm.write(" %s: Already processed but invalid alias mapping - export deleted" % CDF_obj + Style.RESET_ALL)
                    # Fall through to conversion below, where it will be converted
                    # again (possibly w/ the right mapping), and if the mapping is wrong
                    # again, the file will get stored along with the needed mapping in cprj_rev_dict
                else:
                    # revs match, so skip this file.
                    print(Fore.WHITE, Style.DIM, end="")
                    tqdm.write("%s: Already processed; valid alias mapping confirmed" % CDF_obj + Style.RESET_ALL)
                    continue

            try:
                success = CDF_obj.convert(self.export_dir, check_sn=check_SNs)
            except Exception as exception_text:
                self.ActiveGUI_Driver.lose_focus()
                if CDF_obj.has_export():
                    # Remove export that may not have been validated
                    CDF_obj.remove_export()
                print(Fore.CYAN + Style.BRIGHT)
                print(" %s: Encountered exception during processing" % CDF_obj + Style.RESET_ALL)
                print(exception_text)
                print(Fore.GREEN + Style.BRIGHT)
                print("Press Enter to continue with other files, 'e' to exit "
                                "file-conversion loop, or 'q' to quit program.")
                answer = input("> " + Style.RESET_ALL)
                if answer.lower() == "":
                    continue
                elif answer.lower() == "e":
                    break
                else:
                    # Accept anything other than a blank input or 'e' as a quit command.
                    quit()
            else:
                if success:
                    print(Fore.GREEN, end="")
                    tqdm.write(" %s: Processed; valid alias mapping confirmed" % CDF_obj + Style.RESET_ALL)
                else:
                    # success=False (but no exception thrown) means check_cprj_rev_match() failed
                    # Add to dict to be processed with different rev in tail call below.
                    print(Fore.RED, end="")
                    tqdm.write(" %s: Processed but with invalid alias mapping - export deleted" % CDF_obj + Style.RESET_ALL)
                    if CDF_obj.get_ctrl_sw_rev() not in self.cprj_rev_dict:
                        self.cprj_rev_dict[CDF_obj.get_ctrl_sw_rev()] = []
                    self.cprj_rev_dict[CDF_obj.get_ctrl_sw_rev()].append(CDF_obj)
                    # tqdm.write("File %s needs rev-%s cprj file for conversion."
                    #                      % (CDF_obj, CDF_obj.get_ctrl_sw_rev()))

        if self.cprj_rev_dict:
            # Process files needing different cprj file(s):
            self.convert_all(self.ActiveGUI_Driver, check_SNs=check_SNs)


def convert_all_cpfs(source_dir, dest_dir, check_SNs=False):
    file_type = "cpf"

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
        if file_type == "cpf" and (os.path.exists(os.path.join(dest_dir,
                            os.path.splitext(filename)[0] + CPF_COMBINED_EXPORT_SUFFIX))):
            # Skip if already processed this file.
            tqdm.write("Already processed %s" % os.path.basename(filename)) # DEBUG
            continue

        filepath = os.path.join(source_dir, filename)
        if (os.path.isfile(filepath) and
                    os.path.splitext(filename)[-1].lower() == ".%s" % file_type):
            try:
                success = convert_file(filepath, dest_dir, check_sn=check_SNs,
                                                              gui_in_focus=True)
            except Exception as exception_text:
                print(Fore.CYAN + Style.BRIGHT)
                print("\nEncountered exception processing %s" % filename + Style.RESET_ALL)
                print(exception_text)
                print(Fore.GREEN + Style.BRIGHT)
                print("Press Enter to continue with other files, 'e' to exit "
                                "file-conversion loop, or 'q' to quit program.")
                answer = input("> " + Style.RESET_ALL)
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

    print("\nConverting CPF exports from .tsv format (named .XLS) to .xslx (in dir "
                                                    "\n\t\"%s\")..." % dir_path)
    try:
        fixcpf.convert_all_param_exports(dir_path, check_xls=False)
        print("...done")
    except PermissionError:
        # Gets a PermissionError if running on PowerShell most of the time.
        print(Fore.GREEN + Style.BRIGHT)
        input("\nEncountered permission error in removing CPF tsv files.\n"
                        "Press Enter to continue to next part of program.")
        print(Style.RESET_ALL)


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
    parser.add_argument("-s", "--slow", help="Specify factor by which to extend "
                            "pauses b/w GUI commands. >1 extends pauses while "
                                    "<1 speeds them up.", type=float, default=1)
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
        update_from_remote_dirs()
        print(Fore.GREEN + Style.BRIGHT)
        print("Press Enter to proceed to file processing or 'q' to quit program.")
        answer = input("> " + Style.RESET_ALL)
        if answer == "":
            pass
        else:
            # Accept anything other than a blank input as a quit command.
            quit()

    # Convert exports
    if os.name == "nt":
        GUI_PAUSE_MULT = args.slow # Extend or reduce pauses between GUI commands
        gui.PAUSE = 0.5 * GUI_PAUSE_MULT # 500 ms pause after each command.

        try:
            convert_all_cpfs(import_dir, export_dir, check_SNs=check_vehicle_sns)

            GUI_DriverInstance = GUI_Driver()
            CDF_Database = CloneDataFileDB(import_dir, export_dir)
            CDF_Database.convert_all(GUI_DriverInstance, check_SNs=check_vehicle_sns)
            print(Fore.MAGENTA + Style.BRIGHT + "\nGUI interaction done\n" + Style.RESET_ALL)

        except gui.FailSafeException:
            GUI_DriverInstance.lose_focus()
            print(Fore.MAGENTA + Style.BRIGHT + "\n\nUser canceled GUI interaction.")
            print(Style.RESET_ALL)
            time.sleep(3 * GUI_PAUSE_MULT)
            # If user terminates GUI interraction, continue running below.
            pass
    else:
        print(Fore.MAGENTA + Style.BRIGHT + "Skipping GUI interaction "
                                                   "(requires Windows system).")
        print(Style.RESET_ALL)


    if auto_run:
        sync_to_remote_dirs()


def sync_to_remote_dirs():
    # Sync to shared folder
    print(Fore.GREEN + Style.BRIGHT)
    print("\nSync local Controller-export dir to shared folder? Enter to "
                            "proceed, 's' to skip, or 'q' to quit program.")
    answer = input("> " + Style.RESET_ALL)
    if answer == "":
        print("Syncing ctrl exports to shared folder...")
        sync_remote(DIR_EXPORT, os.path.join(DIR_REMOTE_SHARE_CTRL, "Converted"),
                                                purge=True, multilevel=False)
        print("...done")
    elif answer.lower() == "s":
        print("Skipping sync from local ctrl-export dir to shared folder.")
    else:
        # Accept anything other than a blank input as a quit command.
        quit()

    # Sync to second remote (Azure blob)
    # Controller exports  |  local dir --> Azure blob storage
    sync_to_azure(DIR_EXPORT, AZ_BLOB_ADDR_CTRL, "local Controller-export dir")

    # BDX files  |  shared folder --> Azure blob storage
    sync_to_azure(DIR_REMOTE_SHARE_BATT, AZ_BLOB_ADDR_BATT, "shared-folder Battery-export dir")

    # MES (manufacturing) battery-scan data  |  shared folder --> Azure blob storage
    sync_to_azure(DIR_REMOTE_SHARE_MFG, AZ_BLOB_ADDR_MFG, "shared-folder MES batt-scan export dir")


def sync_to_azure(src_dir, dest_dir, src_desc):
    print(Fore.GREEN + Style.BRIGHT)
    print("\nSync %s to Azure blob? Enter to "
                    "proceed, 's' to skip, or 'q' to quit program." % src_desc)
    answer = input("> " + Style.RESET_ALL)
    if answer == "":
        print("\nRunning AzCopy sync job from %s..." % src_desc)
        print(Fore.BLUE + Style.BRIGHT)
        returncode = subprocess.call(["azcopy", "sync",
                                            "--delete-destination", "true",
                                                        "--exclude-path=tmp",
                                            os.path.join(src_dir, ""), dest_dir])
        # https://learn.microsoft.com/en-us/azure/storage/common/storage-ref-azcopy-sync
        # https://stackoverflow.com/questions/68894328/azcopy-copy-exclude-a-folder-and-the-files-inside-it
        # https://stackoverflow.com/a/15010678
        print(Style.RESET_ALL + "...done")
    elif answer.lower() == "s":
        print("Skipping sync from %s to Azure blob." % src_desc)
    else:
        quit()
